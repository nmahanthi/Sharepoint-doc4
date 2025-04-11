import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Pivot, PivotItem, PrimaryButton } from '@fluentui/react';
import styles from './DmsDocumentForm.module.scss';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { IFieldInfo } from '@pnp/sp/fields';
import { IFieldLinkInfo } from '@pnp/sp/content-types';
import { IAdditionalFieldOption, IDmsDynamicFieldProps } from './DmsDynamicFieldBase';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { Logger } from '@pnp/logging';
import { getSPCacheNever } from '../../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import * as strings from 'DmsDocumentFormFormCustomizerStrings';
import { IListItemFormUpdateValue } from '@pnp/sp/lists';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import DmsDynamicField from './DmsDynamicField';
import { FieldNames, DmsDocumentFormConstants, APPLICABLESTATUS, DocumentStatusFieldKey } from '../../../constants';
import { AadHttpClient } from '@microsoft/sp-http';

export interface IDmsDynamicFormProps<T extends { ID: number | undefined;[key: string]: unknown }> {
  displayMode: FormDisplayMode;
  fields: IFieldInfo[];
  fieldLinks: IFieldLinkInfo[];
  initialItem?: T;
  initialItemData?: never;
  onSave: () => void;
  onClose: () => void;
  onError?: (error: string) => void;
  onChange?: (item: T) => void;
  sectionsToFieldsMap: { [sectionName: string]: string[] };
  additionalFieldOptions?: { [fieldName: string]: IAdditionalFieldOption };
  context: FormCustomizerContext;
  isDocumentController: boolean | undefined;
  isProjectAdmin: boolean | undefined;
  projectEndpoint?: string;
  dmsClient: AadHttpClient;
  updateApplicableDocumentEndpoint?: string;
}
export interface IDmsDynamicFormState<T extends { ID: number | undefined }> {
  item: T;
  fieldValidationErrors: { [fieldName: string]: string };
  fieldSubmitErrors: { [fieldName: string]: string };
  fieldLoadErrors: { [fieldName: string]: Error | undefined };
  fieldStates: IDmsDynamicFieldProps[];
  selectedSection: string;
  formHeight?: number;
  isBusy?: boolean;
  currentPresaveOperation?: {
    label: string;
    status: 'running' | 'success' | 'error';
    index: number;
    count: number;
  }
  isStatusApplicableAndDocController: boolean;
}

const LOG_SOURCE: string = 'DmsDynamicForm';
const modeClasses = {
  [FormDisplayMode.Display]: styles.viewMode,
  [FormDisplayMode.Edit]: styles.editMode,
  [FormDisplayMode.New]: styles.newMode
};

export default class DmsDynamicForm<T extends { ID: number | undefined;[key: string]: unknown }> extends React.Component<IDmsDynamicFormProps<T>, IDmsDynamicFormState<T>> {

  private _fieldRefs: { [fieldName: string]: React.RefObject<DmsDynamicField> };
  private _formContainerRefs: { [sectionName: string]: React.RefObject<HTMLDivElement> };
  private _sp: SPFI;
  private _fieldInitBatch: SPFI;
  private _previousProjectRevision: unknown
  private _executeFieldInitBatch: () => Promise<void>;
  /**
   *
   */
  /**
   *
   */
  constructor(props: IDmsDynamicFormProps<T>) {
    super(props);
    // this._sp = getSP();
    this._sp = getSPCacheNever();
    const [fieldInitBatch, executeFieldInitBatch] = this._sp.batched();
    this._fieldInitBatch = fieldInitBatch;
    this._executeFieldInitBatch = executeFieldInitBatch;

    const fieldNames = Object.keys(props.sectionsToFieldsMap).reduce((acc, sectionName) => acc.concat(props.sectionsToFieldsMap[sectionName]), [] as string[]);
    this._fieldRefs = fieldNames.reduce((acc, fieldName) => {
      acc[fieldName] = React.createRef<DmsDynamicField>();
      return acc;
    }, {} as { [fieldName: string]: React.RefObject<DmsDynamicField> });

    this._formContainerRefs = Object.keys(props.sectionsToFieldsMap).reduce((acc, sectionName) => {
      acc[sectionName] = React.createRef<HTMLDivElement>();
      return acc;
    }, {} as { [sectionName: string]: React.RefObject<HTMLDivElement> });
    const item: T = (props.initialItem || {}) as T;
    this.state = {
      item,
      fieldStates: [],
      selectedSection: Object.keys(props.sectionsToFieldsMap)[0],
      fieldValidationErrors: {},
      fieldSubmitErrors: {},
      fieldLoadErrors: {},
      isStatusApplicableAndDocController: (props.isDocumentController || props.isProjectAdmin) && item && item[DocumentStatusFieldKey] === APPLICABLESTATUS ? true : false
    };
    this._previousProjectRevision = this.state.item.ProjectRevision;
  }
  private async _validateForm(): Promise<{ [fieldName: string]: string; }> {
    const [batch, execute] = this._sp.batched();
    const result = Promise.all(
      [
        ...Object.keys(this._fieldRefs)
          .filter(fieldName => this._fieldRefs[fieldName].current && !this._fieldRefs[fieldName].current?.props.readonly && !this._fieldRefs[fieldName].current?.props.fieldInfo.ReadOnlyField)
          .map(fieldName => (this._fieldRefs[fieldName].current?.validate(batch).then(validationError => ({ fieldName, validationError }))))
      ]
    ).then(results => {
      const filteredValidationErrors = results.reduce((acc, result) => {
        if (result?.validationError) {
          acc[result.fieldName] = result.validationError as string;
        }
        return acc;
      }, {} as { [fieldName: string]: string });
      this.setState({ fieldValidationErrors: filteredValidationErrors });
      return filteredValidationErrors;
    });
    await execute();
    return result;
  }
  private async _onSave(): Promise<void> {
    const { onError, context, displayMode, onSave, dmsClient, updateApplicableDocumentEndpoint } = this.props;
    const { isStatusApplicableAndDocController } = this.state;
    try {
      this.setState({ isBusy: true });
      const fieldValidationErrors = await this._validateForm();
      if (Object.keys(fieldValidationErrors).length > 0) {
        this.setState({ isBusy: false });
        return;
      }
      const valuesForUpdate = this._getValuesForUpdate(displayMode === FormDisplayMode.New);
      if (displayMode === FormDisplayMode.Edit) {
        const continueSave = await this._runPresaveOperations(valuesForUpdate);
        if (continueSave) {
          const saveSuccess = await this._saveItem(valuesForUpdate);
          // const saveSuccess = isStatusApplicableAndDocController ? await this._saveItemToProjectRequest(dmsClient,updateApplicableDocumentEndpoint,valuesForUpdate): await this._saveItem(valuesForUpdate);
          if (saveSuccess && onSave) onSave();
        }
      }
      else if (displayMode === FormDisplayMode.Display && isStatusApplicableAndDocController) {
        const continueSave = await this._runPresaveOperations(valuesForUpdate);
        if (continueSave) {
          const saveSuccess = await this._saveItemToProjectRequest(dmsClient, updateApplicableDocumentEndpoint, valuesForUpdate);
          if (saveSuccess && onSave) onSave();
        }
      } else {
        const saveSuccess = await this._saveItem(valuesForUpdate);
        if (saveSuccess) {
          const success = await this._runPresaveOperations(valuesForUpdate);
          if (success && onSave) onSave();
          if (!success) {
            try {
              const folderName = `${this._fieldRefs[FieldNames.ProjectReference].current?.state.value}-${this._fieldRefs[FieldNames.Revision].current?.state.value}`;
              await this._sp.web.lists.getById(context.list.guid.toString()).rootFolder.folders.getByUrl(folderName)
                .deleteWithParams({ BypassSharedLock: true, DeleteIfEmpty: false });
            } catch (error) {
              console.error(error);
              Logger.error(new Error(`${LOG_SOURCE}: Error deleting folder item`));
              if (onError) onError(strings.CleanError);
            }
          }
        }
      }
    } catch (error) {
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}: Error saving item`));
      if (onError) onError(strings.SaveError);
      this.setState({ isBusy: false });
    }
  }

  private async _saveItemToProjectRequest(dmsClient: AadHttpClient, updateApplicableDocumentEndpoint: string | undefined, valuesForUpdate: IListItemFormUpdateValue[]): Promise<boolean> {
    const updatedValues: any[] = valuesForUpdate.map((item: IListItemFormUpdateValue) => {
      if (item.FieldName && DmsDocumentFormConstants.enableInApplicableViewForm.includes(item.FieldName)) {
        if (DmsDocumentFormConstants.enableInApplicableViewFormDateFields.includes(item.FieldName)) {
          const datePattern = /^(0[1-9]|1[0-2])\/(0[1-9]|[12][0-9]|3[01])\/\d{4}$/;
          if (item.FieldValue && item.FieldValue !== "" && datePattern.test(item.FieldValue)) {
            const parts = item.FieldValue.split("/");
            const formattedDate = `${parts[2]}-${parts[0]}-${parts[1]}T00:00:00Z`;
            const date: Date = new Date(formattedDate);

            // Get the month, day, and year
            const month: number = date.getMonth() + 1; // Months are zero-indexed
            const day: number = date.getDate();
            const year: number = date.getFullYear();

            // Format the date as m/d/yyyy without leading zeros
            const americanFormatDate: string = `${month}/${day}/${year}`;

            console.log(americanFormatDate); // Output: 12/1/2024
            item.FieldValue = americanFormatDate;
            // item.FieldValue = formattedDate;
          }
          else item.FieldValue = undefined;
        }
        else if (DmsDocumentFormConstants.enableInApplicableViewFormBooleanFields.includes(item.FieldName)) {
          if (item.FieldValue === "" || (item.FieldValue && item.FieldValue.toLocaleLowerCase() === 'false'))
            item.FieldValue = "0";
          else if (item.FieldValue && item.FieldValue.toLocaleLowerCase() === 'true')
            item.FieldValue = '1';
        }
        return item.FieldValue ? item : null;
      }
    }).filter(i => i);

    const request: any = {
      siteUrl: this.props.context.pageContext.web.serverRelativeUrl,
      properties: updatedValues,
      docSetId: this.state.item.ID
    };
    return updateApplicableDocumentEndpoint ? await dmsClient.post(updateApplicableDocumentEndpoint, AadHttpClient.configurations.v1, {
      body: JSON.stringify(request)
    }).then(() => {
      return true;
    }).catch(err => {
      this.setState({ isBusy: false });
      console.error(err);
      Logger.error(new Error(`${LOG_SOURCE}: ${err}`));
      return false;
    }) : false;
  }


  private async _runPresaveOperations(valuesForUpdate: IListItemFormUpdateValue[]): Promise<boolean> {
    const { additionalFieldOptions, context } = this.props;
    const presaveOperations = Object.entries(additionalFieldOptions || {}).filter(([_, afo]) => afo.preSaveOperation).map(([field, afo]) => ({ operation: afo.preSaveOperation, label: afo.preSaveOperationLabel, field }));
    let continueSave = true;
    for (const [index, { operation, label, field }] of presaveOperations.entries()) {
      const promise = operation ? operation(this._sp, context.spHttpClient, this._fieldRefs[field], valuesForUpdate) : undefined;
      if (!promise) continue;
      this.setState({ currentPresaveOperation: { label: label || '', status: 'running', index: index + 1, count: presaveOperations.length } });
      try {
        await promise;
        this.setState({ currentPresaveOperation: { label: label || '', status: 'success', index: index + 1, count: presaveOperations.length } });
      } catch (error) {
        this.setState({
          currentPresaveOperation: { label: label || '', status: 'error', index: index + 1, count: presaveOperations.length },
          isBusy: false
        });
        console.error(error);
        continueSave = false;
      }
    }
    return continueSave;
  }
  private _getValuesForUpdate(addMissingDefaults: boolean): IListItemFormUpdateValue[] {
    const { fields, fieldLinks } = this.props;
    const result = Object.values(this._fieldRefs).filter(fr =>
      (fr.current && !fr.current?.props.fieldInfo.ReadOnlyField && !fr.current?.props.readonly)
    ).reduce((acc, fieldRef) => {
      if (fieldRef.current) {
        const fieldName = fieldRef.current.props.fieldLink?.Name;
        acc.push({
          FieldName: fieldName,
          FieldValue: this._getValueForUpdate(fieldRef.current.props.fieldInfo, fieldRef.current.state.value)
        }
        );
      }
      return acc;
    }, [] as IListItemFormUpdateValue[]);
    if (addMissingDefaults) {
      fields.filter(f => (
        !result.find(r => r.FieldName === f.InternalName) &&
        !f.ReadOnlyField &&
        f.TypeAsString !== 'Computed' &&
        f.DefaultValue &&
        fieldLinks.find(fl => fl.Name === f.InternalName)
      )).forEach(field => {
        result.push({
          FieldName: field.InternalName,
          FieldValue: this._getValueForUpdate(field, field.DefaultValue)
        });
      });
    }
    return result;
  }
  private async _saveItem(valuesForUpdate: IListItemFormUpdateValue[]): Promise<boolean> {
    const { context, displayMode, onError, initialItem } = this.props;
    const itemId = initialItem?.ID;

    let addOrUpdateResult: IListItemFormUpdateValue[] | undefined = undefined;
    const folderName = `${this._fieldRefs[FieldNames.ProjectReference].current?.state.value}-${this._fieldRefs[FieldNames.Revision].current?.state.value}`;
    if (displayMode === FormDisplayMode.New) {
      if (!folderName) throw new Error('Folder name is required');
      const fldr = await this._sp.web.lists.getById(context.list.guid.toString()).rootFolder.folders.addUsingPath(folderName);
      try {
        const folderItem = await this._sp.web.getFolderByServerRelativePath(fldr.ServerRelativeUrl).getItem();
        addOrUpdateResult = await folderItem.validateUpdateListItem(valuesForUpdate.filter(v => v.FieldName !== 'FileLeafRef').concat(
          { FieldName: 'ContentTypeId', FieldValue: context.contentType.id }
        ));
      } catch (error) {
        console.error(error);
        Logger.error(new Error(`${LOG_SOURCE}: Error updating folder item`));
        if (onError) onError(strings.AddError);
        try {
          await this._sp.web.lists.getById(context.list.guid.toString()).rootFolder.folders.getByUrl(folderName)
            .deleteWithParams({ BypassSharedLock: true, DeleteIfEmpty: false });
        } catch (error) {
          console.error(error);
          Logger.error(new Error(`${LOG_SOURCE}: Error deleting folder item`));
          if (onError) onError(strings.CleanError);
        }
      }
      this.setState({ isBusy: false });
    }
    if (displayMode === FormDisplayMode.Edit) {
      if (!itemId) throw new Error('Item ID is required');
      const newProjectRevision = valuesForUpdate.find(v => v.FieldName === 'ProjectRevision')?.FieldValue;
      if (this._previousProjectRevision === newProjectRevision) {
        addOrUpdateResult = await this._sp.web.lists.getById(context.list.guid.toString()).items.getById(itemId).validateUpdateListItem(valuesForUpdate);
      } else {
        addOrUpdateResult = await this._sp.web.lists.getById(context.list.guid.toString())
          .items.getById(itemId).validateUpdateListItem(valuesForUpdate.filter(v => v.FieldName !== 'FileLeafRef').concat(
            { FieldName: 'FileLeafRef', FieldValue: folderName }
          ));

      }
      this.setState({ isBusy: false });
    }
    if (addOrUpdateResult) {
      const erroredFields = addOrUpdateResult.filter(r => r.HasException).map(r => r.FieldName).join(', ');
      if (erroredFields) {
        Logger.error(new Error(`${LOG_SOURCE}: Error saving item, errored fields: ${erroredFields}`));
        if (onError) onError(`${strings.PartialSaveError}`);
        this.setState({
          fieldSubmitErrors: addOrUpdateResult.filter(r => r.HasException).reduce((acc, r) => {
            if (r.FieldName && r.ErrorMessage) acc[r.FieldName] = r.ErrorMessage;
            return acc;
          }, {} as { [fieldName: string]: string })
        });
        if (displayMode === FormDisplayMode.New) {
          try {
            await this._sp.web.lists.getById(context.list.guid.toString()).rootFolder.folders.getByUrl(folderName)
              .deleteWithParams({ BypassSharedLock: true, DeleteIfEmpty: false });
          } catch (error) {
            console.error(error);
            Logger.error(new Error(`${LOG_SOURCE}: Error deleting folder item`));
            if (onError) onError(strings.CleanError);
          }
        }
      } else {
        return true;
      }
    } else {
      Logger.error(new Error(`${LOG_SOURCE}: Error saving item`));
      if (onError) onError(strings.SaveError);
      this.setState({ isBusy: false });
      return false;
    }
    return false;
  }
  private _getValueForUpdate(fieldInfo: IFieldInfo, value: unknown): string {
    const taxTypedValues = value as { id: string; labels: { name: string }[] }[];
    const dateTypedValue = value ? new Date(value as string) : null;
    const userTypedValue = value as ISiteUserInfo[];
    const multiChoiceTypedValue = value as string[];
    const urlTypedValue = value as { Description: string; Url: string };
    switch (fieldInfo.TypeAsString) {
      case 'User':
      case 'UserMulti':
        return JSON.stringify(userTypedValue ? userTypedValue.map(v => ({ Key: v.LoginName })) : []);
      case 'TaxonomyFieldType':
        if (!taxTypedValues || taxTypedValues.length === 0) return '';
        else
          return `${taxTypedValues[0].labels[0].name}|${taxTypedValues[0].id}`;
      case 'TaxonomyFieldTypeMulti':
        if (!taxTypedValues || taxTypedValues.length === 0) return '';
        else
          return taxTypedValues.map(v => (`${v.labels[0].name}|${v.id}`)).join(';');
      case 'DateTime':
        return dateTypedValue ?
          `${`${dateTypedValue.getMonth() + 1}`.padStart(2, '0')}/${`${dateTypedValue.getDate()}`.padStart(2, '0')}/${dateTypedValue.getFullYear()}` :
          '';
      case 'URL':
        return urlTypedValue ? `${urlTypedValue.Url}, ${urlTypedValue.Description || urlTypedValue.Url}` : '';
      case 'MultiChoice':
        if (!multiChoiceTypedValue || multiChoiceTypedValue.length === 0) return '';
        else {
          if (multiChoiceTypedValue.join) {
            return multiChoiceTypedValue.join('#;');
          }
          return `${multiChoiceTypedValue}`;
        }
      default:
        return value ? (value as string).charAt ? value as string : JSON.stringify(value) : '';
    }
  }
  private _onClose(): void {
    const { onClose } = this.props;
    onClose();
  }

  public componentDidMount(): void {
    const { fieldLinks, fields, onError, initialItemData, displayMode } = this.props;
    const { item } = this.state;
    const fieldStates = fieldLinks.filter(f => {
      const fieldInfo = fields.find(fi => fi.InternalName === f.Name);
      return !f.Hidden && fieldInfo && !fieldInfo.Hidden
    }).map(fieldLink => {
      let name = fieldLink.Name;
      const fieldInfo = fields.find(f => f.InternalName === fieldLink.Name);
      if (fieldInfo?.TypeAsString === 'User' || fieldInfo?.TypeAsString === 'UserMulti') {
        name = fieldLink.Name + 'Id';
      }
      const additionalFieldOptions = this.props.additionalFieldOptions ? this.props.additionalFieldOptions[name] : undefined;
      if (!name || !fieldInfo) {
        Logger.error(new Error(`${LOG_SOURCE}: Field '${name}' not found in fields`));
        if (onError) onError(`Field '${name}' not found in fields`);
        return;
      }

      return {
        value: displayMode === FormDisplayMode.New ? fieldInfo.DefaultValue : item ? item[name] : undefined,
        data: initialItemData ? initialItemData[fieldLink.Name] : undefined,
        fieldInfo,
        fieldLink,
        name,
        additionalFieldOptions
      } as IDmsDynamicFieldProps;
    }).filter(f => f !== undefined) as Partial<IDmsDynamicFieldProps>[];
    ['Author', 'Created', 'Editor', 'Modified'].forEach(fieldName => {
      fieldStates.push({
        value: item[fieldName],
        data: initialItemData ? initialItemData[fieldName] : undefined,
        fieldInfo: fields.find(f => f.InternalName === fieldName) as IFieldInfo,
        name: fieldName,
        readonly: true,
        additionalFieldOptions: {
          [FormDisplayMode.New]: { hidden: Promise.resolve(true) },
          [FormDisplayMode.Edit]: { disabled: Promise.resolve(true) }
        }
      });
    });
    this.setState({
      fieldStates: fieldStates as IDmsDynamicFieldProps[],
    }, () => {
      this._executeFieldInitBatch().catch(error => {
        //TODO: Better error handling
        Logger.error(new Error(`${LOG_SOURCE}: Error initializing fields`));
        if (onError) onError(`${strings.FieldLoadError}`.replace('{0}', error.message));
        console.error(error);
      })
    });
  }
  public componentDidUpdate(_: IDmsDynamicFormProps<T>, prevState: IDmsDynamicFormState<T>): void {
    const { fieldLoadErrors, selectedSection, formHeight } = this.state;
    const { fieldLoadErrors: prevFieldLoadErrors } = prevState;
    const { onError } = this.props;
    const currentHeight = this._formContainerRefs[selectedSection].current?.clientHeight || 0;
    if (currentHeight > (formHeight || 0)) {
      const formHeight = Object.values(this._formContainerRefs)
        .reduce((acc, ref) => Math.max(acc, ref.current?.clientHeight || 0), 0);
      this.setState({ formHeight });
    }
    if (fieldLoadErrors !== prevFieldLoadErrors && onError) {
      onError(`${strings.FieldLoadError}`.replace('{0}', Object.keys(fieldLoadErrors).join(', ')));
    }
  }

  public render(): React.ReactElement<{}> {
    const { sectionsToFieldsMap, context, displayMode } = this.props;
    const { currentPresaveOperation, fieldStates, selectedSection, formHeight,
      fieldValidationErrors, isBusy, fieldSubmitErrors,
      fieldLoadErrors, isStatusApplicableAndDocController } = this.state;
    const sectionsValidationErrors = fieldValidationErrors ? Object.keys(sectionsToFieldsMap).reduce((acc, sectionName) => {
      acc[sectionName] = sectionsToFieldsMap[sectionName].reduce((acc, fieldName) => acc + (fieldValidationErrors[fieldName] ? 1 : 0), 0);
      return acc;
    }, {} as { [sectionName: string]: number }) : undefined;
    if (fieldSubmitErrors && sectionsValidationErrors) {
      Object.keys(fieldSubmitErrors).forEach((key) => {
        const fieldSection = Object.keys(sectionsToFieldsMap).find(section => sectionsToFieldsMap[section].includes(key));
        if (fieldSection) {
          if (sectionsValidationErrors && sectionsValidationErrors[fieldSection]) {
            sectionsValidationErrors[fieldSection] += 1;
          } else {
            sectionsValidationErrors[fieldSection] = 1;
          }
        }
      });
    }
    return <>
      <div className={`${styles.formWrapper} ${modeClasses[displayMode]}`}>
        <Pivot
          headersOnly={true}
          onLinkClick={(item) => {
            this.setState({
              selectedSection: item?.props.itemKey || Object.keys(sectionsToFieldsMap)[0]
            })
          }}>
          {Object.keys(sectionsToFieldsMap).filter(sectionName =>
            sectionsToFieldsMap[sectionName].map(fieldName =>
              this._fieldRefs[fieldName].current &&
              !this._fieldRefs[fieldName].current?.state.isAdditionalOptionHidden).some(Boolean)
          ).map((sectionName, index) =>
            <PivotItem
              key={index}
              itemKey={sectionName}
              headerText={sectionName}
              itemCount={sectionsValidationErrors && sectionsValidationErrors[sectionName] > 0 ? sectionsValidationErrors[sectionName] : undefined}
              itemIcon={sectionsValidationErrors && sectionsValidationErrors[sectionName] ? 'warning' : undefined}
              onRenderItemLink={
                sectionsValidationErrors && sectionsValidationErrors[sectionName] > 0 ?
                  (props, defaultRender) => {
                    return defaultRender ? <div className={styles.invalidSection}>{defaultRender(props)}</div> : <></>;
                  } :
                  undefined} />
          )}
        </Pivot>
        <div className={styles.formOuterContainer} style={{ height: formHeight ? `${formHeight}px` : 'auto' }}>
          <div className={styles.formMidContainer}>
            {Object.keys(sectionsToFieldsMap).map((sectionName, index) =>
              <div className={`${styles.formInnerContainer} ${selectedSection === sectionName ? '' : styles.hidden}`}
                ref={this._formContainerRefs[sectionName]} key={`${sectionName}-${index}`} >
                <div className={styles.column}>
                  <div className={styles.row}>
                    {fieldStates.filter(f => sectionsToFieldsMap[sectionName].includes(f.fieldInfo.InternalName)).map((fieldState, index) => {
                      const fieldRef = this._fieldRefs[fieldState.fieldInfo.InternalName];
                      return <div className={styles.fieldWrapper} key={`${sectionName}-field-${index}`}>
                        {fieldState.additionalFieldOptions && fieldState.additionalFieldOptions.renderer ?
                          fieldState.additionalFieldOptions.renderer({
                            ...fieldState,
                            context,
                            disabled: isStatusApplicableAndDocController && displayMode === FormDisplayMode.Display ? this._fieldDisablityCheck(fieldState.fieldInfo.InternalName, DmsDocumentFormConstants.enableInApplicableViewForm) : (displayMode === FormDisplayMode.Display) || isBusy || !!fieldLoadErrors[fieldState.fieldInfo.InternalName],
                            ref: fieldRef,
                            onChange: (value) => { this._onFieldChange(fieldState.fieldInfo.InternalName, value); },
                            onValidation: (validationError) => { this._onFieldValidated(fieldState.fieldInfo.InternalName, validationError) },
                            errorMessage: fieldSubmitErrors[fieldState.fieldInfo.InternalName],
                            onError: (error) => { this.setState({ fieldLoadErrors: { ...this.state.fieldLoadErrors, [fieldState.fieldInfo.InternalName]: error } }) },
                            initializationBatch: this._fieldInitBatch,
                            formMode: displayMode
                          }) :
                          <DmsDynamicField {...fieldState}
                            context={context} disabled={
                              isStatusApplicableAndDocController && displayMode === FormDisplayMode.Display ? this._fieldDisablityCheck(fieldState.fieldInfo.InternalName, DmsDocumentFormConstants.enableInApplicableViewForm) : (displayMode === FormDisplayMode.Display) || isBusy || !!fieldLoadErrors[fieldState.fieldInfo.InternalName]
                            }
                            ref={fieldRef}
                            onChange={(value) => { this._onFieldChange(fieldState.fieldInfo.InternalName, value); }}
                            onValidation={(validationError) => this._onFieldValidated(fieldState.fieldInfo.InternalName, validationError)}
                            errorMessage={fieldSubmitErrors[fieldState.fieldInfo.InternalName]}
                            onError={(error) => { this.setState({ fieldLoadErrors: { ...this.state.fieldLoadErrors, [fieldState.fieldInfo.InternalName]: error } }) }}
                            initializationBatch={this._fieldInitBatch}
                            formMode={displayMode}
                          />
                        }
                      </div>
                    })
                    }
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
        <div className={styles.footer}>
          {(displayMode === FormDisplayMode.Edit || displayMode === FormDisplayMode.New) &&
            <>
              <DefaultButton text="Cancel" onClick={this._onClose.bind(this)} disabled={isBusy} />
              <PrimaryButton text="Save" onClick={this._onSave.bind(this)} disabled={isBusy} />
            </>
          }
          {displayMode === FormDisplayMode.Display && !isStatusApplicableAndDocController &&
            <DefaultButton text="Close" onClick={this._onClose.bind(this)} />
          }
          {displayMode === FormDisplayMode.Display && isStatusApplicableAndDocController &&
            <>
              <DefaultButton text="Close" onClick={this._onClose.bind(this)} />
              <PrimaryButton text="Save" onClick={this._onSave.bind(this)} disabled={isBusy} />
            </>
          }
        </div>
        <Dialog
          hidden={!currentPresaveOperation}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.RunninPresaveOperations,
            subText: currentPresaveOperation?.status === 'running' ? `${currentPresaveOperation?.label} (${currentPresaveOperation?.index}/${currentPresaveOperation?.count})` :
              currentPresaveOperation?.status === 'success' ? `${currentPresaveOperation?.label} ✅` :
                currentPresaveOperation?.status === 'error' ? `${currentPresaveOperation?.label} ❌` : '',
          }}
        >
          <DialogFooter>
            {((currentPresaveOperation?.status === 'success' && currentPresaveOperation?.index === currentPresaveOperation?.count) ||
              currentPresaveOperation?.status === 'error') &&
              <DefaultButton text="Close" onClick={() => this.setState({ currentPresaveOperation: undefined })} />
            }
          </DialogFooter>
        </Dialog>
      </div>
    </>;
  }

  private _fieldDisablityCheck(name: string, properties: string[]): boolean {
    return !properties.includes(name);
  }

  private _onFieldValidated(name: string, validationError: string | undefined): void {
    const { fieldValidationErrors } = this.state;
    const newFieldValidationErrors = { ...fieldValidationErrors };
    if (!validationError && newFieldValidationErrors[name]) {
      delete newFieldValidationErrors[name];
    }
    if (validationError) {
      newFieldValidationErrors[name] = validationError;
    }
    this.setState({ fieldValidationErrors: newFieldValidationErrors });
  }
  private _onFieldChange(name: string, value: unknown): void {
    this.setState((prevState) => {
      this.setState({
        item: { ...prevState.item, [name]: value }
      });
    }, () => {
      if (this.props.onChange) this.props.onChange(this.state.item);
    });
  }
}

