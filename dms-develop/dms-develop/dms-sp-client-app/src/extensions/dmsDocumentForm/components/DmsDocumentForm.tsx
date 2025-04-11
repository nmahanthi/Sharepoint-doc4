import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { MessageBar, MessageBarType, Label, ActionButton } from '@fluentui/react';
import styles from './DmsDocumentForm.module.scss';
import * as strings from 'DmsDocumentFormFormCustomizerStrings';
import { SPFI } from '@pnp/sp/fi';
// import { getSP, getSPCacheNever } from '../../../pnpjs-config';
import { getSP, getSPCacheNever } from '../../../pnpjs-config';
import { IFieldInfo } from '@pnp/sp/fields';
import { Logger } from '@pnp/logging';
import DmsDocumentFormSqueleton from './DmsDocumentFormSkeleton';
import { CustomPermissionsListTitle, FieldNames, ListUrls } from '../../../constants';
import { CustomPermissionsLevels } from '../../../constants';
import {
  IReadonlyTheme
} from '@microsoft/sp-component-base';
import { ThemeProvider as FluentThemeProvider } from '@fluentui/react';
import { IFieldLinkInfo } from '@pnp/sp/content-types';
import DmsDynamicForm from './DmsDynamicForm';
import { IListItemFormUpdateValue, RenderListDataOptions } from '@pnp/sp/lists';
import { PermissionKind } from '@pnp/sp/security';
import "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/content-types";
import "@pnp/sp/items";
import { customValidators } from '../custom-actions';
import { IAdditionalFieldOption } from './DmsDynamicFieldBase';
import ConfidentialityLevelField from '../custom-fields/ConfidentialityLevelField/ConfidentialityLevelField';
import AlstomPartNumbersField from '../custom-fields/AlstomPartNumbersField/AlstomPartNumbersField';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';
import { IConfidentialityRequest } from '../../../interfaces/IConfidentialityRequest';
import PLOwnershipField from '../custom-fields/ProjectPLOwnershipField/ProjectPLOwnershipField';

export interface IDmsDocumentFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
  sectionsToFieldsMap: { [sectionName: string]: string[] };
  additionalFieldOptions?: { [fieldInternalName: string]: IAdditionalFieldOption };
  theme: IReadonlyTheme;
  dmsClient: AadHttpClient,
  confidentialityEndpoint: string;
  updateApplicableDocumentEndpoint:string;
}
export interface IDmsDocumentFormState {
  selectedSection?: string;
  fields: IFieldInfo[];
  fieldLinks: IFieldLinkInfo[];
  isDocumentController?: boolean;
  isProjectAdmin?: boolean;
  isError?: boolean;
  isLoading?: boolean;
  validationError?: string;
  formErrors: string[];
  sectionsValidationErrors?: { [sectionName: string]: number };
  canEdit?: boolean;
  formValidationFlag?: boolean;
  item?: { ID: number | undefined;[key: string]: unknown };
  itemData?: never;
  libraries?:string[]
}

const LOG_SOURCE: string = 'DmsDocumentForm';
export default class DmsDocumentForm extends React.Component<IDmsDocumentFormProps, IDmsDocumentFormState> {
  private _sp: SPFI;
  private _notificationContainerRef = React.createRef<HTMLDivElement>();
  /**
   *
   */
  constructor(props: IDmsDocumentFormProps) {
    super(props);
    this.state = {
      fieldLinks: [],
      fields: [],
      selectedSection: props.sectionsToFieldsMap ? Object.keys(props.sectionsToFieldsMap)[0] : '',
      isLoading: true,
      sectionsValidationErrors: {},
      formValidationFlag: false,
      formErrors: []
    };
    // this._sp = getSP();
    this._sp = getSPCacheNever();
  }


  public componentDidMount(): void {
    const sp = getSP();
    const [batch, execute] = sp.batched();
    const { context, sectionsToFieldsMap } = this.props;
    const fieldNames = sectionsToFieldsMap ? Object.values(sectionsToFieldsMap).reduce((acc, fields) => acc.concat(fields), [] as string[]) : undefined;
    if (!fieldNames) {
      Logger.error(new Error(`${LOG_SOURCE} no fields found in sectionsToFieldsMap`));
      this.setState({ isError: true, isLoading: false });
      return;
    }
  
    Promise.all([
      batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.ApplicableDocuments}'`)(),
      batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.PreviousVersions}'`)(),
    ]).then(([applicableLib, prevLib]) => {
      if (!applicableLib ||!applicableLib || applicableLib.length === 0 ||
        !prevLib || prevLib.length === 0) {
        console.error(new Error('Document libraries not found'));
        Logger.error(new Error(`${LOG_SOURCE}:Document libraries not found`));
      } else {
        this.setState({
          libraries:[applicableLib[0].Id,prevLib[0].Id]
        });
      }
    }).catch((error) => {
     
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
    });
    execute().catch((error) => {
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
    });

    Promise.all([
      this._getFields(),
      this._hasPermissionConfigItems(),
      context.itemId ? this._sp.web.lists.getById(context.list.guid.toString()).items.getById(context.itemId)() : Promise.resolve(),
      context.itemId ? this._sp.web.lists.getById(context.list.guid.toString()).renderListDataAsStream({
        ViewXml: `<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${context.itemId}</Value></Eq></Where></Query><ViewFields>${fieldNames.map(f => (`<FieldRef Name="${f}"/>`)).join('')}</ViewFields></View>`,
        RenderOptions: RenderListDataOptions.ListData
      }) : undefined,
      context.itemId ? this._sp.web.lists.getById(context.list.guid.toString()).items.getById(context.itemId).currentUserHasPermissions(PermissionKind.EditListItems) : undefined
    ]).then(([[fields, fieldLinks], isDocumentController, item, listDataAsStream, canEdit]) => {
      this.setState({ fields, fieldLinks, isDocumentController, isLoading: false, item, itemData: listDataAsStream?.Row[0], canEdit });
    }).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE} error in componentDidMount, ${error}`));
      this.setState({ isError: true, isLoading: false });
    });

    Promise.all([
      this._hasPermissionConfigItemsProjectAdmin(),
    ]).then(([ isProjectAdmin]) => {
      this.setState({  isProjectAdmin });
    }).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE} error in componentDidMount for Project Admin Check, ${error}`));
      this.setState({ isError: true, isLoading: false });
    });
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: DmsDocumentForm unmounted');
  }

  public componentDidUpdate(_: IDmsDocumentFormProps): void {
    if (this._notificationContainerRef.current && this._notificationContainerRef.current.childElementCount > 0) {
      this._notificationContainerRef.current.focus();
    }
  }
  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error(error, errorInfo);
    Logger.error(new Error(`${LOG_SOURCE} componentDidCatch, ${error}`));
    this.setState({ isError: true, isLoading: false });
  }


  public render(): React.ReactElement<{}> {
    const { context, sectionsToFieldsMap, displayMode,
      theme, onSave, onClose, dmsClient, updateApplicableDocumentEndpoint } = this.props;
    const { isLoading, isError, isDocumentController, validationError, isProjectAdmin,
      formErrors, canEdit, fields, fieldLinks, item, itemData } = this.state;
    let titleHeader = "";
    switch (displayMode) {
      case FormDisplayMode.New:
        titleHeader = strings.NewFormHeader;
        break;
      case FormDisplayMode.Edit:
        titleHeader = strings.EditFormHeader;
        break;
      case FormDisplayMode.Display:
        titleHeader = strings.ViewFormHeader;
        break;
    }
    const isAllowed = (displayMode === FormDisplayMode.New && isDocumentController || displayMode !== FormDisplayMode.New);
    const additionalFieldOptions = this._getAdditionalFieldOptions();
    return <>
      <FluentThemeProvider theme={theme}>
        <div className={styles.dmsDocumentForm}>
          <div className={styles.header}>
            <Label className={styles.titleHeader}>{titleHeader} {context.contentType.name}</Label>
            {displayMode === FormDisplayMode.Display && <ActionButton text={strings.EditButtonText} onClick={() => {
              window.location.href = `${context.list.serverRelativeUrl}/Forms/EditForm.aspx?ID=${context.itemId}&Source=${encodeURIComponent(location.href)}`;
            }} allowDisabledFocus disabled={!canEdit} iconProps={{ iconName: 'Edit' }} />}
          </div>
          <div className={styles.notificationsContianer} ref={this._notificationContainerRef} tabIndex={0}>
            {!isAllowed && !isLoading && !isError &&
              <MessageBar
                messageBarType={MessageBarType.blocked}
                isMultiline={false}
                truncated={false}
              >
                {strings.OnlyDocumentControllerCanCreateDocument}
              </MessageBar>
            }
            {validationError &&
              <MessageBar
                messageBarType={MessageBarType.blocked}
                isMultiline={false}
                truncated={false}
              >
                {validationError}
              </MessageBar>
            }
            {formErrors &&
              formErrors.map((error, index) => (
                <MessageBar
                  key={index}
                  messageBarType={MessageBarType.blocked}
                  isMultiline={true}
                  truncated={false}
                  onDismiss={() => { this.setState({ formErrors: formErrors.filter((_, i) => i !== index) }); }}
                >
                  {error}
                </MessageBar>))
            }
            {isError &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                truncated={false}
              >
                {strings.ErrorLoadingForm}
              </MessageBar>
            }
          </div>
          {
            isAllowed && !isLoading && !isError &&
            <>
              <DmsDynamicForm
                context={context}
                displayMode={displayMode}
                fields={fields}
                fieldLinks={fieldLinks}
                initialItem={item}
                initialItemData={itemData}
                onSave={onSave}
                onClose={onClose}
                onChange={(item) => { this.setState({ item }) }}
                sectionsToFieldsMap={sectionsToFieldsMap}
                onError={(message) => { if (formErrors.indexOf(message)) this.setState({ formErrors: [...formErrors, message] }); }}
                additionalFieldOptions={additionalFieldOptions}
                isDocumentController={isDocumentController}
                dmsClient={dmsClient}
                isProjectAdmin={isProjectAdmin}
                updateApplicableDocumentEndpoint={updateApplicableDocumentEndpoint}
              />
            </>
          }
        </div>

        {isLoading && !isError &&
          <DmsDocumentFormSqueleton />
        }
      </FluentThemeProvider>
    </>;
  }
  private async _hasPermissionConfigItems(): Promise<boolean> {
    const items = await this._sp.web.lists.getByTitle(CustomPermissionsListTitle).items.filter(`Title eq '${CustomPermissionsLevels.DocumentController}'`).top(1)();
    return items.length > 0;
  }
  private async _hasPermissionConfigItemsProjectAdmin(): Promise<boolean> {
    const items = await this._sp.web.lists.getByTitle(CustomPermissionsListTitle).items.filter(`Title eq '${CustomPermissionsLevels.ProjectAdmin}'`).top(1)();
    return items.length > 0;
  }
  private _getFields(): Promise<[IFieldInfo[], IFieldLinkInfo[]]> {
    const { context } = this.props;
    const links = this._sp.web.lists.getById(context.list.guid.toString()).contentTypes.getById(context.contentType.id).fieldLinks();
    const fields = this._sp.web.lists.getById(context.list.guid.toString()).fields<IFieldInfo[]>();
    return Promise.all([fields, links]);
  }
  private _getAdditionalFieldOptions(): { [fieldInternalName: string]: IAdditionalFieldOption } {
    const { additionalFieldOptions: propadditionalFieldOptions, dmsClient, context, confidentialityEndpoint } = this.props;
    let additionalFieldOptions: { [fieldInternalName: string]: IAdditionalFieldOption } = { ...propadditionalFieldOptions };
    const fOwithCustomValidations = Object.keys(customValidators).reduce((acc, fieldName) => {
      const fieldOption = additionalFieldOptions?.[fieldName] || {};
      return {
        ...acc,
        [fieldName]: {
          ...fieldOption,
          validator: customValidators[fieldName].bind(this)
        }
      };
    }, {} as { [fieldInternalName: string]: IAdditionalFieldOption });
    additionalFieldOptions = { ...additionalFieldOptions, ...fOwithCustomValidations };
    // Override sensitivity label with custom field options
    additionalFieldOptions[FieldNames.ConfidentialityLevel] = {
      ...additionalFieldOptions[FieldNames.ConfidentialityLevel],
      renderer: (props) => <ConfidentialityLevelField {...{ 
        ...props,
        dmsClient,
        additionalParams: { confidentialityEndpoint },
        ref: props.ref as React.RefObject<ConfidentialityLevelField>
      }} />,
      preSaveOperation: (_: SPFI, __: SPHttpClient, ref: React.RefObject<ConfidentialityLevelField>, values: IListItemFormUpdateValue[]) => {
        if (!ref.current)
          throw new Error("ConfidentialityLevelField ref is not set");
        const { internalValue: value, hasChanged } = ref.current.state;
        if (!hasChanged) return;
        const confidentialityLevel = value.value;
        const folderName = `${values.find(v => v.FieldName === FieldNames.ProjectReference)?.FieldValue}-${values.find(v => v.FieldName === FieldNames.Revision)?.FieldValue}`;
        return dmsClient.post(confidentialityEndpoint, AadHttpClient.configurations.v1, {
          body: JSON.stringify({
            confidentiality: confidentialityLevel,
            folderUrl: `${context.list.serverRelativeUrl}/${folderName}`,
            itemId: context.itemId,
            libraryId: context.list.guid.toString(),
            siteUrl: context.pageContext.web.serverRelativeUrl,
            users: value?.users?.map(u => u.Email),
          } as IConfidentialityRequest)
        }).then(response => {
          if (!response.ok) {
            throw new Error(response.statusText);
          }
        });
      },
      preSaveOperationLabel: strings.ConfidentialityLevelPreSaveOperationLabel
    };

    additionalFieldOptions[FieldNames.AlstomPartNumbers] = {
      ...additionalFieldOptions[FieldNames.AlstomPartNumbers],
      renderer: (props) => (
        <AlstomPartNumbersField
          {...{ ...props, ref: props.ref as React.RefObject<AlstomPartNumbersField> }}
        />
      ),

      preSaveOperation: async (_: SPFI, __: SPHttpClient, ref: React.RefObject<AlstomPartNumbersField>, values: IListItemFormUpdateValue[]) => {
        if (ref.current) {
          const alstomPartNumbers = ref.current.partNumbersToSingleLine();
          const partNumbersFieldValue = values.find((v) => v.FieldName === FieldNames.AlstomPartNumbers);
          if (partNumbersFieldValue) {
            partNumbersFieldValue.FieldValue = alstomPartNumbers;

          }
        }
      },
    };
    additionalFieldOptions[FieldNames.PLOwnership] = {
      ...additionalFieldOptions[FieldNames.PLOwnership],
      renderer: (props) => (
        <PLOwnershipField
          {...{ ...props, ref: props.ref as React.RefObject<PLOwnershipField> }}
        />
      )
    };
    return additionalFieldOptions;
  }

}
