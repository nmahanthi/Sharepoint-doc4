import * as React from 'react';
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import { IFieldInfo } from '@pnp/sp/fields';
import {
  ComboBox, IComboBox, IComboBoxOption, IPersonaProps, Label, Shimmer, Text, TextField,
  Toggle
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as strings from 'DmsDocumentFormFormCustomizerStrings';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import styles from './DmsDocumentForm.module.scss';
import { ModernTaxonomyPicker } from '@pnp/spfx-controls-react/lib/ModernTaxonomyPicker';
import { Logger } from '@pnp/logging';
import DmsFieldValidationWrapper from './DmsFieldValidationWrapper';
import DmsDynamicFieldBase, { IDmsDynamicFieldProps } from './DmsDynamicFieldBase';

const LOG_SOURCE: string = 'DmsDynamicField';

export default class DmsDynamicField extends DmsDynamicFieldBase {

  constructor(props: IDmsDynamicFieldProps) {
    super(props);
  }

  public componentDidMount(): void {
    const { fieldInfo, onError, additionalFieldOptions, formMode } = this.props;
    this.setState({
      isLoading: true
    });
    Promise.all([
      additionalFieldOptions ? additionalFieldOptions[formMode]?.hidden : Promise.resolve(false),
      additionalFieldOptions ? additionalFieldOptions[formMode]?.disabled : Promise.resolve(false),
      this.initializeField()
    ]).then((
      [hidden, disabled]
    ) => {
      this.setState({
        isLoading: false,
        isAdditionalOptionHidden: hidden,
        isAdditionalOptionReadonly: disabled
      });
    }).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error initializing field ${fieldInfo.Title}, ${error}`));
      console.error(error);
      this.setState({ isErrored: true, isLoading: false });
      if (onError) onError(new Error(`Error initializing field ${fieldInfo.Title}: ${error}`));
    });
  }

  public componentDidUpdate(prevProps: Readonly<IDmsDynamicFieldProps>): void {
    if (prevProps.errorMessage !== this.props.errorMessage) {
      this.setState({
        validationError: this.props.errorMessage
      });
    }
  }
  //#region Render
  public render(): React.ReactElement<{}> {
    const { fieldInfo, fieldLink, disabled: propsDisabled } = this.props;
    const { value, validationError: untypedvalidationError, isLoading, isErrored, isAdditionalOptionHidden, isAdditionalOptionReadonly } = this.state;
    const personFieldInfo = fieldInfo as IFieldInfo & { SelectionMode: number | undefined; SelectionScope: number | undefined };
    const urlValidationError = (untypedvalidationError as { [key: string]: string | undefined })?.Url;
    const stringValidationError = untypedvalidationError?.charAt ? untypedvalidationError as string : undefined;
    const descriptionValidationError = (untypedvalidationError as { [key: string]: string | undefined })?.Description || stringValidationError;
    const disabled = propsDisabled || fieldInfo.ReadOnlyField || isAdditionalOptionReadonly;
    if (isErrored) {
      return <>
        <div>
          <Label>{fieldInfo.Title}</Label>
          <Text role='alert' variant="medium" styles={{ root: { color: 'red' } }}>{strings.errorLoadingField}</Text>
        </div>
      </>;
    }
    if (isLoading) {
      return <>
        <div className={styles.loadingField}>
          <Shimmer width="30%" styles={{ shimmerWrapper: { height: "18px" } }} />
          <Shimmer width="100%" styles={{ shimmerWrapper: { height: "30px" } }} />
        </div>
      </>;
    }
    if (isAdditionalOptionHidden) return <></>;
    switch (fieldInfo.TypeAsString) {
      case 'File':
      case 'Text':
        return <TextField label={fieldInfo.Title} errorMessage={stringValidationError}
          required={fieldLink?.Required} value={value as string}
          onChange={this._onChange.bind(this)} disabled={disabled} />;
      case 'Note':
        return <TextField label={fieldInfo.Title} errorMessage={stringValidationError}
          required={fieldLink?.Required} value={value as string}
          onChange={this._onChange.bind(this)} multiline resizable={false} disabled={disabled} />;
      case 'DateTime':
        return <TextField type={disabled ? undefined : 'date'} label={fieldInfo.Title} errorMessage={stringValidationError}
          required={fieldLink?.Required} value={value as string}
          onChange={this._onChange.bind(this)} disabled={disabled} />
      case 'Choice':
        return <ComboBox
          required={fieldLink?.Required}
          selectedKey={value as string}
          label={fieldInfo.Title}
          autoComplete="on"
          allowFreeform={false}
          errorMessage={stringValidationError}
          options={this._getComboboxOptions()}
          onChange={this._onComboBoxChange.bind(this)}
          disabled={disabled}
        />
      case 'MultiChoice':
        return <ComboBox
          multiSelect
          required={fieldLink?.Required}
          selectedKey={value as string}
          label={fieldInfo.Title}
          autoComplete="on"
          allowFreeform={false}
          errorMessage={stringValidationError}
          options={this._getComboboxOptions()}
          onChange={this._onComboBoxMultiChange.bind(this)} disabled={disabled} />
      case 'User':
      case 'UserMulti':
        return <PeoplePicker
          context={this._peoplePickerContext}
          titleText={fieldInfo.Title}
          personSelectionLimit={fieldInfo.TypeAsString === 'User' ? 1 : 20}
          groupId={personFieldInfo.SelectionScope}
          showtooltip={true}
          required={fieldLink?.Required}
          disabled={disabled}
          onChange={this._onPeoplePickerChange.bind(this)}
          principalTypes={personFieldInfo.SelectionMode === 0 ? [PrincipalType.User] : [PrincipalType.User, PrincipalType.SharePointGroup]}
          defaultSelectedUsers={(value as ISiteUserInfo[])?.map(u => `${u.Email}/${u.Title}`)}
          errorMessage={stringValidationError}
          resolveDelay={1000} ensureUser />
      case 'URL':
        return <>
          <Label>{fieldInfo.Title}</Label>
          <div className={styles.urlField}>
            <TextField label={strings.URLFieldDescription}
              errorMessage={descriptionValidationError}
              required={fieldLink?.Required} value={(value as { Description: string; Url: string })?.Description}
              data-property="Description"
              onChange={this._onUrlFieldChange.bind(this)} disabled={disabled} />
            <TextField label={strings.URLFieldUrl} errorMessage={urlValidationError}
              required={fieldLink?.Required} value={(value as { Description: string; Url: string })?.Url}
              data-property="Url"
              onChange={this._onUrlFieldChange.bind(this)} disabled={disabled} />
          </div>
        </>;
      case 'TaxonomyFieldType':
      case 'TaxonomyFieldTypeMulti':
        return <>
          <DmsFieldValidationWrapper validationError={stringValidationError}>
            <ModernTaxonomyPicker
              allowMultipleSelections={fieldInfo.TypeAsString === 'TaxonomyFieldTypeMulti'}
              termSetId={(fieldInfo as IFieldInfo & { TermSetId: string }).TermSetId}
              panelTitle={strings.taxPickerPanelTitle}
              label={fieldInfo.Title}
              context={this.props.context as never}
              onChange={this._onTaxPickerChange.bind(this)}
              required={fieldLink?.Required}
              initialValues={value as never[]}
              disabled={disabled}
              termPickerProps={{ itemLimit: fieldInfo.TypeAsString === 'TaxonomyFieldType' ? 1 : undefined }}
            />
          </DmsFieldValidationWrapper>
        </>;
      case 'Boolean':
        return <>
          <DmsFieldValidationWrapper validationError={stringValidationError}>
            <Toggle label={fieldInfo.Title}
              checked={value as boolean}
              onChange={this._onToggleChange.bind(this)} 
              disabled={disabled}
              onText={strings.OnText} offText={strings.OffText} />
          </DmsFieldValidationWrapper>
        </>;
      default:
        return <div>{`${fieldInfo.TypeAsString} - ${fieldInfo.Title}: ${JSON.stringify(value)}`}</div>;
    }
  }
  //#endregion

  private _getComboboxOptions(): IComboBoxOption[] {
    const { fieldInfo } = this.props;
    const { value } = this.state;
    const choices = fieldInfo.Choices ? fieldInfo.Choices.map(choice => ({ key: choice, text: choice })) : [];
    if (value && (value as string[]).push) {
      (value as string[]).filter(v => !choices.find(c => c.key === v)).forEach(v => choices.push({ key: v, text: v }));
    }
    if (value && (value as string).charAt && !choices.find(c => c.key === value)) choices.push({ key: value as string, text: value as string });
    return choices;
  }
  //#region Change handlers
  private _onChange(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement> | Date, newValue?: string): void {
    const { fieldInfo } = this.props;
    this.setState({
      value: newValue,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(newValue);
    }
    this.validate(this._sp, newValue, true, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onToggleChange(_: React.MouseEvent<HTMLElement>, checked?: boolean): void {
    const { fieldInfo } = this.props;
    this.setState({
      value: checked,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(checked);
    }
    this.validate(this._sp, checked, true, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onComboBoxChange(event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void {
    const { fieldInfo } = this.props;
    this.setState({
      value: value,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(value);
    }
    this.validate(this._sp, value, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onComboBoxMultiChange(event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void {
    const { fieldInfo } = this.props;
    const { value: prevValue } = this.state;
    const selected = option?.selected;
    let typedValue: string[];
    if (prevValue) {
      if ((prevValue as string).charAt) {
        typedValue = [prevValue as string];
      } else {
        typedValue = prevValue as string[];
      }
    } else {
      typedValue = [];
    }
    const prevSelectedKeys = typedValue ? typedValue as string[] : [];
    const newValue = selected ? [...prevSelectedKeys, option.key as string] : prevSelectedKeys.filter(k => k !== option?.key);
    this.setState({
      value: newValue,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(newValue);
    }
    this.validate(this._sp, newValue, true, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onPeoplePickerChange(items: (IPersonaProps & { loginName: string, email: string })[]): void {
    const { fieldInfo } = this.props;
    const newValue: Partial<ISiteUserInfo>[] | undefined = items.map(i => ({
      Email: i.secondaryText,
      Id: +(i.id || 0),
      Title: i.text,
      LoginName: i.loginName
    }));
    this.setState({
      value: newValue,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(newValue);
    }
    this.validate(this._sp, newValue, true, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onUrlFieldChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    const { value } = this.state;
    const { fieldInfo } = this.props;
    const property = event.currentTarget.getAttribute('data-property');
    if (!property) return;
    const processedValue = { ...(value as { Description: string | undefined; Url: string | undefined }), [property]: newValue };
    this.setState({
      value: processedValue,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(processedValue);
    }
    this.validate(this._sp, processedValue, true, true, property).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }
  private _onTaxPickerChange(terms: never[]): void {
    const { fieldInfo } = this.props;
    this.setState({
      value: terms,
      isTouched: true
    });
    if (this.props.onChange) {
      this.props.onChange(terms);
    }
    this.validate(this._sp, terms, true, true).catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: strings.onErrorDuringValidation });
    });
  }

}

