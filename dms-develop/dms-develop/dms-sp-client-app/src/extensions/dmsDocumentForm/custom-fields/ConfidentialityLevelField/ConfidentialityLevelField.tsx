/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from 'react';
import styles from './ConfidentialityLevelField.module.scss';
import { ComboBox, IComboBox, IComboBoxOption, IPersonaProps } from '@fluentui/react';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import DmsDynamicFieldBase, { IDmsDynamicFieldProps, IDmsDynamicFieldState } from '../../components/DmsDynamicFieldBase';
import { Logger } from '@pnp/logging';
import * as strings from 'DmsDocumentFormFormCustomizerStrings';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SPFI } from '@pnp/sp';
import { AadHttpClient } from '@microsoft/sp-http';

export interface IConfidentialityLevelFieldValue {
  value?: string;
  users?: Partial<ISiteUserInfo>[];
}
export interface IConfidentialityLevelFieldState extends IDmsDynamicFieldState {
  value: string;
  internalValue: IConfidentialityLevelFieldValue;
  groupIds: number[];
  hasChanged?: boolean;
  validationError?: { level?: string; users?: string; };
}
const LOG_SOURCE: string = 'ConfidentialityLevelField';
export default class ConfidentialityLevelField extends DmsDynamicFieldBase<IConfidentialityLevelFieldState> {

  constructor(props: IDmsDynamicFieldProps) {
    super(props);
  }

  public componentDidMount(): void {
    const { fieldInfo, onError } = this.props;
    if (super.componentDidMount) super.componentDidMount();
    this.setState({ isLoading: true });
    this.initializeField().catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error initializing field ${fieldInfo.Title}, ${error}`));
      console.error(error);
      this.setState({ isErrored: true, isLoading: false });
      if (onError) onError(new Error(`Error initializing field ${fieldInfo.Title}: ${error}`));
    });
  }

  public render(): React.ReactElement<{}> {
    const { fieldLink, fieldInfo, disabled: propsDisabled } = this.props;
    const { validationError, groupIds, internalValue , isLoading} = this.state;
    const disabled = propsDisabled || fieldInfo.ReadOnlyField || isLoading;
    return <>
      <div className={styles.ConfidentialityLevelField}>
        <div>
          <ComboBox
            required={fieldLink?.Required}
            selectedKey={internalValue?.value}
            label={fieldInfo.Title}
            autoComplete="on"
            allowFreeform={false}
            errorMessage={validationError?.level}
            options={this._getComboboxOptions()}
            onChange={this._onComboBoxChange.bind(this)}
            disabled={disabled}
          />
        </div>
        {(internalValue?.value === 'Confidential' || internalValue?.value === 'Private') &&
          <div>
            <PeoplePicker
              errorMessage={validationError?.users}
              context={this._peoplePickerContext}
              titleText={`${fieldInfo.Title} ${strings.ConfidentialityLevelUsers}`}
              groupId={groupIds}
              required={true}
              disabled={disabled}
              onChange={this._onPeoplePickerChange.bind(this)}
              personSelectionLimit={20}
              principalTypes={[PrincipalType.User]}
              defaultSelectedUsers={internalValue?.users?.map(u => `${u.Email}/${u.Title}`)}
              resolveDelay={1000} ensureUser />
          </div>
        }
      </div>
    </>;
  }

  public override async initializeField(): Promise<void> {
    const {
      context,
      initializationBatch,
      dmsClient,
      additionalParams,
      value } = this.props;
    const sp = initializationBatch || this._sp;
    const groups = await sp.web.siteGroups();
    if (!additionalParams || !dmsClient) {
      throw new Error('Missing required parameters');
    }
    const response = context.itemId ? await dmsClient?.fetch(`${additionalParams.confidentialityEndpoint}&siteUrl=${encodeURIComponent(context.pageContext.web.serverRelativeUrl)}&libraryId=${context.list.guid.toString()}&itemId=${context.itemId}`,
      AadHttpClient.configurations.v1, { method: 'GET' }) : undefined;
    if (response && !response.ok) {
      throw new Error(`Error fetching confidential users: ${response.statusText}`);
    }
    const users: { Id: number; Email: string; Title: string; }[] = await response?.json();
    this.setState({
      isLoading: false,
      internalValue: { value: value as string, users },
      groupIds: groups.map(group => group.Id),
      value: value as string
    })
  }
  private _getComboboxOptions(): IComboBoxOption[] {
    const { fieldInfo } = this.props;
    const { internalValue } = this.state;
    const choices = fieldInfo.Choices ? fieldInfo.Choices.map(choice => ({ key: choice, text: choice })) : [];
    const value = internalValue?.value;
    if (value && !choices.find(c => c.key === value)) choices.push({ key: value as string, text: value as string });
    return choices;
  }
  private _onComboBoxChange(event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void {
    const { fieldInfo } = this.props;
    const { internalValue: currentValue } = this.state;
    const hasChanged = currentValue.value !== option?.key;
    this.setState({
      value: option?.key as string,
      internalValue: { ...currentValue, value: option?.key as string },
      isTouched: true,
      hasChanged
    });
    if (this.props.onChange) {
      this.props.onChange(value);
    }
    this.validate(this._sp, { ...currentValue, value: option?.key as string }, true, false, 'level').catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: { level: strings.onErrorDuringValidation } });
    });
  }
  private _onPeoplePickerChange(items: (IPersonaProps & { loginName: string, email: string })[]): void {
    const { fieldInfo } = this.props;
    const { internalValue: currentValue } = this.state;
    const newValue: Partial<ISiteUserInfo>[] | undefined = items.map(i => ({
      Email: i.secondaryText,
      Id: +(i.id || 0),
      Title: i.text,
      LoginName: i.loginName
    }));
    this.setState({
      internalValue: { ...currentValue, users: newValue },
      isTouched: true,
      hasChanged: true
    });
    if (this.props.onChange) {
      this.props.onChange(newValue);
    }
    this.validate(this._sp, { ...currentValue, users: newValue }, true, true, 'users').catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error validating field ${fieldInfo.InternalName}, ${error}`));
      console.error(error);
      this.setState({ validationError: { level: strings.onErrorDuringValidation } });
    });
  }
  public async validate(sp: SPFI, newValue?: IConfidentialityLevelFieldValue, useNewValue?: boolean, ignoreRequiredIfNotTouched?: boolean, key?: string): Promise<string | { [key: string]: string | undefined } | undefined> {
    const { fieldLink, onValidation } = this.props;
    const { internalValue, validationError } = this.state;
    const value = useNewValue ? newValue : internalValue;
    let newValidationError: { level?: string; users?: string; } | undefined;
    if (fieldLink?.Required && (!ignoreRequiredIfNotTouched && !this.state.isTouched) && !value?.value) {
      newValidationError = { level: strings.validationErrorRequiredMissingValue };
    } else {
      if ((key === 'users' || !key) && (value?.value === 'Confidential' || value?.value === 'Private')) {
        if (!value?.users || value?.users.length === 0) {
          newValidationError = { ...validationError, users: strings.validationErrorRequiredMissingValue };
        }
      }
    }
    this.setState({
      validationError: newValidationError
    });
    if (onValidation) {
      if (!newValidationError) {
        onValidation();
      }
      else {
        const newValidationErrorObj = newValidationError as { [key: string]: string | undefined };
        onValidation(`${Object.keys(newValidationErrorObj).
          filter(k => newValidationErrorObj[k])
          .map(k => `${k}: ${newValidationErrorObj[k]}`).join(', ')}
        `);
      }
    }
    return newValidationError;
  }
}

