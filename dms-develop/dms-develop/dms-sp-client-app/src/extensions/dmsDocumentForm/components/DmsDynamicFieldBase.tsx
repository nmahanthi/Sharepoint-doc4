import * as React from 'react';
import { IFieldInfo } from '@pnp/sp/fields';
import { IFieldLinkInfo } from '@pnp/sp/content-types';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { SPFI } from '@pnp/sp';
import * as strings from 'DmsDocumentFormFormCustomizerStrings';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { getSP } from '../../../pnpjs-config';
import { IListItemFormUpdateValue } from '@pnp/sp/lists';

export type DmsDynamicFieldValidator = (value: unknown, _: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => Promise<string | { [key: string]: string | undefined } | undefined> | string | { [key: string]: string | undefined } | undefined;

export interface IAdditionalFieldOption {
  [mode: number]: {
    hidden?: Promise<boolean>;
    disabled?: Promise<boolean>;
  };
  validator?: DmsDynamicFieldValidator;
  renderer?: (props: IDmsDynamicFieldProps & { ref: React.RefObject<unknown> }) => React.ReactElement<IDmsDynamicFieldProps, typeof DmsDynamicFieldBase>;
  preSaveOperation?: (sp: SPFI, httpClient: SPHttpClient, ref: React.RefObject<DmsDynamicFieldBase>, values:IListItemFormUpdateValue[]) => Promise<void> | undefined;
  preSaveOperationLabel?: string;
}

export interface IDmsDynamicFieldProps {
  formMode: FormDisplayMode;
  readonly?: boolean;
  disabled?: boolean;
  value: unknown;
  data: unknown;
  fieldInfo: IFieldInfo;
  fieldLink?: IFieldLinkInfo;
  name: string;
  additionalFieldOptions?: IAdditionalFieldOption;
  onChange?: (value: unknown) => void;
  onValidation?: (validationError?: string) => void;
  onError?: (error: Error) => void;
  context: FormCustomizerContext;
  errorMessage?: string;
  initializationBatch?: SPFI;
  dmsClient?:AadHttpClient;
  additionalParams?: { [key: string]: string };
  onInit?: () => void;
}

export interface IDmsDynamicFieldState {
  validationError?: string | { [key: string]: string | undefined };
  value: unknown;
  isLoading?: boolean;
  isErrored?: boolean;
  isTouched?: boolean;
  isAdditionalOptionReadonly?: boolean;
  isAdditionalOptionHidden?: boolean;
}

export default class DmsDynamicFieldBase<T extends IDmsDynamicFieldState = IDmsDynamicFieldState> extends React.Component<IDmsDynamicFieldProps, T> {
  protected _peoplePickerContext: IPeoplePickerContext;
  protected _sp: SPFI;
  /**
   *
   */
  constructor(props: IDmsDynamicFieldProps) {
    super(props);
    this.state = {
      value: undefined,
      validationError: props.errorMessage,
    } as T;
    this._peoplePickerContext = {
      absoluteUrl: props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: props.context.msGraphClientFactory as never,
      spHttpClient: props.context.spHttpClient as never
    };
    this._sp = getSP();
  }
  //#region validations
  public async validate(sp: SPFI, newValue?: unknown, useNewValue?: boolean, ignoreRequiredIfNotTouched?: boolean, key?: string): Promise<string | { [key: string]: string | undefined } | undefined> {
    const { fieldInfo, additionalFieldOptions, onValidation } = this.props;
    const { value: currentValue } = this.state;
    const value = useNewValue ? newValue : currentValue;
    const validator = additionalFieldOptions?.validator || this._validators[fieldInfo.TypeAsString];
    let newValidationError: { [key: string]: string | undefined } | string | undefined;
    if (validator) {
      const validationResult = await validator(value, sp, key, ignoreRequiredIfNotTouched);
      newValidationError = validationResult;
    } else {
      newValidationError = this._handleRequiredValidation(value, key, ignoreRequiredIfNotTouched);
    }
    this.setState({
      validationError: newValidationError
    });
    if (onValidation) {
      if (!newValidationError) {
        onValidation();
      }
      else {
        if (newValidationError?.indexOf) {
          onValidation(newValidationError as string);
        }
        else {
          const newValidationErrorObj = newValidationError as { [key: string]: string | undefined };
          onValidation(`${Object.keys(newValidationErrorObj).
            filter(k => newValidationErrorObj[k])
            .map(k => `${k}: ${newValidationErrorObj[k]}`).join(', ')}
          `);
        }
      }
    }
    return newValidationError;
  }

  private _handleRequiredValidation(value: unknown, key?: string, ignoreRequiredIfNotTouched?: boolean): string | { [key: string]: string | undefined } | undefined {
    const { fieldLink, fieldInfo } = this.props;
    const { isTouched } = this.state;
    if (ignoreRequiredIfNotTouched && !isTouched) return;
    switch (fieldInfo.TypeAsString) {
      case 'URL':
        return this._handleRequiredUrlValidation(value, key || '');
      default:
        if (fieldLink?.Required && (!value || (Array.isArray(value) && value.length === 0))) {
          return strings.validationErrorRequiredMissingValue;
        }
        return undefined;
    }
  }
  private _handleRequiredUrlValidation(value: unknown, key: string): { [key: string]: string | undefined } | undefined {
    const { fieldLink } = this.props;
    const oldValidationError = this.state.validationError as { [key: string]: string | undefined };
    const newValidationError = { ...(oldValidationError || {}) };
    if (fieldLink?.Required && !value && key === 'Url') {
      return { ...newValidationError, Url: strings.validationErrorRequiredMissingValue };
    }
    return;
  }

  private _validators: { [fieldType: string]: (value: unknown, sp: SPFI, key?: string) => Promise<string | { [key: string]: string | undefined } | undefined> | string | undefined } = {
    'UserMulti': async (value: unknown, sp: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      if (!value) return;
      const users = value as ISiteUserInfo[];
      if (!users) return;
      if (users.filter(u => !u.Email).length > 0) return strings.validationErrorInvalidFormat;
      try {
        await Promise.all(users.map(u => sp.web.ensureUser(u.Email)));
        return undefined;
      } catch (error) {
        return strings.validationErrorInvalidFormat;
      }
    },
    'User': async (value: unknown, sp: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      if (!value) return;
      const users = value as ISiteUserInfo[];
      if (!users) return;
      if (users.filter(u => !u.Email).length > 0) return strings.validationErrorInvalidFormat;
      try {
        await Promise.all(users.map(u => sp.web.ensureUser(u.Email)));
        return undefined;
      } catch (error) {
        return strings.validationErrorInvalidFormat;
      }
    },
    'DateTime': async (value: unknown, _: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      if (!value) return;
      const dateValue = new Date(value as string);
      return dateValue?.getTime && !isNaN(dateValue.getTime()) ? undefined : strings.validationErrorInvalidFormat;
    },
    'TaxonomyFieldType': async (value: unknown, _: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const { context, fieldInfo } = this.props;
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      if (!value) return;
      const terms = value as { id: string }[];
      const client = context.spHttpClient;
      const response = await client.get(`/_api/v2.1/termstore/sets/${(fieldInfo as IFieldInfo & { TermSetId: string }).TermSetId}/terms?$select=id`, SPHttpClient.configurations.v1);
      const termsData = (await response.json()).value as { id: string }[];
      return terms.map(t => t.id).filter(id => !termsData.find(td => td.id === id)).length > 0 ? strings.validationErrorInvalidFormat : undefined;
    },
    'TaxonomyFieldTypeMulti': async (value: unknown, _: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const { context, fieldInfo } = this.props;
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      if (!value) return;
      const terms = value as { id: string }[];
      const client = context.spHttpClient;
      const response = await client.get(`/_api/v2.1/termstore/sets/${(fieldInfo as IFieldInfo & { TermSetId: string }).TermSetId}/terms?$select=id`, SPHttpClient.configurations.v1);
      const termsData = (await response.json()).value as { id: string }[];
      return terms.map(t => t.id).filter(id => !termsData.find(td => td.id === id)).length > 0 ? strings.validationErrorInvalidFormat : undefined;
    },
    'URL': async (value: unknown, _: SPFI, key?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const requiredError = this._handleRequiredValidation(value, key, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      const oldValidationError = this.state.validationError as { [key: string]: string | undefined };
      const newValidationError = { ...(oldValidationError || {}) };
      const url = (value as { Description: string; Url: string })?.Url;
      const description = (value as { Description: string; Url: string })?.Description;
      if (!key || key === 'Url') {
        if (url && !url.match(/^https?:\/\//)) return { ...newValidationError, Url: strings.validationErrorInvalidFormat };
      }
      if (!url && description) return { ...newValidationError, Url: strings.validationErrorRequiredMissingValue };
      return;
    },
    'File': async (value: unknown, _: SPFI, __?: string, ignoreRequiredIfNotTouched?: boolean) => {
      const requiredError = this._handleRequiredValidation(value, undefined, ignoreRequiredIfNotTouched);
      if (requiredError) return requiredError;
      const stringValue = value as string;
      if (
        stringValue.indexOf('/') > -1 ||
        stringValue.indexOf('\\') > -1 ||
        stringValue.indexOf(':') > -1 ||
        stringValue.indexOf('?') > -1 ||
        stringValue.indexOf('*') > -1 ||
        stringValue.indexOf('"') > -1 ||
        stringValue.indexOf('<') > -1 ||
        stringValue.indexOf('>') > -1 ||
        stringValue.indexOf('|') > -1
      ) return strings.validationErrorInvalidFormat;
      return;
    }
  };
  //#endregion

  //#region initializations
  public async initializeField(): Promise<void> {
    const { fieldInfo, initializationBatch } = this.props;
    if (this._initializers[fieldInfo.TypeAsString]) {
      const sp = initializationBatch || this._sp;
      await this._initializers[fieldInfo.TypeAsString](sp);
    } else {
      this.setState({
        isLoading: false,
        value: this.props.value
      });
    }
  }
  private _initializers: { [fieldType: string]: (sp: SPFI) => Promise<void> } = {
    'UserMulti': async (sp: SPFI) => {
      const { data } = this.props;
      const typedData1 = data as { email: string, value: string, id: number, picture: string }[];
      const typedData2 = data as { id: number; title: string }[];
      let resolvedUsers: ISiteUserInfo[] = [];
      if (data) {
        resolvedUsers = await Promise.all(typedData1.map((u, i) => {
          if (u.email) return sp.web.ensureUser(u.email);
          return sp.web.getUserById(typedData2[i].id)();
        }));
      }
      this.setState({
        value: data ? resolvedUsers : []
      });
    },
    'User': async (sp: SPFI) => {
      const { data } = this.props;
      const typedData1 = data as { email: string, value: string, id: number, picture: string }[];
      const typedData2 = data as { id: number; title: string }[];
      let resolvedUsers: ISiteUserInfo[] = [];
      if (data) {
        resolvedUsers = await Promise.all(typedData1.map((u, i) => {
          if (u.email) return sp.web.ensureUser(u.email);
          return sp.web.getUserById(typedData2[i].id)();
        }));
      }
      this.setState({
        value: data ? resolvedUsers : []
      });
    },
    'DateTime': async () => {
      const { value } = this.props;
      if (!value) return;
      const dateValue = new Date(value as string);
      const processedValue = `${dateValue?.getFullYear()}-${`${((dateValue?.getMonth() || 0) + 1)}`.padStart(2, '0')}-${`${dateValue?.getDate()}`.padStart(2, '0')}`;
      this.setState({
        value: processedValue
      });
    },
    'File': async () => {
      const { data } = this.props;
      this.setState({
        value: data
      });
    },
    'TaxonomyFieldType': async () => {
      const { value, data } = this.props;
      const typedValue = (value as { TermGuid: string; Label: string });
      if (!typedValue) return;
      let processedValue: { labels: { name: string; isDefault: boolean; languageTag: string }[], id: string }[];
      if (typedValue.TermGuid) {
        if (value && !data) throw new Error('Passing data as stream is mandatory for taxonomy fields');
        processedValue = [{ labels: [{ name: (data as { Label: string }).Label, isDefault: true, languageTag: "en-US" }], id: typedValue.TermGuid }];
      } else {
        if (!/(?:-\d;#[^;]+(?:;#)?)+/.test(value as string)) {
          throw new Error('Invalid taxonomy field value');
        }
        const values = (value as string).split(';#').filter(v => /.*\|(?:[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+)/.test(v));
        processedValue = values.map(v => {
          const [label, id] = v.split('|');
          return { labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id };
        });
      }
      this.setState({
        value: processedValue
      });
    },
    'TaxonomyFieldTypeMulti': async () => {
      const { value } = this.props;
      const typedValue = (value as { TermGuid: string; Label: string }[]);
      if (!typedValue) return;
      let processedValue: { labels: { name: string; isDefault: boolean; languageTag: string }[], id: string }[] | undefined;
      if (typedValue.length > 0 && typedValue[0].TermGuid) {
        processedValue = !value ? undefined :
          typedValue.map(v => (
            { labels: [{ name: v.Label, isDefault: true, languageTag: "en-US" }], id: v.TermGuid }
          ));
      } else {
        if ((value as string).charAt) {
          if (!/(?:-\d;#[^;]+(?:;#)?)+/.test(value as string)) {
            throw new Error('Invalid taxonomy field value');
          }
          const values = (value as string).split(';#').filter(v => /.*\|(?:[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+)/.test(v));
          processedValue = values.map(v => {
            const [label, id] = v.split('|');
            return { labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id };
          });
        }
      }
      this.setState({
        value: processedValue
      });
    },
    'Note': async () => {
      const { value, data } = this.props;
      this.setState({
        value: value || data
      });
    },
    'Boolean': async () => {
      const { value } = this.props;
      this.setState({
        value: (value === 1 || value === true || value === 'true' || value === '1')
      });
    }
  };
  //#endregion
}

