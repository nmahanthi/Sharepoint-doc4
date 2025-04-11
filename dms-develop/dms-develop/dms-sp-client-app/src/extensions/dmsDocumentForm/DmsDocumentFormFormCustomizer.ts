import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { FormDisplayMode, Log } from '@microsoft/sp-core-library';
import {
  ThemeProvider as SPThemeProvider, IReadonlyTheme
} from '@microsoft/sp-component-base';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import DmsDocumentForm, { IDmsDocumentFormProps } from './components/DmsDocumentForm';
import { DmsDocumentFormConstants } from '../../constants/index';
// import { getGraph, getSP } from '../../pnpjs-config';
import { getGraph, getSPCacheNever } from '../../pnpjs-config';
import { IAdditionalFieldOption } from './components/DmsDynamicFieldBase';
import { AadHttpClient } from '@microsoft/sp-http';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDmsDocumentFormFormCustomizerProperties {
  // This is an example; replace with your own property
  sectionsToFieldsMap: { [sectionName: string]: string[] };
  hideInEditForm: string[]; hideInNewForm: string[]; hideInViewForm: string[];
  disableInEditForm: string[]; disableInNewForm: string[];
  dmsUserApiAppId: string;
  confidentialityEndpoint: string;
  updateApplicableDocumentEndpoint: string;
}

const LOG_SOURCE: string = 'DmsDocumentFormFormCustomizer';

export default class DmsDocumentFormFormCustomizer
  extends BaseFormCustomizer<IDmsDocumentFormFormCustomizerProperties> {
  private _themeProvider: SPThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private _dmsClient: AadHttpClient;

  public async onInit(): Promise<void> {
    await super.onInit();
    this._dmsClient = await this.context.aadHttpClientFactory.getClient(this.properties.dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);

    this._themeProvider = this.context.serviceScope.consume(SPThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated DmsDocumentFormFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    // getSP(this.context);
    getSPCacheNever(this.context);
    getGraph(this.context);
  }

  public render(): void {
    const {
      sectionsToFieldsMap: propsSectionsToFieldsMap,
      hideInEditForm: propsHideInEditForm,
      hideInNewForm: propsHideInNewForm,
      hideInViewForm: propsHideInViewForm,
      disableInEditForm: propsDisableInEditForm,
      disableInNewForm: propsDisableInNewForm,
      confidentialityEndpoint: propsConfidentialityEndpoint,
      updateApplicableDocumentEndpoint:propsUpdateApplicableDocumentEndpoint
    } = this.properties;
    const {
      defaultSectionsToFieldsMap,
      hideInEditForm: defaultHideInEditForm, hideInNewForm: defaultHideInNewForm, hideInViewForm: defaultHideInViewForm,
      disableInEditForm: defaultDisableInEditForm, disableInNewForm: defaultDisableInNewForm,
      confidentialityEndpoint: defaultConfidentialityEndpoint, updateApplicableDocumentEndpoint: defaultUpdateApplicableDocumentEndpoint
    } = DmsDocumentFormConstants;
    const sectionsToFieldsMap = propsSectionsToFieldsMap || defaultSectionsToFieldsMap;
    const hideInEditForm = propsHideInEditForm || defaultHideInEditForm;
    const hideInNewForm = propsHideInNewForm || defaultHideInNewForm;
    const hideInViewForm = propsHideInViewForm || defaultHideInViewForm;
    const disableInEditForm = propsDisableInEditForm || defaultDisableInEditForm;
    const disableInNewForm = propsDisableInNewForm || defaultDisableInNewForm;

    const fieldNames = Object.values(sectionsToFieldsMap).reduce((acc, val) => acc.concat(val), []);
    const additionalFieldOptions = fieldNames.reduce((acc, fieldName) => {
      acc[fieldName] = {
        [FormDisplayMode.New]: {
          hidden: Promise.resolve(hideInNewForm.includes(fieldName)),
          disabled: Promise.resolve(disableInNewForm.includes(fieldName))
        },
        [FormDisplayMode.Edit]: {
          hidden: Promise.resolve(hideInEditForm.includes(fieldName)),
          disabled: Promise.resolve(disableInEditForm.includes(fieldName))
        },
        [FormDisplayMode.Display]: {
          hidden: Promise.resolve(hideInViewForm.includes(fieldName))
        }
      };
      return acc;
    }, {} as { [fieldName: string]: IAdditionalFieldOption });
    const dmsDocumentForm: React.ReactElement<{}> =
      React.createElement(DmsDocumentForm, {
        context: this.context,
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose,
        theme: this._themeVariant,
        sectionsToFieldsMap: sectionsToFieldsMap || defaultSectionsToFieldsMap,
        additionalFieldOptions,
        dmsClient: this._dmsClient,
        confidentialityEndpoint: propsConfidentialityEndpoint || defaultConfidentialityEndpoint,
        updateApplicableDocumentEndpoint: propsUpdateApplicableDocumentEndpoint || defaultUpdateApplicableDocumentEndpoint
      } as IDmsDocumentFormProps);

    ReactDOM.render(dmsDocumentForm, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose = (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
