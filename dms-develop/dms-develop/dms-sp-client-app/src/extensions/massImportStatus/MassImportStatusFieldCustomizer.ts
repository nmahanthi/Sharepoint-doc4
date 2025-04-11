import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import MassImportStatus, { IMassImportStatusProps } from './components/MassImportStatus';
import { BatchImportConstants, DmsDocumentFormConstants } from '../../constants';
import { AadHttpClient } from '@microsoft/sp-http';
import { INotificationService, NotificationService } from '../../service/NotificationService';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP } from '../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import { MassImportService } from './service/MassImportService';
import { ProjectPropertiesService } from '../../service/ProjectPropertiesService';
import { IProjectPropertiesService } from '../../service/IProjectPropertiesService';
/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMassImportStatusFieldCustomizerProperties {
  // This is an example; replace with your own property
  massImportEndpointUrl: string;
  dmlFilePrefix: string;
  midFilePrefix: string;
  miaFilePrefix: string;
  dmsUserApiAppId: string;
  maxQueueTime: number;
}

const LOG_SOURCE: string = 'MassImportStatusFieldCustomizer';

export default class MassImportStatusFieldCustomizer
  extends BaseFieldCustomizer<IMassImportStatusFieldCustomizerProperties> {
  private _dmsClient: AadHttpClient;
  private _notificationService: INotificationService;
  private _effectiveProperties: IMassImportStatusFieldCustomizerProperties;
  private _sp: SPFI;
  private _projectPropertiesService: IProjectPropertiesService;

  public async onInit(): Promise<void> {
    const { dmsUserApiAppId } = this.properties;
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated MassImportStatusFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    try {
      this._dmsClient = await this.context.aadHttpClientFactory.getClient(dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);
      const servicesInitPromise = new Promise<void>((resolve) => {
        this.context.serviceScope.whenFinished(() => {
          this._notificationService = this.context.serviceScope.consume(NotificationService.serviceKey);
          this._projectPropertiesService = this.context.serviceScope.consume(ProjectPropertiesService.serviceKey);
          resolve();
        });
      });
      await servicesInitPromise;
      this._sp = getSP(this.context);
      const valueFromProjectProperties = await this._projectPropertiesService.getProperty('MassImportStatusFieldCustomizerProperties');
      this._effectiveProperties = JSON.parse(valueFromProjectProperties || 'null') || this.properties;
      MassImportService.initialize(
        this._dmsClient, 
        this.properties.massImportEndpointUrl || BatchImportConstants.MassImportEndpointUrl,
        this.context.pageContext.web.serverRelativeUrl
      );
    } catch (e) {
      console.error(e);
      throw e;
    }
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const fileName = event.listItem.getValueByName('FileLeafRef');
    const status = event.listItem.getValueByName(BatchImportConstants.BatchImportListFields.Status);
    const massImportStatus: React.ReactElement<{}> =
      React.createElement(MassImportStatus, {
        fileName: fileName,
        status: status,
        dmsClient: this._dmsClient,
        notificationService: this._notificationService,
        dmlFilePrefix: this._effectiveProperties.dmlFilePrefix || BatchImportConstants.DmlFilePrefix,
        midFilePrefix: this._effectiveProperties.midFilePrefix || BatchImportConstants.MidFilePrefix,
        miaFilePrefix: this._effectiveProperties.miaFilePrefix || BatchImportConstants.MiaFilePrefix,
        fileServerRelativeUrl: event.listItem.getValueByName('FileRef'),
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        massImportEndpointUrl: this._effectiveProperties.massImportEndpointUrl || BatchImportConstants.MassImportEndpointUrl,
        spId: event.listItem.getValueByName('ID'),
        spListId: this.context.pageContext.list?.id.toString(),
        sp:this._sp,
        currentUser: this.context.pageContext.user.email,
        maxQueueTime: this.properties.maxQueueTime || 10,
      } as IMassImportStatusProps);

    ReactDOM.render(massImportStatus, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
