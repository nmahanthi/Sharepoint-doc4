import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'DmsApplicationActionsFieldCustomizerStrings';
import DmsApplicationActions, { IDmsApplicationActionsProps } from './components/DmsApplicationActions';
import { IPermissionsService, PermissionsService } from '../../service/PermissionsService';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../pnpjs-config';
import { IRequestService } from '../../service/IRequestService';
import { RequestService } from '../../service/RequestService';
import { INotificationService, NotificationService } from '../../service/NotificationService';
import { DmsDocumentFormConstants } from '../../constants';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDmsApplicationActionsFieldCustomizerProperties {
  // This is an example; replace with your own property
  workflowUrl?: string;
  AVVAWorkflowUrl?:string;
  proofreadingWorkflowUrl?:string;
  logHistoryWorkflowUrl?: string;
  dmsUserApiAppId:string;
}

const LOG_SOURCE: string = 'DmsApplicationActionsFieldCustomizer';

export default class DmsApplicationActionsFieldCustomizer
  extends BaseFieldCustomizer<IDmsApplicationActionsFieldCustomizerProperties> {

    private _permissionsService: IPermissionsService;
    private _requestService: IRequestService;
    private _notificationService: INotificationService;
    private _sp: SPFI;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Activated DmsApplicationActionsFieldCustomizer with properties:');
    console.log("DmsApplicationActionsFieldCustomizer Initialization");
    try{
      this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);
      this._requestService = this.context.serviceScope.consume(RequestService.serviceKey);
      this._notificationService = this.context.serviceScope.consume(NotificationService.serviceKey);
      this._sp = getSP(this.context);
    } catch (e) {
      console.error(e);
      throw e;
    }
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const dmsApplicationActions: React.ReactElement<{}> =
      React.createElement(DmsApplicationActions, { 
        sp: this._sp,
        currentItem: event.listItem,
        listId: this.context.pageContext?.list?.id.toString(),
        listUrl: this.context.pageContext?.list?.serverRelativeUrl.split('/').pop(),
        webRelativeUrl: this.context.pageContext?.web?.serverRelativeUrl,
        permissionService: this._permissionsService,
        strings,
        currentUser: this.context.pageContext.user.loginName,
        context: this.context,
        requestService: this._requestService,
        workflowUrl: this.properties.workflowUrl,
        notificationService: this._notificationService,
        AVVAWorkflowUrl: this.properties.AVVAWorkflowUrl,
        proofreadingWorkflowUrl: this.properties.proofreadingWorkflowUrl,
        logHistoryHttpTriggerEndPoint:  this.properties.logHistoryWorkflowUrl,
        dmsUserApiAppId: this.properties.dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId
      } as IDmsApplicationActionsProps);

    ReactDOM.render(dmsApplicationActions, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
