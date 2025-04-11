import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
} from '@microsoft/sp-listview-extensibility';
import { IMessageBarProps, MessageBar, MessageBarType } from '@fluentui/react';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/content-types";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web"
import "@pnp/sp/publishing-sitepageservice";

import { override } from '@microsoft/decorators';
import { getSP } from '../../pnpjs-config';
import { ContentTypeSiteIds } from '../../constants';
import { DmsRole, IPermissionsService, PermissionsService } from '../../service/PermissionsService';
import { FunctionComponentElement } from 'react';
import strings from 'DmsDocumentSetCreationCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDmsDocumentSetCreationCommandSetProperties {
  // This is an example; replace with your own properties
  targetContentTypeId: string;
  draftlistname: string;
}

const LOG_SOURCE: string = 'DmsDocumentSetCreationCommandSet';
const ERROR_MESSAGE_CONTAINER_ID = 'custom-error-message-container';
export default class DmsDocumentSetCreationCommandSet extends BaseListViewCommandSet<IDmsDocumentSetCreationCommandSetProperties> {
  private sp: SPFI;
  private _errorMessage: string = '';
  private _errorMessageElement: FunctionComponentElement<IMessageBarProps>;
  private _contentTypeId: string;
  private _permissionsService: IPermissionsService;

  private _getContentTypeId(): Promise<string | undefined> {
    if (!this._contentTypeId) {
      const listId = this.context.pageContext.list?.id.toString();
      if (listId) {
        return this.sp.web.lists.getById(listId).contentTypes.select('StringId')
          .filter(`startswith(StringId,'${ContentTypeSiteIds.TechnicalDocuments}')`)().then((contentTypes) => {
            if (contentTypes && contentTypes.length > 0) {
              this._contentTypeId = contentTypes[0].StringId;
              return this._contentTypeId;
            } else {
              return undefined;
            }
          });
      } else {
        throw new Error("No list id found in the context");
      }
    } else {
      return Promise.resolve(this._contentTypeId);
    }
  }
  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized DmsDocumentSetCreationCommandSet');
    let errorMessageContainer = document.getElementById(ERROR_MESSAGE_CONTAINER_ID);
    const libraryContainers = document.getElementsByClassName('od-scrollable-content od-scrollablePane-content-ItemsScopeList od-ItemsScopeList-content-sticky');
    if (!errorMessageContainer) {
      errorMessageContainer = document.createElement('div');
      errorMessageContainer.id = ERROR_MESSAGE_CONTAINER_ID;
      if(libraryContainers.length > 0){
        libraryContainers[0].appendChild(errorMessageContainer);
      }else{
        document.body.appendChild(errorMessageContainer);
      }
    }
    //Initialize PnP JS
    this.sp = getSP(this.context);
    this.context.serviceScope.whenFinished(() => {
      this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);
      this._getContentTypeId().catch((error) => {
        console.error("Error getting content type id: ", error);
        this._showErrorMessage(`${strings.InitializationError}. Details: ${error.message}`);
      });
    });
    // Check if Permissionconfig list contains items and update the visibility of Command 2
    this.updateCommandVisibility().catch((error) => {
      console.error("Error updating command visibility: ", error);
      this._showErrorMessage(`${strings.InitializationError}. Details: ${error.message}`);
    });

    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewStateChanged.bind(this));
    return Promise.resolve();
  }
  private async updateCommandVisibility(): Promise<void> {
    const { draftlistname } = this.properties;
    const command2 = this.tryGetCommand('COMMAND_2');
    if (command2) {
      command2.visible = false;
      this.raiseOnChange();
      const hasPermissions = await this._permissionsService.currentUserHasRole(DmsRole.DocumentController);
      const isDraftLibrary = this.context.pageContext.list?.serverRelativeUrl === `${this.context.pageContext.web.serverRelativeUrl}/${draftlistname}`;
      command2.visible = hasPermissions && isDraftLibrary;
      this.raiseOnChange();
    }
  }

  @override
  public onListViewStateChanged(event: ListViewStateChangedEventArgs): void {
    this._getContentTypeId().catch((error) => {
      console.error("Error getting content type id: ", error);
      this._showErrorMessage(`${strings.InitializationError}. Details: ${error.message}`);
    });
    this.updateCommandVisibility().catch((error) => {
      console.error("Error updating command visibility: ", error);
      this._showErrorMessage(`${strings.InitializationError}. Details: ${error.message}`);
    });
  }
  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    const listId = this.context.pageContext.list?.id.toString();
    const listabsUrl = this.context.pageContext.list?.serverRelativeUrl.toString();
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const ct = await this._getContentTypeId();
    if (!ct) {
      this._showErrorMessage(`${strings.InitializationError}. Details: Content type not found`);
      console.error("Content type not found");
      return;
    }
    const redirecturl = `${webUrl}/_layouts/15/SPListForm.aspx?PageType=8&List=${encodeURIComponent(listId || '').replace(/-/g, '%2D')}&Source=${encodeURIComponent(location.href)}&ContentTypeId=${encodeURIComponent(ct)}&RootFolder=${encodeURIComponent(listabsUrl || '')}`;//working
    switch (event.itemId) {
      case 'COMMAND_2':
        window.location.href = redirecturl;
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  private _showErrorMessage(message: string): void {
    this._errorMessage = message;
    this.render();
  }
  @override
  public render(): void {
    if (this._errorMessage) {
      this._errorMessageElement = React.createElement(MessageBar, {
        messageBarType: MessageBarType.error,
        isMultiline: false,
        onDismiss: () => { this._errorMessage = ''; this.render(); },
        dismissButtonAriaLabel: 'Close'
      }, this._errorMessage);

      const errorMessageContainer = document.getElementById(ERROR_MESSAGE_CONTAINER_ID);
      if (errorMessageContainer) {
        ReactDOM.render(this._errorMessageElement, errorMessageContainer);
      }
    }
  }
  protected onDispose(): void {
    const errorMessageContainer = document.getElementById(ERROR_MESSAGE_CONTAINER_ID);
    if (errorMessageContainer && this._errorMessageElement) {
      ReactDOM.unmountComponentAtNode(this._errorMessageElement as unknown as Element);
    }
  }
}

