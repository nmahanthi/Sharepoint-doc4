import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  RowAccessor,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { DmsRole, IPermissionsService, PermissionsService } from '../../service/PermissionsService';
import '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import ConfirmationPopup from '../../components/Applicable/ApplicableStatus';
import { getSP } from '../../pnpjs-config';

const LOG_SOURCE: string = 'ApplicableMenuCommandSet';

export default class ApplicableMenuCommandSet extends BaseListViewCommandSet<IApplicableMenuCommandSetStrings> {

  private _permissionsService: IPermissionsService;
  container: any;
  private fileLeafRefs:{projectReference:string;projectRevision:string}[]|undefined
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ApplicableMenuCommandSet');

    this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);

    const compareOneCommand: Command = this.tryGetCommand('APPLICABLECOMMAND');
    compareOneCommand.visible = false;
    getSP(this.context);
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'APPLICABLECOMMAND':
        this._showDialog();
        break;

      default:
        throw new Error('Unknown command');
    }
  }

  private _showDialog(): void {
    if (this.container) {
      ReactDOM.unmountComponentAtNode(this.container);
      document.body.removeChild(this.container);
      this.container = null;
    }
    this.container = document.createElement('div');
    document.body.appendChild(this.container);

    const onClose = () => {
      if (this.container) {
        ReactDOM.unmountComponentAtNode(this.container);
        document.body.removeChild(this.container);
        this.container = null;
        this.raiseOnChange();
      }
    };

    const onConfirm = () => {
      this.raiseOnChange();
      onClose();
    };
 this.fileLeafRefs =this.context.listView.selectedRows&& this.context.listView.selectedRows.map((row: any) => {
  return {
    projectReference:row._values?.get("ProjectReference") || "",
    projectRevision:row._values?.get("ProjectRevision") || "",
  }
}).filter(ref => ref.projectReference !== ""&&ref.projectReference!==""); 

    const element = React.createElement(ConfirmationPopup, {
      webUrl: this.context.pageContext.web.serverRelativeUrl,
      onClose,
      onConfirm,
      getSelection: this._getSelection.bind(this),
      getselectedId: this._getselectedId.bind(this),
      fileLeafRefs:this.fileLeafRefs
    });
    ReactDOM.render(element, this.container);
  }
  
  private _getSelection(): readonly RowAccessor[] | undefined{
    return this.context.listView.selectedRows;
  }

  private _getselectedId(): string | undefined {
    const listId = this.context.pageContext.list?.id;
    return listId ? listId.toString() : undefined;  
}

  private _onListViewStateChanged = async (args: ListViewStateChangedEventArgs): Promise<void> => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('APPLICABLECOMMAND');
    const hasDocControllPermissions = await this._permissionsService.currentUserHasRole(DmsRole.DocumentController);

    if (compareOneCommand && hasDocControllPermissions) {
      if (this.context.listView && Array.isArray(this.context.listView.selectedRows)) {
        const selectedRows = this.context.listView.selectedRows;

        const allDraft = selectedRows.every((row: any) => {
          const documentStatus = row._values.get("DocumentStatus");
          const itemChildCount=row._values.get("ItemChildCount");
          return documentStatus === 'Draft'&& itemChildCount> 0;
        });
        compareOneCommand.visible = selectedRows.length >= 1 && allDraft;
      }
    } else {
      compareOneCommand.visible = false;
    }

    this.raiseOnChange();
  }
  
}