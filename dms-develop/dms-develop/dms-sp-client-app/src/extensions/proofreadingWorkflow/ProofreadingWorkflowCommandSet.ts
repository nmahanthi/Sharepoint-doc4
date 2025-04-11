import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  BaseListViewCommandSet,
  ListViewStateChangedEventArgs,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { override } from '@microsoft/decorators';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web"
import "@pnp/sp/site-users/web";
import "@pnp/sp/publishing-sitepageservice";
import ProofreadingDialogContent from '../../components/ProofReading/ProofreadingDialogContent';

import { DmsRole, IPermissionsService, PermissionsService } from '../../service/PermissionsService';
import { ContentTypeSiteIds, ListUrls } from '../../constants';
import { ISelectedItemService, SelectedItemService } from '../../service/SelectedItemService';
import { INotificationService, NotificationService } from '../../service/NotificationService';

export interface IProofreadingWorkflowCommandSetProperties {
  workflowUrl: string;
}

const LOG_SOURCE: string = 'ProofreadingWorkflowCommandSet';


export default class ProofreadingWorkflowCommandSet extends BaseListViewCommandSet<IProofreadingWorkflowCommandSetProperties> {
  private container: HTMLDivElement | null = null;
  private _permissionsService: IPermissionsService;
  private _selectedItemService: ISelectedItemService;
  private _notificationService: INotificationService;

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ProofreadingWorkflowCommandSet');

    this.context.serviceScope.whenFinished(() => {
      this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);
      this._selectedItemService = this.context.serviceScope.consume(SelectedItemService.serviceKey);
      this._notificationService = this.context.serviceScope.consume(NotificationService.serviceKey);
      this._selectedItemService.init(this);
    });
    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewStateChanged.bind(this));
    this._showDialog = this._showDialog.bind(this);
    return Promise.resolve();
  }

  @override
  public async onListViewStateChanged(event: ListViewStateChangedEventArgs): Promise<void> {
    const compareOneCommand: Command = this.tryGetCommand('PROOFREADINGCOMMAND');
    if (compareOneCommand) {
      compareOneCommand.visible = false;
      this.raiseOnChange();
      compareOneCommand.visible = await this.updateCommandVisibility(event);
      this.raiseOnChange();
    }
  }

  private async updateCommandVisibility(event: ListViewStateChangedEventArgs): Promise<boolean> {
    const selectedRows = this.context.listView.selectedRows;
    if (selectedRows?.length !== 1 ||
      !this.context.pageContext.list ||
      (this.context.listView.selectedRows && !this.context.listView.selectedRows[0].getValueByName('ID'))) {
      return false;
    }

    const data = await this._selectedItemService.getSelectedItemData();
    if (!data) {
      return false;
    }
    const hasContributorRole = await this._permissionsService.currentUserHasRole(DmsRole.Contributor);
    return data.docStatus === 'Draft' &&
      data.contentTypeId.startsWith(ContentTypeSiteIds.TechnicalDocuments) &&
      this.context.pageContext.list.serverRelativeUrl.indexOf(ListUrls.Draft) !== -1 &&
      hasContributorRole &&
      data.childItemsCount > 0;
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'PROOFREADINGCOMMAND':
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

        console.log("after raise change");
      }
    };

    const defaultWorkflowUrl = "https://prod-104.westeurope.logic.azure.com:443/workflows/d30aa75d1dfc40f18d4e51544189d2bf/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oJi2TWhsh1Syjh69gTmyImGuVZFxNVkJgaEKKoH3xVc";
    if (
      !this.context.listView.selectedRows ||
      this.context.listView.selectedRows.length === 0 ||
      !this.context.pageContext.list) {
      console.error('No selected rows');
      return;
    }
    const itemId = this.context.listView.selectedRows[0].getValueByName('ID') as number;

    this._selectedItemService.getSelectedItemData()
      .then(data => {
        if(!data) {
          console.error('No data');
          return;
        }
        const element = React.createElement(ProofreadingDialogContent, {
          onClose,
          context: this.context,
          itemID: itemId,
          libraryID: this.context.pageContext.list?.id.toString() ?? '',
          siteURL: this.context.pageContext.web.absoluteUrl,
          currentUser: data.userName,
          currentUserEmail: data.userEmail,
          docGUId: data.docIDextracted,
          createdBy: data.createdByobj,
          projReference: data.projReference,
          projRevision: data.projRevision,
          docTitle: data.docTitle,
          itemLink: data.itemLink,
          workflowUrl: this.properties.workflowUrl || defaultWorkflowUrl,
          notificationService: this._notificationService,
        });
        ReactDOM.render(element, this.container);
      })
      .catch(error => {
        console.error('Error fetching item data:', error);
      });
  }
}
