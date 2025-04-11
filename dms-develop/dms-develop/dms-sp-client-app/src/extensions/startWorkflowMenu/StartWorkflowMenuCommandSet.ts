import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  ListViewStateChangedEventArgs,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import WorkflowDialogContent from '../../components/Workflow/WorkflowDialogContent';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/graph/groups"; // Import for groups
import "@pnp/graph/groups";
import "@pnp/graph/users";
import "@pnp/sp/publishing-sitepageservice";
import { override } from '@microsoft/decorators';
import { DmsRole, IPermissionsService, PermissionsService } from '../../service/PermissionsService';
import { ISelectedItemService, SelectedItemService } from '../../service/SelectedItemService';
import { ContentTypeSiteIds, ListUrls } from '../../constants';
import { INotificationService, NotificationService } from '../../service/NotificationService';

export interface IStartWorkflowMenuCommandSetProperties {
  workflowUrl: string;
}

const LOG_SOURCE: string = 'StartWorkflowMenuCommandSet';

export default class StartWorkflowMenuCommandSet extends BaseListViewCommandSet<IStartWorkflowMenuCommandSetProperties> {
  private container: HTMLDivElement | null = null;
  private _permissionsService: IPermissionsService;
  private _selectedItemService: ISelectedItemService;
  private _notificationService: INotificationService;


  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized StartWorkflowMenuCommandSet');
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
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      compareOneCommand.visible = false;
      this.raiseOnChange();
      const visibility = await this.updateCommandVisibility(event);
      compareOneCommand.visible = visibility;
      this.raiseOnChange();  // Trigger a re-render of the command bar
    }

  }

  private async updateCommandVisibility(event: ListViewStateChangedEventArgs): Promise<boolean> {
    console.log("Start AVVA workflow");
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
      case 'COMMAND_1':
        this._showDialog();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  private _showDialog(): void {
    if (
      !this.context.listView.selectedRows ||
      this.context.listView.selectedRows.length === 0 ||
      !this.context.pageContext.list) {
      console.error('No selected rows');
      return;
    }
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
      }
    };
    const defaultWorkflowUrl = 'https://prod-44.westeurope.logic.azure.com:443/workflows/b5d5e9e2f34f476fbd47a205c836d9d0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=z_4_ooxtm1cBxgazpBZ72HQ3mIeYLiQb-EtsQFj8CX0';
       const itemId = this.context.listView.selectedRows[0].getValueByName('ID') as number;
    this._selectedItemService.getSelectedItemData().then((data) => {
      if (!data) {
        console.error('No data');
        return;
      }
      const element = React.createElement(WorkflowDialogContent, {
        onClose, context: this.context,
        itemID: itemId,
        libraryID: this.context.pageContext.list?.id.toString() ?? '',
        siteURL: this.context.pageContext.web.absoluteUrl,
        currentUser: data.userName,
        currentUserEmail: data.userEmail,
        projReference: data.projReference,
        projRevision: data.projRevision,
        docTitle: data.docTitle,
        itemLink: data.itemLink,
        docGUId: data.docIDextracted,// Passing the URL as a property
        createdBy: data.createdByobj,
        workflowUrl: this.properties.workflowUrl || defaultWorkflowUrl,
        notificationService: this._notificationService,
      });
      ReactDOM.render(element, this.container);
    }).catch((error) => {
      console.error('Error fetching item data:', error);
    });
  }
}
