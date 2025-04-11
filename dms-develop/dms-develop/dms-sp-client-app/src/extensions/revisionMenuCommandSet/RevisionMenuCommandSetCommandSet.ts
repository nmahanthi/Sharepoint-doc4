import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { IRequestService } from '../../service/IRequestService';
import { RequestService } from '../../service/RequestService';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { assign} from '@fluentui/react';
import { IItem } from "@pnp/sp/items/types";
import WorkflowsPanel, { IWorkflowsPanelProps } from '../../components/WorkflowsPanel/WorkflowsPanel';
import { getSP } from '../../pnpjs-config';
import { AadHttpClient } from '@microsoft/sp-http';
import { DmsDocumentFormConstants } from '../../constants';
import BaselineService from '../../service/BaselineService';

export interface IRevisionMenuCommandSetCommandSetProperties {
  requestListId: string;
  taskListId: string;
  workflowUrl: string;
  httpTriggerEndPoint: string;
  dmsUserApiAppId: string;
}

const LOG_SOURCE: string = 'RevisionMenuCommandSetCommandSet';

export default class RevisionMenuCommandSetCommandSet extends BaseListViewCommandSet<IRevisionMenuCommandSetCommandSetProperties> {
  private _sp: SPFI;
  private _requestService: IRequestService;
  private panelPlaceHolder: HTMLDivElement | null = null;
  private itemData: any;
  private errorMessage: string | null = null;
  private _dmsClient: AadHttpClient;
  

  public  async onInit(): Promise<void> {
    console.log('Initialized RevisionMenuCommandSetCommandSet');
    Log.info(LOG_SOURCE, 'Initialized RevisionMenuCommandSetCommandSet');
    const { dmsUserApiAppId} = this.properties;

    await BaselineService.init(this.context,dmsUserApiAppId);
    this._dmsClient = await this.context.aadHttpClientFactory.getClient(dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);

    const openWFPanel: Command = this.tryGetCommand('COMMAND_1');
    openWFPanel.visible = true;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._sp = getSP(this.context);
    this.context.serviceScope.whenFinished(() => {
      this._requestService = this.context.serviceScope.consume(RequestService.serviceKey);
    });

    // Create the container for our React component
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

    return Promise.resolve();
  }

  private _dismissPanel() {
    this._renderPanelComponent({ showPanel: false });
  }

  private async getDetails(projectReference: string): Promise<{ draftItems: IItem[], applicableItems: IItem[], previousVersionsItems: IItem[] }> {
    const draftItems = await this.getItems('Draft', projectReference);
    const applicableItems = await this.getItems('Applicable Documents', projectReference);
    const previousVersionsItems = await this.getItems('Previous Versions', projectReference);
    return { draftItems, applicableItems, previousVersionsItems };
  }

  private async getItems(listTitle: string, projectReference: string): Promise<IItem[]> {
    try {
      const items: IItem[] = await this._sp.web.lists.getByTitle(listTitle).items
        .select('Title', 'Modified', 'Author/Title', 'ProjectReference', 'ProjectRevision', 'FileRef')
        .expand('Author')
        .filter(`ProjectReference eq '${projectReference}'`)();
      return items;
    } catch (error) {
      console.error(`Error fetching items from ${listTitle}:`, error);
      this.errorMessage = `Error fetching items from ${listTitle}. Please try again later.`;
      return [];
    }
  }

  private async showRevisionDetails(): Promise<void> {
    const selectedRows = this.context.listView.selectedRows;
    if (!selectedRows || selectedRows.length === 0) {
      return;
    }

    const selectedItem = selectedRows[0];
    const projectReference = selectedItem.getValueByName('ProjectReference');

    if (!projectReference) {
      this.errorMessage = 'No Project Reference found for the selected item.';
      this._renderPanelComponent({ showPanel: true });
      return;
    }

    try {
      const details = await this.getDetails(projectReference);

      this.itemData = [
        ...details.draftItems,
        ...details.applicableItems,
        ...details.previousVersionsItems
      ].map(item => ({
        title: (item as any).Title,
        modified: new Date((item as any).Modified).toLocaleDateString('en-GB'),
        createdBy: (item as any).Author ? (item as any).Author.Title : 'Unknown',
        projectRevision: (item as any).ProjectRevision,
        link: (item as any).FileRef
      }));

      if (this.itemData && this.itemData.length > 0) {
        const listItemId = selectedItem.getValueByName('ID') as number;
        const docId = selectedItem.getValueByName('_dlc_DocIdUrl.desc');
        this._showPanel(listItemId,docId);
      } else {
        this.errorMessage = 'No document data available for the selected project.';
        this._renderPanelComponent({ showPanel: true });
      }
    } catch (error) {
      console.error('Error fetching document details:', error);
      this.errorMessage = 'An error occurred while retrieving document details. Please try again later.';
      this._renderPanelComponent({ showPanel: true });
    }
  }

  private _showPanel(itemId: number, docId: string): void {
    try {
      if (!this.context.pageContext || !this.context.listView || !this.itemData) {
        this.errorMessage = 'An error occurred while preparing the panel. Please try again.';
        this._renderPanelComponent({ showPanel: true });
        return;
      }
      const httpTriggerEndPoint = this.properties.httpTriggerEndPoint || "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/documentChangeHttpTrigger";
      const workflowUrl = this.properties.workflowUrl || 'https://prod-150.westeurope.logic.azure.com:443/workflows/e2e86240560746d0a5d8191a09ed10ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=t-PsdqesaVR5HFrJKke4oqYiJo-G6a1vMOGZIvD9ceE';

      this._renderPanelComponent({
        showPanel: true,
        setShowPanel: this._dismissPanel.bind(this),
        documentId: itemId,
        requestService: this._requestService,
        cultureName: this.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
        currentUserLogin: this.context.pageContext.user.loginName,
        workflowUrl: workflowUrl,
        itemData: this.itemData,
        context: this.context,
        dmsClient: this._dmsClient,
        httpTriggerEndPoint: `${httpTriggerEndPoint}?documentId=${encodeURIComponent(docId)}`
      });
    } catch (error) {
      console.error('Error showing panel:', error);
      this.errorMessage = 'An error occurred while displaying the panel. Please try again.';
      this._renderPanelComponent({ showPanel: true });
    }
  }

  private _renderPanelComponent(props: any) {
    try {
      if (!props.showPanel) {
        ReactDom.unmountComponentAtNode(this.panelPlaceHolder!);
        return;
      }
      const element: React.ReactElement<IWorkflowsPanelProps> = React.createElement(WorkflowsPanel, assign({
        showPanel: false,
        setShowPanel: null,
        documentId: null,
        requestService: this._requestService,
        cultureName: this.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
        currentUserLogin: this.context.pageContext.user.loginName,
        isthreeSixtyDegree: true,
        itemData: this.itemData,
        errorMessage: this.errorMessage // Pass the error message to the panel component
      }, props));
      ReactDom.render(element, this.panelPlaceHolder!);
    } catch (error) {
      console.error('Error rendering panel component:', error);
      this.errorMessage = 'An error occurred while rendering the panel. Please try again.';
      this._renderPanelComponent({ showPanel: true });
    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    try {
      const listGuid = this.context.listView.list?.guid?.toString() || '';
      switch (event.itemId) {
        case 'COMMAND_1':
          this._requestService.configure(listGuid, this._sp, this.context);
          await this.showRevisionDetails();
          break;
        default:
          throw new Error('Unknown command');
      }
    } catch (error) {
      console.error('Error executing command:', error);
      this.errorMessage = 'An error occurred while executing the command. Please try again.';
      this._renderPanelComponent({ showPanel: true });
    }
  }

  private isDocumentSet(selectedItem: RowAccessor) {
    return selectedItem.getValueByName('ContentTypeId').startsWith('0x0120D520');
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const openWFPanel: Command = this.tryGetCommand('COMMAND_1');
    if (openWFPanel) {
      if (this.context.listView.selectedRows?.length === 1 && this.isDocumentSet(this.context.listView.selectedRows[0])) {
        openWFPanel.visible = true;
      } else {
        openWFPanel.visible = false;
      }
    }

    this.raiseOnChange();
  }
  
}
