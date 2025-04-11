/* eslint-disable no-case-declarations */
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
import WorkflowsPanel, { IWorkflowsPanelProps } from '../../components/WorkflowsPanel/WorkflowsPanel';
import { assign } from '@fluentui/react';
import { getSP } from '../../pnpjs-config';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDmsContexualMenuCommandSetCommandSetProperties {
  // This is an example; replace with your own properties 
  requestListId: string;
  taskListId: string;
  workflowUrl: string;
}

const LOG_SOURCE: string = 'DmsContexualMenuCommandSetCommandSet';

export default class DmsContexualMenuCommandSetCommandSet extends BaseListViewCommandSet<IDmsContexualMenuCommandSetCommandSetProperties> {
  private _sp: SPFI;
  private _requestService: IRequestService;
  private panelPlaceHolder: HTMLDivElement | null=null;
  public onInit(): Promise<void> {
    console.log('Initialized DmsContexualMenuCommandSetCommandSet');
    Log.info(LOG_SOURCE, 'Initialized DmsContexualMenuCommandSetCommandSet');

    //const workflowUrl=this.properties.workflowUrl || 'https://prod-150.westeurope.logic.azure.com:443/workflows/e2e86240560746d0a5d8191a09ed10ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=t-PsdqesaVR5HFrJKke4oqYiJo-G6a1vMOGZIvD9ceE';

    const openWFPanel: Command = this.tryGetCommand('OPEN_WFPANEL');
    openWFPanel.visible = true;//false;

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

  private _showPanel(itemId: number) {
    if(!this.context.pageContext || !this.context.listView)
    {
      return;
    }
    const workflowUrl=this.properties.workflowUrl || 'https://prod-150.westeurope.logic.azure.com:443/workflows/e2e86240560746d0a5d8191a09ed10ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=t-PsdqesaVR5HFrJKke4oqYiJo-G6a1vMOGZIvD9ceE';

    this._renderPanelComponent({
      showPanel: true,
      setShowPanel: this._dismissPanel.bind(this),//(val: boolean) => void;
      documentId: itemId,
      requestService: this._requestService,
      cultureName: this.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(), 
      currentUserLogin: this.context.pageContext.user.loginName,
      workflowUrl:workflowUrl
  //    onClose: this._dismissPanel
    });
  }

  private _renderPanelComponent(props: any) {
    if(!props.showPanel)
        {ReactDom.unmountComponentAtNode(this.panelPlaceHolder!);
          return;}
    const element: React.ReactElement<IWorkflowsPanelProps> = React.createElement(WorkflowsPanel, assign({
      showPanel: false,
      setShowPanel: null,
      documentId: null,
      requestService: this._requestService,
      cultureName: this.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(), 
      currentUserLogin: this.context.pageContext.user.loginName,
      isthreeSixtyDegree:false
    }, props));
    ReactDom.render(element, this.panelPlaceHolder!);
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> { 
    switch (event.itemId) {      
      case 'OPEN_WFPANEL':
      const listGuid=this.context.listView.list?.guid?.toString() || '';
        this._requestService.configure(
          listGuid,
          this._sp,
          this.context
        );
        //console.log(event);
        let selectedItem = event.selectedRows[0];
        const listItemId = selectedItem.getValueByName('ID') as number;
        this._showPanel(listItemId);


        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private isDocumentSet(selectedItem: RowAccessor)
  {
    if(selectedItem.getValueByName('ContentTypeId').startsWith('0x0120D520'))
    {
      return true
    }
    else
    {
      return false;
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const openWFPanel: Command = this.tryGetCommand('OPEN_WFPANEL');
    if (openWFPanel) {
      // This command should be hidden unless exactly one row is selected.
     // if (this.context.listView.selectedRows?.length === 1 && this.context.pageContext.list?.title.toLowerCase() === "draft") {
     if(this.context.listView.selectedRows?.length === 1 && this.isDocumentSet(this.context.listView.selectedRows[0]))  {
     openWFPanel.visible = true;
      }
      else {
        openWFPanel.visible = false;
      }
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
