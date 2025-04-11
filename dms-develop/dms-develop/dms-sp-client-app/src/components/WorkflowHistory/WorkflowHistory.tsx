import { Log } from "@microsoft/sp-core-library";
import * as React from "react";
import styles from './WorkflowHistory.module.scss';
import { IDraftRequest } from "../../interfaces/IDraftRequest";
import * as strings from "DmsContexualMenuCommandSetCommandSetStrings";
import { Modal } from "@fluentui/react/lib/Modal";
import { IconButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { DetailsList, DetailsListLayoutMode, DetailsRow, IColumn, IDetailsRowProps, SelectionMode } from "@fluentui/react/lib/DetailsList";
//import { TextField } from "@fluentui/react/lib/TextField";
//import { Label } from "@fluentui/react/lib/Label";
//import WorkflowTimeline, { IWorkflowTimelineStep } from "../WorkflowTimeline/WorkflowTimeline";
import WorkflowTimeline from "../WorkflowTimeline/WorkflowTimeline";
import { IWfTask } from "../../interfaces/IWfTask";
import { IRequestService } from "../../service/IRequestService";
import { IHistoryTimeLine } from "../../interfaces/IGlobalInterfaces";

const LOG_SOURCE: string = 'WorkflowHistory';


export interface IWorkflowHistoryProps {
  isLoading?: boolean;
  isError?: boolean;
  requests: IDraftRequest[];
  isOpen: boolean;
  containerClassName?: string;
  onDismiss: () => void;
  title: string;
  cultureName: string;
  requestService: IRequestService;
  timeLineTasks:IWfTask[]
  historyData?:IHistoryTimeLine[];
}
export interface IWorkflowHistoryState {
  requests?: (IHistoryTimeLine & { showTimeline?: boolean; isTasksLoading?: boolean; isTasksError?: boolean; loadedTasks?: IWfTask[] })[];
}
export default class WorkflowHistory extends React.Component<IWorkflowHistoryProps, IWorkflowHistoryState> {
  private _columns: IColumn[];

  constructor(props: IWorkflowHistoryProps) {
    super(props);
    this.state = { requests: props.historyData || [] };
    this._columns = [
      {
        key: 'expandColumn',
        name: 'Details',
        // className: classNames.fileIconCell,
        // iconClassName: classNames.fileIconHeaderIcon,
        iconName: 'FullHistory',
        isIconOnly: true,
        fieldName: 'Details',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IHistoryTimeLine & { showTimeline?: boolean, isTasksLoading: boolean, isTasksError: boolean }) => (
          <IconButton iconProps={{ iconName: item.showTimeline ? 'CollapseContentSingle' : 'ExploreContentSingle' }} onClick={() => {
            const newRequests = [...(this.state.requests || [])];
            const newRequest = newRequests.find(r=>r.id === item.id);
            if(newRequest){
              newRequest.showTimeline = !newRequest.showTimeline;
              this.setState({
                requests: newRequests
              });
            }
           
          }} />
        ),
      },
      {
        key: 'requestorColumn',
        name: strings.WfHistoryColumnHeaderRequestor,
        fieldName: 'requestor',
        minWidth: 80,
        maxWidth: 200,
        isResizable: true,
        isCollapsible: true,
        onRender: (item: IHistoryTimeLine) => {
          return <span>{item.requestor}</span>;
        }
      },
      {
        key: 'createdColumn',
        name: strings.WfHistoryColumnHeaderCreated,
        fieldName: 'created',
        minWidth: 80,
        maxWidth: 90,
        isResizable: true,
        onRender: (item: IHistoryTimeLine) => {
          return <span>{new Date(item.created).toLocaleDateString(this.props.cultureName, { year: "numeric", month: "numeric", day: "numeric" })}</span>;
        },
        isPadded: true
      },
      {
        key: 'modifiedColumn',
        name: strings.WfHistoryColumnHeaderLastUpdated,
        fieldName: 'modified',
        minWidth: 80,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        onRender: (item: IHistoryTimeLine) => {
          return <span>{new Date(item.modified).toLocaleDateString(this.props.cultureName, { year: "numeric", month: "numeric", day: "numeric" })}</span>;
        },
        isPadded: true
      },
      {
        key: 'statusColumn',
        name: strings.WfHistoryColumnHeaderStatus,
        fieldName: 'workflowstatus',
        minWidth: 80,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: false,
        onRender: (item: IHistoryTimeLine) => {
          return <span>{item.workflowstatus}</span>;
        },
        isPadded: true
      }      
    ];
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: WorkflowHistory mounted');
    const { historyData } = this.props;
    this.setState({ requests: historyData });
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: WorkflowHistory unmounted');
  }

  public render(): React.ReactElement<{}> {
    //const { isOpen, onDismiss, containerClassName, title, timeLineTasks } = this.props;
    const { isOpen, onDismiss, containerClassName, title } = this.props;
    const { requests } = this.state;
    return <Modal
      isBlocking={false} 
      isOpen={isOpen}
      onDismiss={onDismiss}
      containerClassName={containerClassName}
    >
      <div className={styles.header}>
        <h2 className={styles.heading}>
          {title} Workflows
        </h2>
        <IconButton
          className={styles.CloseButton}
          iconProps={{ iconName: 'Cancel' }}
          onClick={onDismiss}
        />
      </div>
      <Stack className={styles.body}>

        {requests && <DetailsList
          items={requests}
          columns={this._columns}
          selectionMode={SelectionMode.none}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onItemInvoked={this._onItemInvoked}
          onRenderRow={this._onRenderRow.bind(this)}
        />}
      </Stack>
    </Modal>;
  }

  private _getKey(item: IDraftRequest, index?: number): string {
    return '' + item.id;
  }

  private _onItemInvoked(item: IHistoryTimeLine & { showTimeline?: boolean, isTasksLoading?: boolean, isTasksError?: boolean, loadedTasks?: IWfTask[] }): void {
    const newRequests = [...(this.state.requests || [])];
    const newRequest = newRequests.find(r=>r.id === item.id);
    if(newRequest)
    {
      newRequest.showTimeline = !newRequest.showTimeline;
      this.setState({
        requests: newRequests
      });
    }
    
  }

  private _onRenderRow(props: IDetailsRowProps) {
    //const {  cultureName, requestService } = this.props;
    //const { requests } = this.state;
    if (props) {
      const request: (IHistoryTimeLine & { showTimeline?: boolean; isTasksLoading?: boolean; isTasksError?: boolean; loadedTasks?: IWfTask[] }) = props.item;
     /* if (request.showTimeline && !request.isTasksLoading && !request.isTasksError && !request.loadedTasks) {
        const newRequests = [...requests];
        newRequests.find(r=>r.id === request.id).isTasksLoading = true;
        this.setState({
          requests: newRequests
        });
        requestService.getRequestTasks(request.id).then(tasks => {
          const newRequests = [...requests];
          const newRequest = newRequests.find(r=>r.id === request.id);
          newRequest.isTasksLoading = false;
          newRequest.loadedTasks = tasks;
          this.setState({
            requests: newRequests
          });
        }).catch(err => {
          const newRequests = [...requests];
          const newRequest = newRequests.find(r=>r.id === request.id);
          newRequest.isTasksLoading = false;
          newRequest.isTasksError = true;
          this.setState({
            requests: newRequests
          });
          this.setState({
            requests: newRequests
          });
          Log.error(LOG_SOURCE, err);
          console.error(err);
        });
      }
        <div style={request.showTimeline ? undefined : { visibility: "hidden" }}>

    */
      return <React.Fragment>
        <DetailsRow {...props} />
        <div style={request.showTimeline ? undefined : { display: "none" }}>
          <WorkflowTimeline steps={request.steps} graphHeight={15} isLoading={request.isTasksLoading} isError={request.isTasksError} />
        </div>
      </React.Fragment>;
    }
    return null;
  }

}