/* eslint-disable @typescript-eslint/no-empty-function */
import { Log } from "@microsoft/sp-core-library";
import * as React from "react";
import styles from "./WorkflowsPanel.module.scss";
import { Panel } from "@fluentui/react/lib/Panel";
import * as strings from "DmsContexualMenuCommandSetCommandSetStrings";
import WorkflowTimeline, {
  IWorkflowTimelineStep,
} from "../WorkflowTimeline/WorkflowTimeline";
import { PanelType } from "@fluentui/react/lib/Panel";
import { Text } from "@fluentui/react/lib/Text";
import { Stack } from "@fluentui/react/lib/Stack";
import { Link } from "@fluentui/react/lib/Link";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { IRequestService } from "../../service/IRequestService";
import { IWfTask } from "../../interfaces/IWfTask";
import {
  Dialog,
  DialogFooter,
  DialogType,
  IDialogContentProps,
} from "@fluentui/react/lib/Dialog";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import CONSTANTS from "../../_Constants";
import { TextField } from "@fluentui/react/lib/TextField";
import { IDraftRequest } from "../../interfaces/IDraftRequest";
import WorkflowHistory from "../WorkflowHistory/WorkflowHistory";
import { IWfStatus } from "../../interfaces/IWfStatus";
import {
  ICancelWFDetails,
  IHistoryTimeLine,
} from "../../interfaces/IGlobalInterfaces";
import RevisionDetailsSet from "../RevisionMenu/RevisionDetailsSet";
import { MessageBar, MessageBarType } from "@fluentui/react";
import LogHistory from "../RevisionMenu/LogHistory";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { AadHttpClient } from "@microsoft/sp-http";
const LOG_SOURCE: string = "WorkflowsPanel";

export interface IWorkflowsPanelProps {
  showPanel: boolean;
  setShowPanel: () => void; //: (val: boolean) => void;
  documentId: number;
  requestService: IRequestService;
  cultureName: string;
  currentUserLogin: string;
  workflowUrl: string;
  isthreeSixtyDegree: boolean;
  itemData: any;
  errorMessage?: string;
  context: ListViewCommandSetContext,
  dmsClient: AadHttpClient,
  httpTriggerEndPoint: string
}
export interface IWorkflowsPanelState {
  hasError?: boolean;
  isLoading?: boolean;
  requests?: IDraftRequest;
  disableCancelButton?: boolean;
  isValidUser?: boolean;
  wfCancelDetails?: ICancelWFDetails[];
  timeLineSections?: any[];
  showProofReadingWfHistory?: boolean;
  showAvvaWfHistory?: boolean;
  showHistory: any;
  showConfirmCancelation?: boolean;
  isCanceling?: boolean;
  isCancelingError?: boolean;
  isCancelingSuccess?: boolean;
  cancelationReason?: string;
  isCheckDocument:boolean;
}

export interface ITimeLineSection {
  title: string;
  steps: IWorkflowTimelineStep[];
  isError: boolean;
  requests: IDraftRequest;
  requestOutcome: string;
  timeLineTasks: IWfTask[];
  historyData: IHistoryTimeLine[];
}
export default class WorkflowsPanel extends React.Component<
  IWorkflowsPanelProps,
  IWorkflowsPanelState
> {
  /**
   *
   */
  constructor(props: IWorkflowsPanelProps) {
    super(props);
    this.state = {
      isLoading: true,
      disableCancelButton: true,
      isCheckDocument:true,
      timeLineSections: [],
      showHistory: {
        [strings.ProofReadingTitle]: false,
        [strings.AvvaTitle]: false,
      },
    };
  }

  public async componentDidMount() {
    await this.loadPanel();
  }

  private async loadPanel() {
    Log.info(LOG_SOURCE, "React Element: WorkflowsPanel mounted");
    const { requestService, documentId, cultureName, currentUserLogin } =
      this.props;
    requestService
      .getDraftDocumentRequest(documentId)
      .then((requests) => {
        //Check Document status is application then disable cancel workflow button
        const pendingWF = requests?.wfStatus.filter(
          (wf) => wf.workflowStatus.toLowerCase() === "in progress"
        );
        this.setState({
          disableCancelButton:
            requests?.documentstatus?.toLowerCase() === "applicable" ||
            pendingWF.length === 0
              ? true
              : false,
        });
        this.setState({
          isCheckDocument:requests?.documentstatus?.toLowerCase() === "applicable" || requests?.documentstatus?.toLowerCase() === "draft" ? true: false,
      });
        //Format tasks for any delegated tasks
        let orderedTasks: IWfTask[] = [];
        requests.tasks.map((t) => {
          const delegTasks = requests.tasks.filter(
            (nxTsk) =>
              t.id === parseInt(nxTsk.prevTaskId) && nxTsk.isDelegatedTask
          );
          if (delegTasks.length === 0)
            this.checkExists(orderedTasks, t.id) ? "" : orderedTasks.push(t);
          else {
            t.outcome = "Delegated";
            orderedTasks.push(t);
            let delTask = delegTasks[0];
            delTask.delegator = t.assignee?.displayName;
            orderedTasks.push(delTask);
          }
        });

        let wfCancelDetails: ICancelWFDetails[] = [];
        let avvaHistoryTimeLine: IHistoryTimeLine[] = [];
        let prHistoryTimeLine: IHistoryTimeLine[] = [];
        //Get AVVA tasks
        const avvaTasks = orderedTasks.filter(
          (t) => t.taskAssigneeType !== CONSTANTS.ProofReadingStatusName
        );
        let avvSteps: IWfTask[] = [];

        let wfstatus = requests.wfStatus.filter(
          (wft) =>
            wft.workflowType === "AVVA" &&
            wft.revisionId === requests.documentId
        );
        if (wfstatus.length > 0) {
          const currInstance = wfstatus[0].instanceId;
          avvSteps = avvaTasks.filter((avt) => avt.instanceId === currInstance);

          //Collect information for WF cancellation for AVVA
          if (wfstatus[0].workflowStatus.toLowerCase() === "in progress") {
            wfCancelDetails.push({
              wfID: wfstatus[0].id,
              //taskIds: avvSteps.filter(p => p.outcome.toLowerCase() === "in progress").map(t => t.id)
              taskIds: avvSteps.map((t) => t.id),
            });
          }

          wfstatus.shift();
          avvaHistoryTimeLine = this.processHistoryTimeLineData(
            wfstatus,
            avvaTasks,
            cultureName,
            "Document Approved",
            requests
          );
        }
        //Get Proffreading Tasks
        const proofReadTasks = orderedTasks.filter(
          (t) => t.taskAssigneeType === CONSTANTS.ProofReadingStatusName
        );
        let prSteps: IWfTask[] = [];

        wfstatus = requests.wfStatus.filter(
          (wft) =>
            wft.workflowType === "Proofreading" &&
            wft.revisionId === requests.documentId
        );
        if (wfstatus.length > 0) {
          const currInstance = wfstatus[0].instanceId;
          prSteps = proofReadTasks.filter(
            (avt) => avt.instanceId === currInstance
          );

          //Collect information for WF cancellation for Proofreading
          if (wfstatus[0].workflowStatus.toLowerCase() === "in progress") {
            wfCancelDetails.push({
              wfID: wfstatus[0].id,
              // taskIds: prSteps.filter(p => p.outcome.toLowerCase() === "in progress").map(t => t.id)
              taskIds: prSteps.map((t) => t.id),
            });
          }

          wfstatus.shift();
          prHistoryTimeLine = this.processHistoryTimeLineData(
            wfstatus,
            proofReadTasks,
            cultureName,
            "Document Reviewed",
            requests
          );
        }

        let sections: ITimeLineSection[] = [];

        // AVVA section TBD
        sections.push({
          title: strings.AvvaTitle,
          //steps: this._tasksToSteps(requests, avvSteps, cultureName, "Submitted", "Draft ready"),
          steps:
            avvaTasks.length > 0
              ? this._tasksToSteps(
                  requests,
                  avvSteps,
                  cultureName,
                  "Submitted",
                  "Document Approved"
                )
              : [],
          isError: false,
          requests: requests,
          requestOutcome: "",
          timeLineTasks: avvaTasks,
          historyData: avvaHistoryTimeLine,
        });

        //show Proofreading section in timeline only is there any proofreadin review tasks created
        //    if (proofReadTasks.length > 0) {
        sections.push({
          title: strings.ProofReadingTitle,
          steps:
            proofReadTasks.length > 0
              ? this._tasksToSteps(
                  requests,
                  prSteps,
                  cultureName,
                  "Submitted",
                  "Document Reviewed"
                )
              : [],
          isError: false,
          requests: requests,
          requestOutcome: "",
          timeLineTasks: proofReadTasks,
          historyData: prHistoryTimeLine,
        });
        //   }

        this.setState({
          timeLineSections: sections,
          isLoading: false,
          wfCancelDetails: wfCancelDetails,
          requests,
        });
      })
      .catch((err) => {
        Log.error(LOG_SOURCE, err);
        console.error(err);
        this.setState({ isLoading: false, hasError: true });
      });

    this.setState({
      isValidUser: await requestService.canCurrentUserCancel(
        documentId,
        currentUserLogin
      ),
    });
  }

  private processHistoryTimeLineData(
    wfstatus: IWfStatus[],
    wfTasks: IWfTask[],
    cultureName: string,
    lastStepName: string,
    request: IDraftRequest
  ) {
    let historyData: IHistoryTimeLine[] = [];
    wfstatus.map((wfstat) => {
      historyData.push({
        created: wfstat.startDate,
        modified: wfstat.modified,
        workflowstatus: wfstat.workflowStatus,
        requestor: wfstat.requestor, //request?.author?.displayName,
        steps: this._tasksToSteps(
          request,
          wfTasks.filter((wft) => wft.instanceId === wfstat.instanceId),
          cultureName,
          "Submitted",
          lastStepName
        ),
        tasks: wfTasks,
        id: wfstat.id,
      });
    });
    return historyData;
  }

  private checkExists(tasks: IWfTask[], id: number) {
    return tasks.findIndex((task) => task.id === id) >= 0 ? true : false;
  }

  private _tasksToSteps(
    request: IDraftRequest,
    tasks: IWfTask[],
    cultureName: string,
    firstStepName: string,
    lastStepName: string
  ): IWorkflowTimelineStep[] {
    if (!request) {
      return [];
    }

    //get Initiator name
    let initiatorName = "";
    let wfInst= null;
 
    if (tasks.length > 0) {
      // const wfInst = request.wfStatus.filter(wf => wf.instanceId === tasks[0].instanceId);
      wfInst = request.wfStatus.filter(wf => wf.instanceId === tasks[0].instanceId);
      if (wfInst.length > 0)
        initiatorName = wfInst[0].requestor;
    }


    const pendingTasks = tasks.filter(
      (upd) => upd.outcome !== "Delegated" && upd.outcome === "In Progress"
    );
    const CancellledTasks = tasks.filter(
      (upd) => upd.outcome === "Cancelled" || upd.outcome === "Rejected"
    );
    return [
      // { name: firstStepName, person: request.author.displayName, date: new Date(request.created).toLocaleDateString(cultureName, { year: "numeric", month: "numeric", day: "numeric" }), status: 'completed' },
      {
        name: firstStepName,
       // person: request.author.displayName,
       person:initiatorName,
        // date: new Date(request.created).toLocaleDateString(cultureName, {
          date: new Date(wfInst?wfInst[0].startDate:new Date()).toLocaleDateString(cultureName, {
          year: "numeric",
          month: "numeric",
          day: "numeric",
        }),
        status:
          CancellledTasks.length === tasks.length ? "cancelled" : "completed",
      },

      ...tasks.map((task) => {
        return {
          name: task.taskAssigneeType,
          comments: task.comments,
          date:
            task.outcome === "Completed" ||
            task.outcome === "Delegated" ||
            task.outcome === "Cancelled"
              ? "" +
                new Date(task.endDate).toLocaleDateString(cultureName, {
                  year: "numeric",
                  month: "numeric",
                  day: "numeric",
                })
              : undefined,
          person: task.assignee?.displayName,
          delegator: task.delegator,
          reminderUrl: `mailto:${task.assignee?.eMail}`,
          status: task.outcome?.toLocaleLowerCase(),
          taskId: task.id,
        };
      }),
      {
        name: lastStepName,
        status:
          pendingTasks.length === 0 &&
          tasks.length > 0 &&
          CancellledTasks.length === 0
            ? "completed"
            : "",
      },
    ] as IWorkflowTimelineStep[];
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, "React Element: WorkflowsPanel unmounted");
  }

  public render(): React.ReactElement<{}> {
    const { showPanel, setShowPanel, cultureName, requestService } = this.props;
    const {
      hasError,
      timeLineSections,
      showConfirmCancelation,
      isCanceling,
      isCancelingError,
      isCancelingSuccess,
      cancelationReason,
      requests,
    } = this.state;

    const dialogContentProps: IDialogContentProps = {
      type: DialogType.normal,
      title: strings.WfPanelConfirmCancelationTitle,
      closeButtonAriaLabel: "Close",
      subText: isCanceling
        ? strings.WfPanelConfirmCancelationMessageIsLoading
        : isCancelingError
        ? strings.WfPanelConfirmCancelationMessageError
        : isCancelingSuccess
        ? strings.WfPanelConfirmCancelationMessageSuccess
        : strings.WfPanelConfirmCancelationMessage,
    };

    return (
      <div className={styles.WorkflowsPanel} id="amswfpanelcontainer">
        <div>
              {this.props.errorMessage && (
                <MessageBar
                  messageBarType={MessageBarType.error}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >
                  {this.props.errorMessage}
                </MessageBar>
              )}
            </div>
        {timeLineSections && (
          <Panel
            isLightDismiss
            isOpen={showPanel}
            onDismiss={() => {
              setShowPanel();
            }}
            closeButtonAriaLabel="Close"
            headerText={
              this.props.isthreeSixtyDegree
                ? strings.ThreeSixtyTitle
                : strings.WfPanelTitle
            }
            type={PanelType.medium}
            className={styles.Panel}
          >
            {this.props.isthreeSixtyDegree && (
              <RevisionDetailsSet
                items={this.props.itemData}
                onCancel={function (): void {
                  throw new Error("Function not implemented.");
                }}
              />
            )}
             {this.props.isthreeSixtyDegree && <div><Text  variant="large" as="h3" className={styles.headerContainer}>Workflow(s)</Text></div>}
            <Dialog
              hidden={!showConfirmCancelation}
              dialogContentProps={dialogContentProps}
              modalProps={{
                styles: { main: { maxWidth: 450 } },
                isBlocking: true,
              }}
              onDismiss={() => {
                if (isCancelingSuccess) {
                  setShowPanel();
                } else {
                  this.setState({
                    showConfirmCancelation: false,
                    isCancelingError: false,
                    isCancelingSuccess: false,
                  });
                }
              }}
            >
              {isCanceling && <Spinner />}
              {!isCanceling && !isCancelingError && !isCancelingSuccess && (
                <TextField
                  label={strings.WfPanelCancelationReason}
                  value={cancelationReason}
                  required
                  multiline
                  onChange={(
                    _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                    newValue?: string
                  ) => {
                    this.setState({ cancelationReason: newValue });
                  }}
                />
              )}
              <DialogFooter>
                {!isCanceling && !isCancelingError && !isCancelingSuccess && (
                  <PrimaryButton
                    onClick={() => {
                      this._cancelRequest();
                    }}
                    text={strings.WfPanelConfirmCancelationYes}
                    disabled={!cancelationReason}
                  />
                )}
                {!isCanceling && (
                  <DefaultButton
                    onClick={() => {
                      if (isCancelingSuccess) {
                        this.props.setShowPanel();
                      } else {
                        this.setState({
                          showConfirmCancelation: false,
                          isCancelingError: false,
                          isCancelingSuccess: false,
                        });
                      }
                    }}
                    text={strings.WfPanelConfirmCancelationNo}
                  />
                )}
              </DialogFooter>
            </Dialog>

            {hasError && (
              <Stack horizontalAlign="center" grow>
                <Text variant="small" className={styles.ErrorMessage}>
                  {strings.WfTimelineErrorMessage}
                </Text>
              </Stack>
            )}
            <Stack tokens={{ childrenGap: 10 }}>
              {timeLineSections.map((section, sectionIndex) => {
                return (
                  <React.Fragment>
                    <Stack.Item>
                      <Stack
                        horizontal
                        tokens={{ childrenGap: 20 }}
                        horizontalAlign="start"
                        verticalAlign="baseline"
                      >
                        <Text variant="large" as="h3">
                          <b>{section.title}</b>
                        </Text>
                      </Stack>
                    </Stack.Item>
                    {section.requests && section.steps.length > 0 && (
                      <Stack.Item>
                        <Text
                          block
                          variant="smallPlus"
                          className={styles.TimelineTitle}
                        >
                          {strings.TimelineTitle}
                        </Text>
                      </Stack.Item>
                    )}
                    <Stack.Item>
                      <Stack horizontal style={{ width: "100%" }}>
                        <Stack.Item grow>
                          <WorkflowTimeline
                            requestOutcome={section.requestOutcome}
                            steps={section.steps}
                            graphHeight={30}
                            isLoading={!section.steps}
                            isError={section.isError}
                            timeLineType={section.title}
                          />
                        </Stack.Item>
                      </Stack>
                    </Stack.Item>
                    {section.historyData.length > 0 && (
                      <Stack.Item>
                        <Link
                          onClick={() => {
                            this.setState({
                              showHistory: { [section.title]: true },
                            });
                          }}
                          title={strings.ViewHistoryTitle}
                        >
                          <Text
                            block
                            variant="smallPlus"
                            className={styles.ViewHistorylineTitle}
                          >
                            {strings.ViewHistoryTitle}
                          </Text>
                        </Link>
                      </Stack.Item>
                    )}

                    <WorkflowHistory
                      requests={section.requests}
                      historyData={section.historyData}
                      isOpen={this.state.showHistory[section.title]}
                      key={`wfh-${section.requests?.length || 0}`}
                      containerClassName={styles.DialogContainer}
                      title={`${section.title}`}
                      onDismiss={() => {
                        this.setState({
                          showHistory: {
                            [strings.ProofReadingTitle]: false,
                            [strings.AvvaTitle]: false,
                          },
                        });
                      }}
                      cultureName={cultureName}
                      timeLineTasks={section.timeLineTasks}
                      requestService={requestService}
                    />
                  </React.Fragment>
                );
              })}
            </Stack>
            {requests && requests.tasks && requests.tasks.length > 0 && (
              <Stack>
                <div className={styles.cancelWorkflow}>
                  {
                    this.state.isValidUser && <PrimaryButton
                    onClick={() => {
                      this.setState({
                        showConfirmCancelation: true,
                        cancelationReason: undefined,
                      });
                    }}
                    text={"Cancel Workflow"}
                    disabled={this.state.disableCancelButton && this.state.isValidUser && this.state.isCheckDocument}
                  />
                  }
                 
                </div>
              </Stack>
            )}
            {this.props.isthreeSixtyDegree && (
              <LogHistory 
              context={this.props.context}
              dmsClient={this.props.dmsClient}
  httpTriggerEndPoint={this.props.httpTriggerEndPoint}
              />
            )}
          </Panel>
        )}
      </div>
    );
  }

  private _cancelRequest = async () => {
    const { requestService } = this.props;
    const { cancelationReason, wfCancelDetails } = this.state;
    //let taskID: number[] = [];
    if (!wfCancelDetails) {
      return;
    }

    try{
      this.setState({ isCanceling: true });
      await requestService
      .triggerFlow(
        wfCancelDetails,
        cancelationReason || "",
        this.props.workflowUrl
      );
      this.setState({ isCancelingSuccess: true, isCanceling: false });
    } catch(err){
        Log.error(LOG_SOURCE, err);
        console.error(err);
        this.setState({ isCancelingError: true, isCanceling: false });
    }
    

      /*.then(() => {
        this.setState({ isCancelingSuccess: true, isCanceling: false });
      })
      .catch((err) => {
        Log.error(LOG_SOURCE, err);
        console.error(err);
        this.setState({ isCancelingError: true, isCanceling: false });
      });*/
    /* await this.state.timeLineSections.map(async timeSec => {
       const iprog = timeSec.steps.filter(s => s.status.toLowerCase() === 'in progress');
       const d = iprog.map(a => a.taskId);
       d.map(i=>{
         taskID.push(i);
       })
     });
     //requestService.triggerFlow();
     await requestService.cancelWorkflow(taskID, cancelationReason);*/
    // this.setState({ isCancelingSuccess: true, isCanceling: false });
  };
}
