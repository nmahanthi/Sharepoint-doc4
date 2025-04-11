import { Log } from "@microsoft/sp-core-library";
import { Icon } from '@fluentui/react/lib/Icon'
import * as React from "react";
import styles from './WorkflowTimeline.module.scss';
import { Stack } from "@fluentui/react/lib/Stack";
import { Callout } from "@fluentui/react/lib/Callout";
import { IconButton } from "@fluentui/react/lib/Button";
import { Text } from "@fluentui/react/lib/Text";
import * as strings from "DmsContexualMenuCommandSetCommandSetStrings";
import { IShimmerElement, Shimmer, ShimmerElementType } from "@fluentui/react/lib/Shimmer";
import CONSTANTS from "../../_Constants";

const LOG_SOURCE: string = 'WorkflowTimeline';

export interface IWorkflowTimelineStep {
  taskId?: number;
  name: string;
  person?: string;
  date?: string;
  delegator?:string;
  status?: 'completed' | 'in progress' | 'rejected' | 'error' | 'delegated' | 'cancelled' ;
  //status?: 'approved' | 'rejected' | 'error' | 'delegated';
  comments?: string;
  reminderUrl?: string;
}
export interface IWorkflowTimelineProps {
  graphHeight: number;
  steps: IWorkflowTimelineStep[];
  isLoading?: boolean;
  isError?: boolean;
  requestOutcome?: string;
  selectedId?: number;
  timeLineType?: string;
}
export interface IWorkflowTimelineState {
  lineWidth: number;
  stepCalloutVisible: { [stepName: string]: boolean };
}
export default class WorkflowTimeline extends React.Component<IWorkflowTimelineProps, IWorkflowTimelineState> {
  private _containerRef: React.RefObject<HTMLDivElement>;
  /**
   *
   */
 // private _uniqueId: Guid;

  constructor(props: IWorkflowTimelineProps) {
    super(props);
    this.state = { lineWidth: props.graphHeight / 2, stepCalloutVisible: {} };
    this._containerRef = React.createRef();
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: WorkflowTimeline mounted');
   // this._uniqueId = Guid.newGuid();
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: WorkflowTimeline unmounted');
  }

  public render(): React.ReactElement<{}> {
    const { steps, graphHeight, isLoading, isError, requestOutcome, timeLineType } = this.props;
    const { stepCalloutVisible } = this.state;
    const lineWidth: number = this._getLineWidth(steps?.length || 0, graphHeight);
    let inProgressIndex: number = -1;
    const svgString: string = steps?.reduce((prev: { result: string; isPrevDisabled: boolean }, current: IWorkflowTimelineStep, index: number) => {
      let circleXoffset = lineWidth / 2;
      if (index > 0) {
        const isLineDisabled = prev.isPrevDisabled || (steps[index - 1].status === 'error' || steps[index - 1].status === 'rejected')
        const lineColor = isLineDisabled ? 'rgb(217, 217, 217)' : 'rgb(22, 33, 52)';
        const lineX1 = (index * graphHeight) + (lineWidth * (index - 1)) + circleXoffset;
        const lineX2 = (index * graphHeight) + (lineWidth * index) + circleXoffset;
        prev.result += `<line x1='${lineX1}' y1='${graphHeight / 2}' x2='${lineX2}' y2='${graphHeight / 2}' stroke='${lineColor}' stroke-width='5' />`;
        circleXoffset = lineX2
      }

      if (this.props.timeLineType === strings.AvvaTitle) {
        if (current.status === 'in progress' && inProgressIndex=== -1)
          inProgressIndex = index;
      } else {
        if (current.status === 'in progress')
          inProgressIndex = index;
      }
      // const isCurrentStepActive = (
      //   index === 0 && !current.status ||
      //   (index > 0 && !!steps[index - 1].status && !current.status)
      // );
     // const isCurrentStepDisabled = !current.status && requestOutcome === CONSTANTS.AbortedStatus ||
     const isCurrentStepDisabled = !current.status && requestOutcome === CONSTANTS.AbortedStatus || steps[index].status==='cancelled' ||
        (index > 0 && (steps[index - 1].status === 'error' || steps[index - 1].status === 'rejected' || prev.isPrevDisabled));
      const circleX = (graphHeight / 2) + circleXoffset;
      if (current.taskId && current.taskId === this.props.selectedId) {
        prev.result += `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${(graphHeight / 2) - 1.5}' stroke='rgb(214, 47, 32)' stroke-width='3' stroke-dasharray="0 6" stroke-linecap="round" fill="transparent" />`;
      } else {
        if (isCurrentStepDisabled) {
          prev.result += `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${graphHeight / 2}' stroke='rgb(217, 217, 217)' stroke-width='0' fill='rgb(217, 217, 217)' />`;
        } else {
          if (inProgressIndex === index) {
            //   if (isCurrentStepActive) {
            prev.result += `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${(graphHeight / 2) - 2.5}' stroke='rgb(22, 33, 52)' stroke-width='5' fill='rgb(214, 47, 32)' />`;
          } else {
            prev.result += `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${graphHeight / 2}' stroke='rgb(22, 33, 52)' stroke-width='0' fill='rgb(22, 33, 52)' />`;
          }
        }
      }
      // prev.result += isCurrentStepDisabled ?
      //   `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${graphHeight / 2}' stroke='rgb(217, 217, 217)' stroke-width='0' fill='rgb(217, 217, 217)' />` : isCurrentStepActive ?
      //     `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${(graphHeight / 2) - 2.5}' stroke='rgb(22, 33, 52)' stroke-width='5' fill='rgb(214, 47, 32)' />` :
      //     `<circle cx='${circleX}' cy='${graphHeight / 2}' r='${graphHeight / 2}' stroke='rgb(22, 33, 52)' stroke-width='0' fill='rgb(22, 33, 52)' />`
      prev.isPrevDisabled = isCurrentStepDisabled;

      return prev;
    }, { result: '', isPrevDisabled: false }).result;

    const shimmerVerticalElement: IShimmerElement[] = [
      { type: ShimmerElementType.circle, height: graphHeight },
      { type: ShimmerElementType.line, verticalAlign: 'center' },
      { type: ShimmerElementType.circle, height: graphHeight },
      { type: ShimmerElementType.line, verticalAlign: 'center' },
      { type: ShimmerElementType.circle, height: graphHeight }
    ];
    return <div className={styles.WorkflowTimeline} ref={this._containerRef}>
      {steps && steps.length > 0 && !isLoading && !isError ? <React.Fragment>
        <div className={styles.StatusIconContainer} >
          {
            steps.map(step => {
              let iconName = '';
              let iconColor = '';
              switch (step.status) {
                case 'completed':
                  iconName = 'CheckMark';
                  iconColor = 'lightseagreen';
                  break;
                case 'rejected':
                  iconName = 'Cancel';
                  iconColor = 'red';
                  break;
                case 'delegated':
                  iconName = 'Switch';
                  iconColor = 'orange';
                  break;
                  // case 'cancelled':
                  //   iconName = 'Cancel';
                  //   iconColor = 'red';
                  //   break;
                case 'error':
                  iconName = 'Error';
                  iconColor = 'red';
                  break;
              }
              return <div style={{ width: `${lineWidth + graphHeight}px` }} >
                {iconName && <Icon title={`${step.status?.substring(0, 1).toUpperCase()}${step.status?.substring(1).toLocaleLowerCase()}`}
                  iconName={iconName} styles={{ root: { color: iconColor, marginLeft: `${graphHeight / 2}px`, fontWeight: 'bolder' } }} />}
              </div>;
            })
          }
        </div>
        <svg width='100%' height={`${graphHeight}px`} dangerouslySetInnerHTML={{ __html: svgString }}>
        </svg>
        <div className={styles.ContentContainer} >
          {
            steps.map((step, stepIndex) => {
              return <div style={{ width: `${lineWidth + graphHeight}px` }} >
                <Stack>
                  <Stack.Item className={styles.Title}>{step.name.replace(CONSTANTS.ProofReadingStatusName,"")}</Stack.Item>
                  <Stack.Item className={styles.Subtitle} title={step.person}>{step.person}</Stack.Item>
                  {step.delegator && <Stack.Item className={styles.Subtitle} >Delegated by : {step.delegator}</Stack.Item>}
                  <Stack.Item className={styles.Subtitle}>{step.date}</Stack.Item>
                  <Stack.Item>
                    {step.comments &&
                      <React.Fragment>
                        <IconButton
                          id={`${timeLineType}${stepIndex}-comments-${timeLineType}`}
                          onClick={() => {
                            const newCalloutState = { ...stepCalloutVisible };
                            newCalloutState[stepIndex] = !newCalloutState[stepIndex];
                            this.setState({ stepCalloutVisible: newCalloutState })
                          }}
                          iconProps={{ iconName: 'Message' }}
                        />
                        {stepCalloutVisible[stepIndex] && (
                          <Callout
                            role="dialog"
                            gapSpace={0}
                            target={`#${timeLineType}${stepIndex}-comments-${timeLineType}`}
                            onDismiss={() => {
                              const newCalloutState = { ...stepCalloutVisible };
                              newCalloutState[stepIndex] = !newCalloutState[stepIndex];
                              this.setState({ stepCalloutVisible: newCalloutState })
                            }}
                            setInitialFocus
                            className={styles.CommentsCallout}
                          >
                            <Text block variant="small">
                              {step.comments}
                            </Text>
                          </Callout>)
                        }
                      </React.Fragment>
                    }
                    {!step.status && step.reminderUrl && (stepIndex === 0 || steps[stepIndex - 1].status === 'completed') &&
                      <IconButton onClick={() => {
                        if(step.reminderUrl){
                          window.location.href = step.reminderUrl 
                        }
                      }} 
                      iconProps={{ iconName: 'Mail' }} />
                    }
                  </Stack.Item>
                </Stack>
              </div>;
            })
          }
        </div></React.Fragment>
        :
        (!steps || steps.length === 0) && !isLoading && !isError ?
          <Stack styles={{ root: { height: `${graphHeight}px` } }} horizontalAlign="center">
            <Text variant="small" className={styles.NoWfMessage}>{strings.NoWfMessage}</Text>
          </Stack> :
          !isError ?
            <Stack styles={{ root: { height: `${graphHeight}px`, paddingLeft: 30, paddingRight: 30 } }} horizontalAlign="space-evenly" grow>
              <Shimmer shimmerElements={shimmerVerticalElement} />
            </Stack> :
            <Stack styles={{ root: { height: `${graphHeight}px` } }} horizontalAlign="center" grow>
              <Text variant="small" className={styles.ErrorMessage}>{strings.WfTimelineErrorMessage}</Text>
            </Stack>
      }
    </div>;
  }
  private _getLineWidth(stepsLength: number, graphHeight: number): number {
       const panelWidth= 1000;//Adjusted based on panel width
     return ((panelWidth) - (graphHeight * stepsLength)) / Math.max(stepsLength, 1);
   /* return this._containerRef?.current ? Math.max(
      ((this._containerRef?.current?.clientWidth || 0) - (graphHeight * stepsLength)) / Math.max(stepsLength, 1),
      graphHeight
    ) : graphHeight;*/
  }
}