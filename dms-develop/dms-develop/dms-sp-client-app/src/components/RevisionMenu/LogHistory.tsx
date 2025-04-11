import * as React from "react";
import { Icon, Stack, Text, Link, mergeStyleSets, Spinner, MessageBar, MessageBarType } from "@fluentui/react";
import { ActivityItem } from "@fluentui/react";
import { FieldCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { AadHttpClient } from "@microsoft/sp-http";
import { LogHistoryService } from "../../service/LogHistoryService";
import { ILogHistoryItem } from "../../interfaces/ILogHistoryItem";

interface LogHistoryState {
  logHistoryItems: ILogHistoryItem[];
  isLoading: boolean;
  error: string | null;
}

interface ILogHistoryProps {
  context: ListViewCommandSetContext |FieldCustomizerContext;
  dmsClient: AadHttpClient;
  httpTriggerEndPoint: string;
}

const classNames = mergeStyleSets({
  exampleRoot: {
    marginTop: "20px",
  },
  nameText: {
    fontWeight: "bold",
  },
  scrollableStack: {
    maxHeight: "700px",
    overflowY: "auto",
  },
});

class LogHistory extends React.Component<ILogHistoryProps, LogHistoryState> {
  private logHistoryService: LogHistoryService;

  constructor(props: ILogHistoryProps) {
    super(props);
    this.state = {
      logHistoryItems: [],
      isLoading: true,
      error: null,
    };
    this.logHistoryService = new LogHistoryService(this.props.dmsClient);
  }

  async componentDidMount(): Promise<void> {
    try {
      const logHistoryItems = await this.logHistoryService.getLogHistoryItems(this.props.httpTriggerEndPoint);
      this.setState({ logHistoryItems: logHistoryItems.reverse(), isLoading: false });
    } catch (error) {
      console.error("Error fetching log history:", error);
      this.setState({ error: "Failed to load log history data.", isLoading: false });
    }
  }

  render(): JSX.Element {
    const { logHistoryItems, isLoading, error } = this.state;

    return (
      <>
        <Stack
          horizontalAlign="center"
          verticalAlign="center"
          styles={{
            root: {
              backgroundColor: "#163350d6",
              color: "white",
              marginTop: 14,
              height: 50,
            },
          }}
        >
          <Text
            variant="large"
            as="h3"
            styles={{ root: { color: "white", fontSize: 22 } }}
          >
            Log History
          </Text>
        </Stack>

        {isLoading ? (
          <Spinner label="Loading log history..." />
        ) : error ? (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        ) : logHistoryItems.length === 0 ? (
          <Text variant="large" styles={{ root: { marginTop: 20 } }}>
            No update history data
          </Text>
        ) : (
          <Stack
            style={{ height: "200px", overflow: "auto", marginBottom: "20px" }}
            className={classNames.scrollableStack}
          >
            {logHistoryItems.map((item) => this.renderActivityItem(item))}
          </Stack>
        )}
      </>
    );
  }

  private renderActivityItem(item: ILogHistoryItem): JSX.Element {
    return (
      <ActivityItem
        key={item.documentId}
        activityDescription={this.getActivityDescription(item)}
        activityIcon={<Icon iconName={this.getIconName(item.actionType)} />}
        comments={this.getComments(item)}
        timeStamp={item.date}
        className={classNames.exampleRoot}
      />
    );
  }

  private getActivityDescription(item: ILogHistoryItem): JSX.Element[] {
    return [
      <Link
        key={1}
        className={classNames.nameText}
      >
        {item.userDisplayName}
      </Link>,
      <span key={2}> {item.actionType.toLowerCase()}d</span>,
      <>
        <div key={3}><span>Project Reference: </span>{item.projectReference}</div>
        <div key={4}><span>Project Revision: </span>{item.projectRevision}</div>
      </>
    ];
  }

  private getComments(item: ILogHistoryItem): JSX.Element {
    return (
      <ul style={{ marginTop: '5px', paddingLeft: '20px' }}>
     { item.changedFields?.filter(flds=>!flds.isHidden && !flds.isReadOnly).map((field, index) => (
        // {item.changedFields?.map((field, index) => (
          <li key={index}>
            <strong>{field.fieldTitle}:</strong>
            <span style={{paddingLeft:'3px',color:'#000'}} > {field.newValue} </span>
             {field.oldValue &&<span style={{padding:'5px'}}>-</span> }
             <span style={{ color: '#bec0c6' }}> {field.oldValue}</span>
          </li>
        ))}
      </ul>
    );
  }

  private getIconName(actionType: string): string {
    switch (actionType) {
      case "create":
        return "Add";
      case "update":
        return "Edit";
      case "delete":
        return "Trash";
      default:
        return "Info";
    }
  }
}

export default LogHistory;