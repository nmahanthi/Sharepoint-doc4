/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  PrimaryButton,
} from "@fluentui/react";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { TextField } from '@fluentui/react/lib/TextField';
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FieldCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp/fi";
import { RenderListDataOptions } from "@pnp/sp/lists";
import { getSPCacheNever } from "../../pnpjs-config";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { INotificationService } from "../../service/NotificationService";
import { Guid } from "@microsoft/sp-core-library";
import { FieldNames, ListUrls } from "../../constants";
import styles from "./ProofreadingDialog.module.scss";
export interface IWorkflowDialogContentProps {
  onClose: () => void;
  context: ListViewCommandSetContext | FieldCustomizerContext;
  itemID: number;
  libraryID: string;
  siteURL: string;
  currentUser: string;
  currentUserEmail: string;
  docGUId: string;
  createdBy: { EMail: string; Title: string; UserName: string; ID: number };
  projReference: string;
  projRevision:string;
  docTitle: string;
  itemLink: string;
  workflowUrl: string;
  notificationService: INotificationService;
}
export interface IProofreadingDialogContentState {
  reviewers: any[];
  isDisabled: boolean;
  taskStatus: string;
  workflowStatus: string;
  closureDate:string;
  groupIds: number[];
  documentStatus?: string;
  refreshInterval?: NodeJS.Timeout;
  isProcessing?: boolean;
  isRevisionValue:boolean;
  errorMessage:string
}
class ProofreadingDialogContent extends React.Component<
  IWorkflowDialogContentProps,
  IProofreadingDialogContentState
> {
  private async _pollStatus(OperationID: Guid, attempt: number) {
    const { notificationService, itemID, libraryID } = this.props;
    const item = await this._sp.web.lists.getById(libraryID).items.getById(itemID)
      .select(FieldNames.DocumentStatus)();
    if (item[FieldNames.DocumentStatus] === "Draft" && attempt < 20) {
      setTimeout(async () => {
        await this._pollStatus(OperationID, attempt + 1);
      }, 3000);
    } else {      
      this.props.onClose();
      notificationService.setOperationCompleted(OperationID, true);
      window.location.reload();
    }
  }
  private _sp: SPFI;
  constructor(props: IWorkflowDialogContentProps) {
    super(props);
    this.state = {
      reviewers: [],
      isDisabled: true,
      taskStatus: "TaskStatus",
      workflowStatus: "WorkflowStatus",
      closureDate:'',
      groupIds: [],
      isProcessing: false,
      isRevisionValue:false,
      errorMessage:''
    };
    // this._sp = getSP(props.context);
    this._sp = getSPCacheNever(props.context);
  }
 /* private _getPeoplePickerItems(items: any[]) {
    if (items.length > 0) {
      this.setState({ reviewers: items, isDisabled: false });
    } else {
      this.setState({ reviewers: [], isDisabled: true });
    }
    console.log("reviewers", this.state.reviewers);
  }*/
 /* _onClosureDateChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ closureDate: newValue || '' });
  };*/


  private _getPeoplePickerItems(items: any[]) {
    if (items.length > 0) {
      this.setState({ reviewers: items }, () => {
        this._updateStartButtonState();
      });
    } else {
      this.setState({ reviewers: [] }, () => {
        this._updateStartButtonState();
      });
    }
  }
  // _onClosureDateChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
  //   this.setState({ closureDate: newValue || '' }, () => {
  //     this._updateStartButtonState();
  //   });
  // };
 
  // private _updateStartButtonState() {
  //   const isReviewersSelected = this.state.reviewers.length > 0;
  //   const isClosureDateSelected = this.state.closureDate.trim() !== '';
  //   this.setState({ isDisabled: !(isReviewersSelected && isClosureDateSelected) });
  // }

  private _onClosureDateChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    const newClosureDate = newValue || '';
    const today = new Date();
  
    // Set the time of 'today' to the start of the day to compare only dates
    today.setHours(0, 0, 0, 0);

    const selectedDate = new Date(newClosureDate);
    let errorMessage = '';
  
    // Check if the selected date is before today
    if (selectedDate < today) {
        errorMessage = "Past Date should not be set as closure Date.";
    }
  
    this.setState({
        closureDate: newClosureDate,
        errorMessage: errorMessage
    }, () => {
        this._updateStartButtonState();
    });
};

private _updateStartButtonState() {
  const isReviewersSelected = this.state.reviewers.length > 0;
  const isClosureDateSelected = this.state.closureDate.trim() !== '';
  const isValidDate = !this.state.errorMessage; // Ensure there are no error messages

  this.setState({ isDisabled: !(isReviewersSelected && isClosureDateSelected && isValidDate) });
}
 

  public async componentDidMount(): Promise<void> {
    const [batch, execute] = this._sp.batched();
    Promise.all([
      batch.web.lists.select("Id").filter(`RootFolder/ServerRelativeUrl eq '${this.props.context.pageContext.web.serverRelativeUrl}/${ListUrls.Draft}'`)(),
      batch.web.lists.select("Id").filter(`RootFolder/ServerRelativeUrl eq '${this.props.context.pageContext.web.serverRelativeUrl}/${ListUrls.ApplicableDocuments}'`)(),
      batch.web.lists.select("Id").filter(`RootFolder/ServerRelativeUrl eq '${this.props.context.pageContext.web.serverRelativeUrl}/${ListUrls.PreviousVersions}'`)(),
    ])
      .then(async ([draftLib, applicableLib, prevLib]) => {
        if (
          !draftLib ||
          draftLib.length === 0 ||
          !applicableLib ||
          applicableLib.length === 0 ||
          !prevLib ||
          prevLib.length === 0
        ) {
          console.error(new Error("Document libraries not found"));
        } else {
          const draftLibraryId = draftLib[0].Id;
          const applicableLibraryId = applicableLib[0].Id;
          const prevLibraryId = prevLib[0].Id;
          const libraries = [draftLibraryId,applicableLibraryId,prevLibraryId,];
          let totalCount = 0; 
          try {
            const camlQuery = `<View Scope="RecursiveAll"><Query><Where><And><Eq><FieldRef Name="ProjectReference"/><Value Type="Text">${this.props.projReference}</Value></Eq><Eq><FieldRef Name="ProjectRevision"/><Value Type="Text">${this.props.projRevision}</Value></Eq></And></Where></Query></View>`;
            for (const libraryId of libraries) {
              const listItems = await this._sp.web.lists
                .getById(libraryId)
                .renderListDataAsStream({
                  ViewXml: camlQuery,
                  RenderOptions: RenderListDataOptions.ListData,
                });
              const itemCount = listItems?.Row?.length || 0;
              totalCount += itemCount; 
            }
            if (totalCount > 1) {
              this.setState({isRevisionValue: true});
            }
          } catch (error) {
            console.error("Error fetching data from libraries:", error);
          }
        }
      })
      .catch((error) => {
        console.error(error);
      });

    execute().catch((error) => {
      console.error(error);
    });
  }

  private _startWorkflow = async () => {

   
    if(!this.state.isRevisionValue){
       const requestUrl = this.props.workflowUrl;
        this.setState({ isProcessing: true });
         const OperationID =
        this.props.notificationService.addOngoingOperation("StartingWorkflow");
       try {
        const response: HttpClientResponse =
        await this.props.context.httpClient.post(
          requestUrl,
          HttpClient.configurations.v1,
          {
            headers: {
              "Content-type": "application/json",
              Accept: "application/json",
            },
            body: JSON.stringify({
              itemID: this.props.itemID,
              libraryID: this.props.libraryID,
              siteURL: this.props.siteURL,
              reviewers: this.state.reviewers,
              closureDate:this.state.closureDate,
              currentUser: this.props.currentUser,
              docGUId: this.props.docGUId,
              createdBy: this.props.createdBy,
              currentUserEmail: this.props.currentUserEmail,
              projReference: this.props.projReference,
              docTitle: this.props.docTitle,
              itemLink: this.props.itemLink,
              taskStatusList: this.state.taskStatus,
              workflowStatusList: this.state.workflowStatus,
            }),
          }
        );
        if (response.ok) {
        await this._pollStatus(OperationID, 0);
        console.log("Flow triggered successfully");
         } else {
        this.props.notificationService.setOperationCompleted(
          OperationID,
          false,
          response.statusText
        );
        console.error("Flow triggered faild", response.statusText);
         }
        } catch (error) {
         this.props.notificationService.setOperationCompleted(OperationID,false,error );
  
       }
    }else{
       this.setState({
        errorMessage:"Revision value not unique"
       })
      }
  };
  public render(): React.ReactElement<{}> {
    
    return (
      <>
        <Dialog
          hidden={false}
          modalProps={{ isBlocking: true }}
          onDismiss={this.props.onClose}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Start Proofreading Workflow",
          }}
        >
          {this.state.errorMessage && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.errorMessage}
              </MessageBar>
            )}
          
          {" "}
          {
            this.state.isProcessing ?
              <div>
                <h4>
                  Processing, Please wait...
                </h4>
              </div>
              :
              <div>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Select Reviewers"
                  personSelectionLimit={10}
                  required={true}
                  placeholder="Please add Reviewers"
                  // groupName={"DMS-DEV1"}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={100}
                  onChange={this._getPeoplePickerItems.bind(this)}
                  // onChange={_onPeoplePickerChange(setReviewers)}
                  webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
                  suggestionsLimit={10}
                />
            <TextField
              label="Closure Date"
              type='date'
              value={this.state.closureDate}
              onChange={this._onClosureDateChange}
              styles={{ root: { width: '48%' } }}
              required
              min={new Date().toISOString().split("T")[0]} 
            />
            </div>
             

          }
          <DialogFooter>
            {this.state.isProcessing ?
              <div>
              </div>
              :
              <div>
                <PrimaryButton
                  onClick={this._startWorkflow}
                  text="Start"
                  disabled={this.state.isDisabled}
                  className={styles.saveButton}
                />
                <DefaultButton onClick={this.props.onClose} text="Cancel" />
              </div>
            }
          </DialogFooter>
        </Dialog >
      </>
    );
  }
}
export default ProofreadingDialogContent;
