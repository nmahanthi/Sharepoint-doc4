import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, PrimaryButton } from '@fluentui/react';
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { TextField } from '@fluentui/react/lib/TextField';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPFI } from "@pnp/sp/fi";
import { RenderListDataOptions } from "@pnp/sp/lists";
import { FieldCustomizerContext, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import * as moment from 'moment';
import styles from './AVVADialog.module.scss';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { Guid } from '@microsoft/sp-core-library';
import { ListUrls } from "../../constants";
import { INotificationService } from '../../service/NotificationService';
import { FieldNames } from '../../constants';
import { getSPCacheNever } from "../../pnpjs-config";

export interface IWorkflowDialogContentProps {
  onClose: () => void;
  context: ListViewCommandSetContext | FieldCustomizerContext;
  itemID: number;
  libraryID: string;
  siteURL: string;
  currentUser: string;
  currentUserEmail: string;
  projReference: string;
  projRevision: string;
  docTitle: string;
  itemLink: string;
  docGUId: string;
  workflowUrl: string;
  createdBy: { EMail: string; Title: string; UserName: string; ID: number; };
  notificationService: INotificationService;
}
export interface IWorkflowDialogContentState {
  author: any[];
  verifier: any[];
  validator: any[];
  approver: any[];
  textFieldValue: string;
  dueDate1: string;
  dueDate2: string;
  dueDate3: string;
  dueDate4: string;
  eSign: boolean;
  validationError: string | null;
  groupIds: number[];
  isEsignEnabled: boolean;
  isProcessing?: boolean;
  isRevisionValue: boolean;
  errorMessage: string;
}

class WorkflowDialogContent extends React.Component<IWorkflowDialogContentProps, IWorkflowDialogContentState> {
  private _sp: SPFI;
  constructor(props: IWorkflowDialogContentProps) {
    super(props);
    this.state = {
      author: [],
      verifier: [],
      validator: [],
      approver: [],
      textFieldValue: '',
      dueDate1: '',
      dueDate2: '',
      dueDate3: '',
      dueDate4: '',
      eSign: false,
      validationError: null,
      groupIds: [],
      isEsignEnabled: false,
      isRevisionValue: false,
      errorMessage: ''
    };
    this._sp = getSPCacheNever(props.context);
  }

  componentDidMount() {
    this.fetchGroupIds()
      .catch(error => {
        console.error('Error fetching group IDs:', error);
      });
    const initialDate = moment();
    const authorDueDate = this.addBusinessDays(initialDate, 2);
    const verifierDueDate = this.addBusinessDays(authorDueDate, 2);
    const validatorDueDate = this.addBusinessDays(verifierDueDate, 2);
    const approverDueDate = this.addBusinessDays(validatorDueDate, 2);

    this.setState({
      dueDate1: authorDueDate.format('YYYY-MM-DD'),
      dueDate2: verifierDueDate.format('YYYY-MM-DD'),
      dueDate3: validatorDueDate.format('YYYY-MM-DD'),
      dueDate4: approverDueDate.format('YYYY-MM-DD'),
    });

    //check revision and reference
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
          const libraries = [draftLibraryId, applicableLibraryId, prevLibraryId,];
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
              this.setState({ isRevisionValue: true });
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

  private async _pollStatus(OperationID: Guid, attempt: number) {

    const { notificationService, itemID, libraryID } = this.props;
    console.log("item:" + await this._sp.web.lists.getById(libraryID).items.getById(itemID));
    const item = await this._sp.web.lists.getById(libraryID).items.getById(itemID).select(FieldNames.DocumentStatus)();
    console.log(item);
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

  addBusinessDays = (startDate: moment.Moment, days: number) => {
    let currentDate = moment(startDate);
    let addedDays = 0;

    while (addedDays < days) {
      currentDate = currentDate.add(1, 'days');

      if (currentDate.isoWeekday() !== 6 && currentDate.isoWeekday() !== 7) {
        addedDays += 1;
      }
    }

    return currentDate;
  };

  fetchGroupIds = async () => {
    try {
      const groupRequestHeaders: HeadersInit = new Headers();
      groupRequestHeaders.append("Accept", "application/json;odata=nometadata");

      const groupApiUrl = `${this.props.siteURL}/_api/web/sitegroups`;
      const response: HttpClientResponse = await this.props.context.httpClient.get(
        groupApiUrl,
        HttpClient.configurations.v1,
        {
          headers: groupRequestHeaders,
        }
      );

      if (!response.ok) {
        throw new Error('Error fetching SharePoint groups');
      }
      const groupData = await response.json();
      const groups = groupData.value;

      const ids = groups.map((group: any) => group.Id);
      const groupIdsArray: number[] = ids.map((id: any) => parseInt(id));

      this.setState({ groupIds: groupIdsArray });
      console.log('Group IDs:', groupIdsArray);
      console.log('Type of first groupId:', typeof groupIdsArray[0]);
    } catch (error) {
      console.error('Error fetching group IDs:', error);
    }
  };

  _onAuthorPeoplePickerChange = (items: any[]) => {
    console.log('Author People Picker Items:', items);
    this.setState({ author: items });
  };

  _onVerifierPeoplePickerChange = (items: any[]) => {
    console.log('Verifier People Picker Items:', items);
    this.setState({ verifier: items });
  };

  _onValidatorPeoplePickerChange = (items: any[]) => {
    console.log('Validator People Picker Items:', items);
    this.setState({ validator: items });
  };

  _onApproverPeoplePickerChange = (items: any[]) => {
    console.log('Approver People Picker Items:', items);
    this.setState({ approver: items });
  };

  _onDueDate1Change = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ dueDate1: newValue || '' });
  };

  _onDueDate2Change = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ dueDate2: newValue || '' });
  };

  _onDueDate3Change = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ dueDate3: newValue || '' });
  };

  _onDueDate4Change = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ dueDate4: newValue || '' });
  };

  _onESignChange = (event: React.FormEvent<HTMLElement>, checked?: boolean) => {
    this.setState({ eSign: !!checked });
  };

  _startWorkflow = async () => {
    if (!this.state.isRevisionValue) {

      const currentDate = moment().format('YYYY-MM-DD');

      const { author, verifier, validator, approver, dueDate1, dueDate2, dueDate3, dueDate4, eSign } = this.state;


      if (
        !author.length || !verifier.length || !validator.length || !approver.length ||
        !dueDate1 || !dueDate2 || !dueDate3 || !dueDate4
      ) {
        this.setState({ validationError: "Please ensure all fields are filled." });
        return;
      }

      const peopleIds = [author[0]?.id, verifier[0]?.id, validator[0]?.id, approver[0]?.id];
      const uniquePeopleIds = new Set(peopleIds);

      if (uniquePeopleIds.size < 4) {
        this.setState({ validationError: "Author, Verifier, Validator, and Approver cannot be the same person." });
        return;
      }

      if (
        !moment(dueDate1).isAfter(currentDate) ||
        !moment(dueDate2).isAfter(currentDate) ||
        !moment(dueDate3).isAfter(currentDate) ||
        !moment(dueDate4).isAfter(currentDate)
      ) {
        this.setState({ validationError: "All due dates must be greater than the current date." });
        return;
      }

      if (
        !moment(dueDate1).isBefore(dueDate2) ||
        !moment(dueDate1).isBefore(dueDate3) ||
        !moment(dueDate1).isBefore(dueDate4)
      ) {
        this.setState({
          validationError:
            "Author's due date should be less than Verifier's, Validator's, and Approver's due dates. Starting from Author to Approver, the due date should be in ascending order."
        });
        return;
      }

      if (
        !moment(dueDate2).isBefore(dueDate3) ||
        !moment(dueDate2).isBefore(dueDate4) ||
        !moment(dueDate2).isAfter(dueDate1)
      ) {
        this.setState({
          validationError:
            "Verifier's due date should be less than Validator's and Approver's due dates, and greater than Author's due date. Starting from Author to Approver, the due date should be in ascending order."
        });
        return;
      }

      if (
        !moment(dueDate3).isBefore(dueDate4) ||
        !moment(dueDate3).isAfter(dueDate1) ||
        !moment(dueDate3).isAfter(dueDate2)
      ) {
        this.setState({
          validationError:
            "Validator's due date should be less than Approver's due date, and greater than Author's and Verifier's due dates. Starting from Author to Approver, the due date should be in ascending order."
        });
        return;
      }

      if (
        !moment(dueDate4).isAfter(dueDate1) ||
        !moment(dueDate4).isAfter(dueDate2) ||
        !moment(dueDate4).isAfter(dueDate3)
      ) {
        this.setState({
          validationError:
            "Approver's due date should be greater than Author's, Verifier's, and Validator's due dates. Starting from Author to Approver, the due date should be in ascending order."
        });
        return;
      }

      console.log('Starting workflow with:', {
        author,
        verifier,
        validator,
        approver,
        dueDate1,
        dueDate2,
        dueDate3,
        dueDate4,
        eSign
      });

      this.setState({ isProcessing: true });
      const OperationID = this.props.notificationService.addOngoingOperation("StartingWorkflow");
      const requestBody = {
        author: author[0],
        verifier: verifier[0],
        validator: validator[0],
        approver: approver[0],
        dueDate1,
        dueDate2,
        dueDate3,
        dueDate4,
        eSign,
        itemID: this.props.itemID,
        libraryID: this.props.libraryID,
        siteURL: this.props.siteURL,
        currentUser: this.props.currentUser,
        currentUserEmail: this.props.currentUserEmail,
        docGUId: this.props.docGUId,
        requestor: this.props.createdBy,
        projReference: this.props.projReference,
        projRevision: this.props.projRevision,
        docTitle: this.props.docTitle,
        itemLink: this.props.itemLink
      };
      console.log(requestBody);
      const requestOptions: IHttpClientOptions = {
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
      };

      try {
        const response: HttpClientResponse = await this.props.context.httpClient.post(
          this.props.workflowUrl,
          HttpClient.configurations.v1,
          requestOptions
        );

        if (!response.ok) {
          this.props.notificationService.setOperationCompleted(
            OperationID,
            false,
            response.statusText
          );
          throw new Error('Network response was not ok');
        }
        // const data = await response.json();
        await this._pollStatus(OperationID, 0);
        this.props.onClose();
        // console.log('Response from HTTP request:', data);
      } catch (error) {
        this.props.notificationService.setOperationCompleted(
          OperationID,
          false,
          error
        );
        console.error('Error sending HTTP request:', error);
      }
    }
    else {
      this.setState({
        errorMessage: "Revision value not unique"
      })
    }
   // this.props.onClose();
  };

  render() {
    const {

      dueDate1,
      dueDate2,
      dueDate3,
      dueDate4,
      eSign,
      validationError,
      groupIds,
      isEsignEnabled
    } = this.state;
    const props = this.props;
    if (groupIds.length === 0) {
      return <div>Loading...</div>;
    }
    return (
      <Dialog
        hidden={false}
        onDismiss={props.onClose}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Start AVVA Workflow',
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { minWidth: 800, maxWidth: 800 } },
        }}
        containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
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
        {
          !this.state.isProcessing &&
          <div>
            {validationError && <div style={{ color: 'red', marginBottom: '10px' }}>{validationError}</div>}
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              {groupIds.length > 0 && (
                <PeoplePicker
                  context={props.context as any}
                  titleText="Author"
                  personSelectionLimit={1}
                  groupId={groupIds}
                  groupName={""}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  onChange={this._onAuthorPeoplePickerChange}
                  webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                  styles={{ root: { width: '48%', marginRight: '4%' } }}
                  required
                />
              )}
              <TextField
                label="Due date"
                type='date'
                value={dueDate1}
                onChange={this._onDueDate1Change}
                styles={{ root: { width: '48%' } }}
                required
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              {groupIds.length > 0 && (
                <PeoplePicker
                  context={props.context as any}
                  titleText="Verifier"
                  personSelectionLimit={1}
                  groupId={groupIds}
                  groupName={""}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  onChange={this._onVerifierPeoplePickerChange}
                  webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                  styles={{ root: { width: '48%', marginRight: '4%' } }}
                  required
                />
              )}
              <TextField
                label="Due date"
                type='date'
                value={dueDate2}
                onChange={this._onDueDate2Change}
                styles={{ root: { width: '48%' } }}
                required
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              {groupIds.length > 0 && (
                <PeoplePicker
                  context={props.context as any}
                  titleText="Validator"
                  personSelectionLimit={1}
                  groupId={groupIds}
                  groupName={""}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  onChange={this._onValidatorPeoplePickerChange}
                  webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                  styles={{ root: { width: '48%', marginRight: '4%' } }}
                  required
                />
              )}
              <TextField
                label="Due date"
                type='date'
                value={dueDate3}
                onChange={this._onDueDate3Change}
                styles={{ root: { width: '48%' } }}
                required
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              {groupIds.length > 0 && (
                <PeoplePicker
                  context={props.context as any}
                  titleText="Approver"
                  personSelectionLimit={1}
                  groupId={groupIds}
                  groupName={""}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  onChange={this._onApproverPeoplePickerChange}
                  webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                  styles={{ root: { width: '48%', marginRight: '4%' } }}
                  required
                />
              )}
              <TextField
                label="Due date"
                type='date'
                value={dueDate4}
                onChange={this._onDueDate4Change}
                styles={{ root: { width: '48%' } }}
                required
              />
            </div>
            {isEsignEnabled && (
              <Checkbox
                label="eSign"
                checked={eSign}
                onChange={this._onESignChange}
                styles={{ root: { marginBottom: '20px' } }}
                required
              />
            )}
          </div>
        }
        {
          this.state.isProcessing &&
          <div>
            <h4>
              Processing, Please wait...
            </h4>
          </div>
        }

        <DialogFooter>
          {this.state.isProcessing ?
            <div>
            </div>
            :
            <div>
              <PrimaryButton onClick={this._startWorkflow} text="Start" 
               className={styles.saveButton}
              />
              <DefaultButton onClick={props.onClose} text="Cancel" />
            </div>
          }
        </DialogFooter>
      </Dialog>
    );
  }
}

export default WorkflowDialogContent;