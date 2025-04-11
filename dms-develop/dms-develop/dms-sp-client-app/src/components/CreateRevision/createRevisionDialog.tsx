import { Log } from "@microsoft/sp-core-library";
import * as React from "react";
import * as strings from "CreateRevisionCommandSetStrings";
import {
  Label, DefaultButton, Spinner, Link, Separator
} from "@fluentui/react";
import { IStackItemStyles, Stack } from "@fluentui/react/lib/Stack";
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/folder";
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { IFolder } from "@pnp/sp/folders";
import { IListItemFormUpdateValue, RenderListDataOptions } from "@pnp/sp/lists";
import { ListUrls } from "../../constants";
//import styles from "./createRevisionDialog.module.scss";


const LOG_SOURCE: string = 'DMS Create Revision';


export interface IRevisionDialogProps {
  close: () => void;
  projectReference: string;
  sp: SPFI;
  documentSetRelativeUrl: string;
  showDialog: boolean;
}

export interface IRevisionDialogState {
  showDialog: boolean;
  showSuccess: boolean;
  showError: boolean;
  isValidating: boolean;
  isProcessing: boolean;
  errorMessage: string;
  newDraftLink: string;
}

export default class CreateRevisionDialog extends React.Component<IRevisionDialogProps, IRevisionDialogState> {

  constructor(props: IRevisionDialogProps) {
    super(props);

    this.state = {
      showDialog: this.props.showDialog,
      showSuccess: false,
      showError: false,
      isValidating: false,
      isProcessing: false,
      errorMessage: "",
      newDraftLink: ""
    }
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: createRevisionDialog unmounted');
  }

  public async componentDidMount(): Promise<void> {
    Log.info(LOG_SOURCE, 'React Element: createRevisionDialog componentdidmount');
    try {
      const draftExist = await this.checkIfDraftExist();
      if (draftExist) {
        this.setState({ isValidating: false, showError: true, errorMessage: strings.CreateRevisionErrorDraftExistMessage });
        return;
      }

      //proceed with draft creation
      const newDraft = await this.createDraft();
      await this.updateDocSetMetadata(newDraft.newDocSet);
      this.setState({ isValidating: false, isProcessing: false, showSuccess: true, showError: false, errorMessage: "", newDraftLink: newDraft.docSetPath });
    } catch (error) {
      console.log(error);
      this.setState({ isValidating: false, isProcessing: false, errorMessage: error.message, showError: true, showSuccess: false });
    }
  }

  public render(): React.ReactElement<{}> {

    const { showError, showSuccess, isValidating, isProcessing, errorMessage, newDraftLink } = this.state;

    const stackItemStyles: IStackItemStyles = {
      root: {
        paddingLeft: 20,
        paddingRight: 20
      },
    };

    return (<div>

      {
        <Stack verticalFill style={{ borderRadius: 5 }}>
          {
            <Stack.Item styles={stackItemStyles} style={{ paddingTop: 20 }}>
              <Label style={{ fontWeight: "bold" }}>
                {showError ? "Error" : showSuccess ? "Success" : strings.DoNotRefresh}
              </Label>
            </Stack.Item>
          }
          {isValidating && !showError && !showSuccess &&
            <Stack.Item styles={stackItemStyles} style={{ paddingTop: 20 }}>
              <Spinner label={strings.CreateRevisionValidatingMessage} ariaLive="assertive" labelPosition="top" />
            </Stack.Item>
          }
          {
            isProcessing && !showError && !showSuccess &&
            <Stack.Item styles={stackItemStyles} style={{ paddingTop: 20 }}>
              <Spinner label={strings.CreateRevisionProcessingMessage} ariaLive="assertive" labelPosition="top" />
            </Stack.Item>
          }
          {
            showError &&
            <Stack.Item styles={stackItemStyles} style={{ paddingTop: 20 }}>
              <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>
            </Stack.Item>
          }
          {
            showSuccess &&
            <Stack.Item styles={stackItemStyles} style={{ paddingTop: 20 }}>
              <MessageBar messageBarType={MessageBarType.success}>{strings.CreateRevisionSuccessMessage} <Link href={newDraftLink}>Click here</Link> to open.</MessageBar>
            </Stack.Item>
          }
          <Stack.Item style={{ paddingTop: 20 }}>
            <Separator />
          </Stack.Item>
          <Stack.Item align="center" styles={stackItemStyles} style={{ paddingBottom: 20 }}>
            <DefaultButton onClick={() => {
              this.setState({
                showDialog: false,
                showSuccess: false,
                showError: false,
                isValidating: false,
                isProcessing: false,
                errorMessage: ""
              }); this.props.close();
            }} text={"Close"} disabled={!(showError || showSuccess)} />
          </Stack.Item>
        </Stack>
      }
    </div>
    );
  }

  private async checkIfDraftExist(): Promise<boolean> {
    const { sp, projectReference } = this.props;
    this.setState({ isValidating: true });
    const camlQuery = `<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="ProjectReference"/><Value Type="Text">${projectReference}</Value></Eq></Where></Query></View>`;
    const items = await sp.web.lists
      .getByTitle(ListUrls.Draft)
      .renderListDataAsStream({
        ViewXml: camlQuery,
        RenderOptions: RenderListDataOptions.ListData,
      });
    return ((items?.Row?.length || 0) > 0);
  }

  private async createDraft(): Promise<{ newDocSet: IFolder, docSetPath: string }> {
    const { sp, documentSetRelativeUrl } = this.props;
    this.setState({ isValidating: false, isProcessing: true });
    //get current document set
    const docSetServerRelativeUrl = documentSetRelativeUrl;
    const docSet = sp.web.getFolderByServerRelativePath(docSetServerRelativeUrl);

    //get target folder path
    const draftRootFolder = await sp.web.lists.getByTitle(ListUrls.Draft).rootFolder();
    const currentName = docSetServerRelativeUrl.split("/").pop();
    const currentRevision = currentName?.split("-").pop();
    const newRevision = this.getNewProjectRevision(currentRevision || "");
    const newName = currentName?.split("-").slice(0, -1).join("-") + "-" + newRevision;
    const targetFolder = draftRootFolder.ServerRelativeUrl + '/' + newName;

    const newDocset = await docSet.copyByPath(targetFolder, false);

    return { newDocSet: newDocset, docSetPath: targetFolder };
  }

  private async updateDocSetMetadata(newDocSet: IFolder): Promise<void> {
    const { sp } = this.props;
    const allFields = await newDocSet.listItemAllFields();
    const formValues = [
      { 'FieldName': 'DocumentStatus', 'FieldValue': 'Draft' },
      { 'FieldName': 'ActualSubmissionDate', 'FieldValue': '' },
      { 'FieldName': 'ProjectRevision', 'FieldValue': this.getNewProjectRevision(allFields.ProjectRevision) },
      { 'FieldName': 'ApprovalCycleLaunchDate', 'FieldValue': '' },
      { 'FieldName': 'Verifiedby', 'FieldValue': '' },
      { 'FieldName': 'VerificationDate', 'FieldValue': '' },
      { 'FieldName': 'Validatedby', 'FieldValue': '' },
      { 'FieldName': 'ValidationDate', 'FieldValue': '' },
      { 'FieldName': 'Approvedby', 'FieldValue': '' },
      { 'FieldName': 'ApprovalDate', 'FieldValue': '' },
    ];
    const result: IListItemFormUpdateValue[] = await sp.web.lists.getByTitle(ListUrls.Draft).items.getById(allFields.Id).validateUpdateListItem(formValues, true);
    console.log(result);
    if (result && result.length > 0) {
      const errorFields = result.filter(f => f.HasException).map(f => { return { fieldName: f.FieldName, errorMessage: f.ErrorMessage } });
      if (errorFields.length > 0) {
        console.log(errorFields);
        await sp.web.lists.getByTitle(ListUrls.Draft).items.getById(allFields.Id).delete();
        throw new Error(strings.CreateRevisionErrorUnknownMessage);
      }
    }
  }

  private getNewProjectRevision(currentRevisionNumber: string): string {
    const arrRevision = currentRevisionNumber.split('');
    //if current Revision is empty, return empty
    if (arrRevision.length == 0) {
      return currentRevisionNumber;
    }

    //check if invalid Revision format, return empty
    const invalidCodes = arrRevision.filter(c => !(c.charCodeAt(0) >= 65 && c.charCodeAt(0) <= 90));
    if (invalidCodes.length > 0) {
      return currentRevisionNumber;
    }

    const newRevision = [];
    let increment = 1;
    while (arrRevision.length > 0) {
      let c = arrRevision.pop();
      if (c) {
        let ascii = c.charCodeAt(0) + increment;
        if (ascii > 90) {
          newRevision.push('A');
          increment = 1;
        } else {
          newRevision.push(String.fromCharCode(ascii));
          increment = 0;
        }
      }
    }
    if (increment > 0) {
      return ['A', ...newRevision].reverse().join('');
    }
    return newRevision.reverse().join('');
  }
}