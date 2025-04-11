import * as React from "react";
import { CommandBar, ICommandBarItemProps } from "@fluentui/react/lib/CommandBar";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import * as strings from "BaselineFormsFormCustomizerStrings";
import { CreateNewVersion } from "../CreateNewVersion/CreateNewVersion";
import { ICustomSearchResults, ILinkedItems,IRelatedBaselines, ITechnicalEvent,IBaseLineExportEvent, IBaselineEvent, IBaselineExport, } from "../../../../interfaces/IGlobalInterfaces";
import {
  AadHttpClient
} from "@microsoft/sp-http";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { IFreezeBaselineRequest } from "../../../../interfaces/IFreezeBaselineRequest";
import { INotificationService } from "../../../../service/NotificationService";
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../../pnpjs-config";
import { BaselineContentType, SearchConstants, TechnicalDocumentContentType } from "../../../../constants";
import { ISearchResult } from "@pnp/sp/search";
import { DmsRole, IPermissionsService } from "../../../../service/PermissionsService";
import { Dialog, DialogType, ProgressIndicator, DialogFooter, PrimaryButton } from "@fluentui/react";
import CONSTANTS from "../../../../_Constants";
import BaselineService from "../../../../service/BaselineService";
export interface IBaselineCommandBarProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  itemID: number;
  item: ILinkedItems;
  linkedSearchItems: ILinkedItems[];
  checkPendingValue: boolean;
  isCheckLikedItem: boolean;
  onClose: () => void;
  onSave: () => void;
  freezeEndpoint: string;
  dmsClient: AadHttpClient;
  notificationService: INotificationService;
  permissionsService: IPermissionsService;
}
export interface IBaselineCommandBarState {
  isFreezeDialogVisible: boolean;
  isSuccessMsg: boolean;
  isNewVersion: boolean;
  loading: boolean;
  dialogMsg: string;
  isValidUser:boolean;
  showTriggerDialog:boolean;
  title: string,
  baselineStatus: string,
  baselineVersion: string,
  cRID: string,
  code: string,
  label: string,
  baselineComment: string,
  dialogMessage:string
}
export class BaselineCommanBar extends React.Component<
  IBaselineCommandBarProps,
  IBaselineCommandBarState
> {
  private exportedBaselines: IBaselineExport[] = [];
  private _sp: SPFI;
  constructor(props: IBaselineCommandBarProps) {
    super(props);

    this.state = {
      isFreezeDialogVisible: false,
      isSuccessMsg: false,
      isNewVersion: false,
      loading: true,
      dialogMsg: '',
      isValidUser:false,
      showTriggerDialog:false,
      title: "",
      baselineStatus: "",
      baselineVersion: "",
      cRID: "",
      code: "",
      label: "",
      baselineComment: "",
      dialogMessage: "Please wait... preparing to begin the process..",
    };
    this.exportedBaselines = [];
  }

  private _newVersion(): void {
    this.setState((preState) => ({ isNewVersion: !preState.isNewVersion }));
  }

  public async componentDidMount() {
    this._sp = getSP(this.props.context);
    await this.loadPanel();
  }
  private async loadPanel() {
    const contributor=await this.props.permissionsService.currentUserHasRole(DmsRole.Contributor);
   const prjAdmin=await this.props.permissionsService.currentUserHasRole(DmsRole.ProjectAdmin);
      this.setState({
        isValidUser: contributor||prjAdmin
      });
    
      console.log("this",this.state.isValidUser);
      
  }

  private async validateTechnicalDocuments(items: ILinkedItems[]): Promise<{
    item: ICustomSearchResults;
    issue: string | undefined;
  }[]> {
    //Search by doc ID, split into multiple queries of 2048 chars (SP limit)
    const searchQueries = items.reduce((acc, item) => {
      let query = '';
      switch (item.ContentType) {
        case "Baseline":
          query = `(${SearchConstants.ItemIdManagedPropName}:${item.ChildBaselineId} AND ${SearchConstants.SitePathManagedPropName}:${item.SiteUrl} AND ${SearchConstants.ContentTypeManagedPropName}:${BaselineContentType})`;
          break;
        case "Technical Document":
          query = `${SearchConstants.DocumentIdManagedPropName}:${item.DocumentId}`;
          break;
        default:
          throw new Error(strings.UnknownContentTypeIssue);
      }
      const curretAccQuery = acc[acc.length - 1];
      if (!curretAccQuery || curretAccQuery.length + query.length + 4 > 2048) {
        acc.push(query);
      } else {
        acc[acc.length - 1] = `${curretAccQuery} OR ${query}`;
      }
      return acc;
    }, [] as string[]);
    const [batch, execute] = this._sp.batched();
    const AllQueriesPromise = Promise.all(searchQueries.map(query =>
      batch.search({
        Querytext: query,
        RowLimit: 500,
        TrimDuplicates: false, ClientType: 'ContentSearchRegular',
        SelectProperties: ["DlcDocId", "DocId", "ListItemID", "ContentType", "ProjectReferenceOWSTEXT", "ProjectRevisionOWSTEXT", "CRIDOWSTEXT", "ChangeRequestIDOWSTEXT", "CodeOWSTEXT", "Title", "SitePath", "LabelOWSTEXT", "BaselineStatusOWSTEXT", "DocumentStatusOWSCHCS", "DocumentLink", "Path"]
      }).then((response) => response.PrimarySearchResults.map((r: ISearchResult & { [prop: string]: unknown }) =>
      ({
        DocumentId: r.DlcDocId as string,
        ItemId: r.ListItemID as number,
        ProjectReference: r.ProjectReferenceOWSTEXT as string,
        ProjectRevision: r.ProjectRevisionOWSTEXT as string,
        CRID: r.CRIDOWSTEXT as string,
        ChangeRequestT: r.ChangeRequestIDOWSTEXT as string,
        Code: r.CodeOWSTEXT as string,
        Title: r.Title as string,
        SitePath: r.SitePath as string,
        isSelected: false,
        DocId: `${r.DocId}`,
        ContentType: r.ContentType as string,
        Label: r.LabelOWSTEXT as string,
        BaselineStatus: r.BaselineStatusOWSTEXT as string,
        DocumentStatus: r.DocumentStatusOWSCHCS as string,
        DocumentLink: r.Path,
        index: 0
      }
      ))),
    ));
    await execute();
    const results = (await AllQueriesPromise).reduce((acc, result) => acc.concat(result), []);
    return results.map((result: ICustomSearchResults) => {
      switch (result.ContentType) {
        case BaselineContentType:
          return {
            item: result,
            issue: result.BaselineStatus?.toLocaleLowerCase() === "frozen" ? undefined : strings.BaselineIssue
          };
        case TechnicalDocumentContentType:
          return {
            item: result,
            issue: result.DocumentStatus?.toLocaleLowerCase() === "applicable" ? undefined : strings.TechnicalDocIssue
          };
        default:
          return {
            item: result,
            issue: strings.UnknownContentTypeIssue
          };
      }
    });
  }

  private async _processInputData(
    inputData: ILinkedItems[]
  ): Promise<void> {

    const { dmsClient, freezeEndpoint, context, itemID, notificationService } = this.props;

    const checkDialogId = notificationService.addOngoingOperation(strings.FreezeCheckTitle);
    try {
      const checks = await this.validateTechnicalDocuments(inputData);
      if (!checks.every((check) => !check.issue)) {
        checks.reduce((acc, check) => {
          if (check.issue) {
            acc.push({ message: check.issue, count: acc.filter((c) => c.message === check.issue).length + 1 });
          }
          return acc;
        }, [] as { message: string, count: number }[]);
        notificationService.setOperationCompletedSilent(checkDialogId);
        notificationService.notifyWithDetails(strings.FreezeNotificationTitle, `${strings.UnableToFreezeBaseline}`,
          checks.filter(c => !!c.issue).map((c) => `${c.item.ProjectReference || c.item.Title || c.item.Label}${c.item.ProjectRevision}: ${c.issue}`),
          "error");
        return;
      }
    } catch (error) {
      notificationService.setOperationCompleted(checkDialogId, false, strings.FreezeCheckError);
      console.error("ERROR", error);
      return;
    }

    const dialogId = notificationService.addOngoingOperation(strings.FreezeNotificationTitle);
    try {
      const response = await dmsClient.post(freezeEndpoint, AadHttpClient.configurations.v1, {
        body: JSON.stringify({
          baselineId: itemID,
          listId: context.list?.guid?.toString(),
          siteUrl: context.pageContext.web.serverRelativeUrl
        } as IFreezeBaselineRequest),
      });
      if (response.ok) {
        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/Lists/Baselines/AllItems.aspx`;
      } else {
        notificationService.setOperationCompleted(dialogId, false, strings.FreezeError);
        console.error("ERROR", response.statusText);
      }
    } catch (error) {
      notificationService.setOperationCompleted(dialogId, false, strings.FreezeError);
      console.error("ERROR", error);
    }
  }
  //Download Baselines

  public getRelativePath(url?: string) {
    const removePart = this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, "");
    return url?.replace(removePart, "");
  }

  private getBaselineData = async (baselineId: number, url: string) => {
    try {
      console.log(`Fetching baseline data for ID: ${baselineId}`);
      const baslineList = await BaselineService.getReferenceListItem(CONSTANTS.ListNames.BaselineRefernceList, baselineId.toString(), url);
      const technicalList = await BaselineService.getReferenceListItem(CONSTANTS.ListNames.TechnicalReferenceList, baselineId.toString(), url);

      const baselineDocs: IBaselineEvent[] = [];
      await Promise.all(baslineList.map(async (b) => {
        if (this.exportedBaselines.filter(rec => rec.ItemId === parseInt(b[CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID])).length === 0)
          baselineDocs.push({
            SitePath: await this.getRelativePath(b[CONSTANTS.BaselineRefernceListFieldNames.SiteURL]),
            ItemId: b[CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID],
            Title: b[CONSTANTS.BaselineRefernceListFieldNames.Title],
            Status: b[CONSTANTS.BaselineRefernceListFieldNames.Status],
            relatedBaseline: await this.getBaselineData(b[CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID], b[CONSTANTS.BaselineRefernceListFieldNames.SiteURL])
          });
      }));
      console.log(`Baseline : ${baselineId}-${technicalList}`);
      const techDocs: ITechnicalEvent[] = [];
      const res = await BaselineService.getSearchBatchResults(technicalList);
      res.map(result =>{
        techDocs.push({
          DocumentId: result.DocumentId,
          SitePath: this.getRelativePath(result.SitePath),
          DocumentLink: this.getRelativePath(result.DocumentLink),
          Status: result.DocumentStatus,
          ProjcectReference: result.ProjectReference,
          ProjcectRevision: result.ProjectRevision,
          FileName: result.Title
        });
      })


      // await Promise.all(technicalList.map(async (t) => {
      //   const result = await BaselineService.getSearchResults(`DlcDocId=${t[CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]}`);
      //   if (result.length > 0) {
      //     techDocs.push({
      //       DocumentId: result[0].DocumentId,
      //       SitePath: this.getRelativePath(result[0].SitePath),
      //       DocumentLink: this.getRelativePath(result[0].DocumentLink),
      //       Status: result[0].DocumentStatus,
      //       ProjcectReference: result[0].ProjectReference,
      //       ProjcectRevision: result[0].ProjectRevision,
      //       FileName: result[0].Title
      //     });
      //   }
      // }));

      console.log(`Baseline and technical documents fetched successfully for baselineId: ${baselineId}`);
      return { baseline: baselineDocs, technical: techDocs };
    } catch (error) {
      console.error(`Error fetching baseline data for baselineId ${baselineId}:`, error);
      throw error;
    }
  }

  public exportBaseline = async () => {
    this.setState({
      showTriggerDialog: true,
      dialogMessage: "Please wait...preparing to begin the process"
    });

    try {
      this.exportedBaselines.push({ ItemId: this.props.itemID });
      const baselineDocuments: IRelatedBaselines = await this.getBaselineData(this.props.itemID, this.props.context.pageContext.site.absoluteUrl);

      const rootBasseline: IBaselineEvent = {
        ItemId: this.props.itemID,
        SitePath: this.props.context.pageContext.site.serverRelativeUrl,
        Title: this.state.title + " - " + this.state.baselineVersion,
        Status: this.state.baselineStatus,
        Code:this.state.code,
        BaselineVersion:this.state.baselineVersion
      };

      const messagePayLoad: IBaseLineExportEvent = {
        ExportDocuments: baselineDocuments,
        UserEmail: this.props.context.pageContext.user.email,
        SiteUrl: this.props.context.pageContext.site.serverRelativeUrl,
        RootBaseLine: rootBasseline
      };
      console.log('Sending baseline export request:', messagePayLoad);
      const result = await BaselineService.sendmessage(messagePayLoad);

      console.log('Export result:', result);
      if (result === "OK") {
        this.setState({ dialogMessage: "Your baseline download is being prepared. You will shortly receive an email with your download link!" });
      } else {
        console.error('Error during baseline export:', result);
        this.setState({ dialogMessage: `Some error occurred, please contact Admin with below details, Error details: ${result}` });
      }
    } catch (error) {
      console.error('Error in exportBaseline:', error);
      this.setState({ dialogMessage: 'An unexpected error occurred during the export. Please try again later.' });
    }
  }



  public render(): React.ReactElement<{}> {
    const { item, checkPendingValue, isCheckLikedItem} = this.props;
    let items: ICommandBarItemProps[];
    

    if (item) {
      const isFreezeDisabled = item.BaselineStatus !== "Ongoing" || !checkPendingValue || !isCheckLikedItem;
      const isNewVersionDisabled = item.BaselineStatus !== "Frozen";

      items = [
        {
          key: "newVersion",
          text: "New Version",
          iconProps: { iconName: "Add" },
          onClick: () => this._newVersion(),
          disabled: isNewVersionDisabled,
        },
        ...(this.state.isValidUser ? [{
          key: "freeze",
          text: "Freeze",
          iconProps: { iconName: "Lock" },
          onClick: () => this._processInputData(this.props.linkedSearchItems),
          disabled: isFreezeDisabled,
      }] : []),
        {
          key: "download",
          text: strings.DownloadBaseline,
          iconProps: { iconName: "Download" },
          onClick: () =>  this.exportBaseline(),
          disabled: !checkPendingValue || !isCheckLikedItem,
        },
        ...(this.props.displayMode === FormDisplayMode.Edit ? [
          {
            key: "save",
            text: "Save",
            iconProps: { iconName: "Save" },
            onClick: () => this.props.onSave(),
          },
        ] : []),
        {
          key: "cancel",
          text: "Cancel",
          iconProps: { iconName: "Cancel" },
          onClick: () => this.props.onClose(),
        },
      ];
    } else {
      items = [
        {
          key: "save",
          text: "Save",
          iconProps: { iconName: "Save" },
          onClick: () => this.props.onSave(),
        },
        {
          key: "cancel",
          text: "Cancel",
          iconProps: { iconName: "Cancel" },
          onClick: () => this.props.onClose(),
        },
      ];
    }


    return (
      <>
        {this.state.isNewVersion && (
          <CreateNewVersion
            item={this.props.item}
            context={this.props.context}
            isNewVersion={this.state.isNewVersion}
            updateVersion={this._newVersion.bind(this)}
          />
        )}

{this.state.showTriggerDialog &&
          <Dialog
            hidden={!this.state.showTriggerDialog}
            onDismiss={() => { this.setState({ showTriggerDialog: false }) }}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Download Baseline",
            }}
            modalProps={{
              isBlocking: true,
              containerClassName: "ms-dialogMainOverride",
            }}
          >
            {this.state.dialogMessage}
            <p />
            {this.state.dialogMessage.indexOf("Please wait") === 0 &&
              <ProgressIndicator label={""} description="" />
            }
            <DialogFooter>
              <PrimaryButton
                onClick={() => { this.setState({ showTriggerDialog: false }) }}
                text="OK"
              />
            </DialogFooter>
          </Dialog>

        }
        <CommandBar items={items} />
      </>
    );
  }
}
