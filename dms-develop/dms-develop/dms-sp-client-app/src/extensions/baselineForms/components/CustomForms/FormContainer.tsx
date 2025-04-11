import * as React from "react";
import styles from "../BaselineForms.module.scss";
import { TooltipDelay, TooltipHost } from "@fluentui/react/lib/Tooltip";
import { TextField } from "@fluentui/react/lib/TextField";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { Icon } from "@fluentui/react/lib/Icon";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import RelatedItems from "../RelatedItems/RelatedItems";
import CONSTANTS from "../../../../_Constants";
import BaselineService from "../../../../service/BaselineService";
import {
  ICustomSearchResults,
  ILinkedItems,
  IRequiredFields,
} from "../../../../interfaces/IGlobalInterfaces";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { AgGridReact } from "ag-grid-react";
import { GridApi, GridReadyEvent, RowSelectedEvent } from "ag-grid-community";
import "ag-grid-community/dist/styles/ag-grid.css";
import "ag-grid-community/dist/styles/ag-theme-alpine.css";
import Dialog, { DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import UserMessage from "../controls/UserMessage";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { BaselineHelper } from "../../../../service/BaselineHelper";
import { BaselineCommanBar } from "../FreezeBaseline/BaselineCommanBar";
import { AadHttpClient } from "@microsoft/sp-http";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { INotificationService } from "../../../../service/NotificationService";
import strings from "BaselineFormsFormCustomizerStrings";
import { IPermissionsService } from "../../../../service/PermissionsService";

export interface IFormContainerState {
  openSearchDialog: boolean;
  openUnLinkDialog: boolean;
  title: string;
  baselineStatus: string;
  baselineVersion: string;
  cRID: string;
  code: string;
  label: string;
  baselineComment: string;
  errorMessage: string;
  errorControls: { stateName: string; fieldName: string; }[];
  linkedSearchItems: ILinkedItems[];
  technicalIds: number[];
  baselineIDs: number[];
  isPageLoading: boolean;
  isSaving: boolean;
  gridAPI?: GridApi;
  updatedRows: ILinkedItems[];
  currentMode: string;
  checkPendingValue: boolean;
  isCheckLikedItem: boolean;
  currentPresaveOperation?: {
    label: string;
    status: "running" | "success" | "error";
    index: number;
    count: number;
  };
}

export interface IFormContainerProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
  itemID: number;
  item: ILinkedItems;
  freezeEndpoint: string;
  dmsClient: AadHttpClient;
  notificationService: INotificationService;
  permissionsService:IPermissionsService;
}

export default class FormContainer extends React.Component<
  IFormContainerProps,
  IFormContainerState
> {
  constructor(props: IFormContainerProps) {
    super(props);
    this.state = {
      openSearchDialog: false,
      openUnLinkDialog: false,
      title: "",
      baselineStatus: "Ongoing",
      baselineVersion: "",
      cRID: "",
      code: "",
      label: "",
      baselineComment: "",
      errorMessage: "",
      errorControls: [],
      linkedSearchItems: [],
      technicalIds: [],
      baselineIDs: [],
      isPageLoading: true,
      isSaving: false,
      gridAPI: undefined,
      updatedRows: [],
      currentMode: "New",
      checkPendingValue: true,
      isCheckLikedItem: true,
    };
  }

  public async componentDidMount(): Promise<void> {
    const { item, itemID } = this.props;
    this.setState({ isPageLoading: false });
    if (itemID !== undefined) {
      this.setState({ currentMode: "Edit" });
      const baslineList = await BaselineService.getReferenceListItem(
        CONSTANTS.ListNames.BaselineRefernceList,
        itemID.toString(),
        this.props.context.pageContext.site.absoluteUrl
      );
      const technicalList = await BaselineService.getReferenceListItem(
        CONSTANTS.ListNames.TechnicalReferenceList,
        itemID.toString(),
        this.props.context.pageContext.site.absoluteUrl
      );
      const linkedSearchItems: ILinkedItems[] = [];
      let index = 0;
      baslineList.map((b) => {
        linkedSearchItems.push({
          DocumentId: null,
          SiteUrl: b[CONSTANTS.BaselineRefernceListFieldNames.SiteURL],
          ChildBaselineId:
            b[CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID],
          ContentType: CONSTANTS.ContentTypeNames.Baseline,
          isSelected: false,
          Title: b[CONSTANTS.BaselineRefernceListFieldNames.Title],
          ProjectReference:b[CONSTANTS.BaselineListFieldNames.Code],
          ProjectRevision:b[CONSTANTS.BaselineListFieldNames.BaselineVersion],
          index: index++,
          Status: b[CONSTANTS.BaselineRefernceListFieldNames.Status],
        });
      });
      await Promise.all(
        technicalList.map(async (t) => {
          const result = await BaselineService.getSearchResults(
            `DlcDocId:${t[CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]
            }`
          );
          if (result.length > 0)
            linkedSearchItems.push({
              DocumentId: result[0].DocumentId,
              SiteUrl: result[0].SitePath,
              ChildBaselineId: null,
              ContentType: result[0].ContentType,
              isSelected: false,
              Title: result[0].Title,
              index: index++,
              Status: result[0].DocumentStatus,
              ProjectReference: result[0].ProjectReference,
              ProjectRevision: result[0].ProjectRevision,
            });
        })
      );
      this.setState(
        {
          title: item[
            CONSTANTS.BaselineListFieldNames.Title as keyof ILinkedItems
          ] as string,
          baselineStatus: item[
            CONSTANTS.BaselineListFieldNames
              .BaselineStatus as keyof ILinkedItems
          ] as string,
          baselineVersion: item[
            CONSTANTS.BaselineListFieldNames
              .BaselineVersion as keyof ILinkedItems
          ] as string,
          cRID: item[
            CONSTANTS.BaselineListFieldNames.CRID as keyof ILinkedItems
          ] as string,
          code: item[
            CONSTANTS.BaselineListFieldNames.Code as keyof ILinkedItems
          ] as string,
          label: item[
            CONSTANTS.BaselineListFieldNames.Title as keyof ILinkedItems
          ] as string,
          baselineComment: item[
            CONSTANTS.BaselineListFieldNames
              .BaselineComments as keyof ILinkedItems
          ] as string,
          linkedSearchItems: linkedSearchItems,
          technicalIds: technicalList.map(({ Id }) => Id),
          baselineIDs: baslineList.map(({ Id }) => Id),
        },
        () => {
          this.state.gridAPI?.sizeColumnsToFit();
        }
      );
    }
  }

  public gridReadyRowResize(event: GridReadyEvent): void {
    const api = event.api;
    setTimeout(() => {
      api.sizeColumnsToFit();
    }, 1500);
    this.setState({ gridAPI: api });
  }

  private onSelectionChanged(_: RowSelectedEvent): void {
    const selectedRows = this.state.gridAPI?.getSelectedRows();
    const rows = cloneDeep(this.state.linkedSearchItems);
    rows?.map((r) => {
      const exists = selectedRows?.some((srch) => srch.index === r.index);
      if (exists) r.isSelected = true;
      else r.isSelected = false;
    });
    this.setState({ updatedRows: rows });
  }

  //RESET ERROR CONTROLS
  public resetError(field: string): void {
    const errors = this.state.errorControls;
    const index = errors.findIndex((obj) => obj.stateName === field);
    if (index >= 0) {
      errors.splice(index, 1);
      this.setState({
        errorControls: errors,
        errorMessage: errors.length > 0 ? this.state.errorMessage : "",
      });
    }
  }

  // TEXTFIELDS ON CHANGE
  private onTextChange(field: keyof IFormContainerState, value: string): void {
    const newState = { [field]: value } as unknown as Pick<IFormContainerState, keyof IFormContainerState>;
    this.setState(newState);
    const errors = this.state.errorControls;
    const index = errors.findIndex((obj) => obj.stateName === field);
    if (index >= 0) {
      errors.splice(index, 1);
      this.setState({
        errorControls: errors,
        errorMessage: errors.length > 0 ? this.state.errorMessage : "",
      });
    }
    this.resetError(field);
  }

  public async validateFields(): Promise<void> {
    const requiredFields: IRequiredFields[] = [
      { stateName: "baselineStatus", fieldName: "Baseline Status" },
      { stateName: "baselineVersion", fieldName: "Baseline Version" },
      { stateName: "code", fieldName: "Code" },
      { stateName: "label", fieldName: "Label" },
    ];
    // Required field validation
    try {
      BaselineHelper.checkRequiredFields(this.state, requiredFields);
    } catch (e) {
      this.setState({
        errorControls: e.targetFields,
        errorMessage: e.message,
      });
      throw new Error(e.message);
    }
    return; // No validation issues.
  }

  //ON FORM SUBMIT
  private async submitForm(): Promise<void> {
    const { notificationService } = this.props;
    await this.validateFields();
    await this.validateVersions();
    
    this.setState({ isSaving: true });
    const notif = notificationService.addOngoingOperation(strings.CreatingItem);
    try {
      const createObject = {
        BaselineComments: this.state.baselineComment,
        BaselineStatus: this.state.baselineStatus,
        BaselineVersion: this.state.baselineVersion,
        Code: this.state.code,
        CRID: this.state.cRID,
        Title: this.state.label,
        ContentTypeId: await BaselineService.getBaseLineContentTypeId(),
      };
      const res = await BaselineService.createListItem(
        this.props.context.list.title,
        createObject
      );
      const newLItemId = res.Id.toString();
      await Promise.all(this.state.linkedSearchItems.map(async (data) => {
        if (data.ContentType === CONSTANTS.ContentTypeNames.Baseline) {
          const baslineObject = {
            [CONSTANTS.BaselineRefernceListFieldNames.BaselineParentID]: newLItemId,
            [CONSTANTS.BaselineRefernceListFieldNames.SiteURL]: data.SiteUrl,
            [CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID]: data.ChildBaselineId,
            Title: data.Title
          };
          await BaselineService.createListItem(
            CONSTANTS.ListNames.BaselineRefernceList,
            baslineObject
          );
        }
        if (data.ContentType === CONSTANTS.ContentTypeNames.Technical) {
          const technicalObject = {
            [CONSTANTS.TechnicalRefernceListFieldNames.BaselineParentID]:
              newLItemId,
            [CONSTANTS.TechnicalRefernceListFieldNames.SiteURL]: data.SiteUrl,
            [CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]:
              data.DocumentId,
            Title: data.Title,
          };
          await BaselineService.createListItem(
            CONSTANTS.ListNames.TechnicalReferenceList,
            technicalObject
          );
        }
      }));
      notificationService.setOperationCompleted(notif, true);
      this.props.onSave();
    } catch (err) {
      this.setState({ isSaving: false });
      console.error(err);
      notificationService.setOperationCompleted(notif, false, err.message);
    }
  }

  private async UpdateForm(): Promise<void> {
    try {
      await this.validateFields();
      await this.validateVersions();


    } catch (e) {
      console.log(e);
      return;
    }

    this.setState({ isSaving: true });
    const createObject = {
      BaselineComments: this.state.baselineComment,
      BaselineStatus: this.state.baselineStatus,
      BaselineVersion: this.state.baselineVersion,
      Code: this.state.code,
      CRID: this.state.cRID,
      Title: this.state.label,
    };
    await BaselineService.updateListItem(
      this.props.context.list.title,
      createObject,
      this.props.itemID
    )
      .then(async (res) => {
        const newLItemId = this.props.itemID.toString();
        await BaselineService.batchDelete(
          CONSTANTS.ListNames.BaselineRefernceList,
          this.state.baselineIDs
        );

        await BaselineService.batchDelete(
          CONSTANTS.ListNames.TechnicalReferenceList,
          this.state.technicalIds
        );

        const logs = await Promise.all(this.state.linkedSearchItems.map(async (data) => {

          if (data.ContentType === CONSTANTS.ContentTypeNames.Baseline) {
            const baslineObject = {
              [CONSTANTS.BaselineRefernceListFieldNames.BaselineParentID]: newLItemId,
              [CONSTANTS.BaselineRefernceListFieldNames.SiteURL]: data.SiteUrl,
              [CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID]: data.ChildBaselineId,
              Title: data.Title
              // [CONSTANTS.BaselineRefernceListFieldNames.Status]:data.Status
            };

            await BaselineService.createListItem(
              CONSTANTS.ListNames.BaselineRefernceList,
              baslineObject
            );
          }
          if (data.ContentType === CONSTANTS.ContentTypeNames.Technical) {
            const technicalObject = {
              [CONSTANTS.TechnicalRefernceListFieldNames.BaselineParentID]:
                newLItemId,
              [CONSTANTS.TechnicalRefernceListFieldNames.SiteURL]:
                data.SiteUrl,
              [CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]:
                data.DocumentId,
              Title: data.Title,
            };
            await BaselineService.createListItem(
              CONSTANTS.ListNames.TechnicalReferenceList,
              technicalObject
            );
          }
        })
        );
        console.log(logs);
        this.setState({
          title: "",
          baselineStatus: "",
          baselineVersion: "",
          cRID: "",
          code: "",
          label: "",
          baselineComment: "",
          errorMessage: "",
          errorControls: [],
        });
        this.setState({
          currentPresaveOperation: { status: 'success', label: 'Update successful', index: 1, count: 1 }
        });
        this.props.onSave();

        window.location.reload();
      })
      .catch((err) => {
        console.log(err);
        this.setState({ isSaving: false });
        this.setState({
          currentPresaveOperation: { status: 'error', label: err.message, index: 1, count: 1 }
        });
      });
  }

  //ON UNLINK CONFIRMATION, REMOVE THE LINKED REFERENCES
  private unlinkReferences(): void {
    const items = cloneDeep(this.state.updatedRows);
    const result = items.filter((l) => l.isSelected === false);
    result.map((r) => {
      r.isSelected = false;
    });

    this.setState(
      { linkedSearchItems: result, openUnLinkDialog: false, updatedRows: [] },
      () => {
        this.state.gridAPI?.sizeColumnsToFit();
      }
    );
  }

  private updatedLinkedItems(searchResults: ICustomSearchResults[]): void {
    const activeResults = searchResults.filter((r) => r.isSelected === true);
    let linkedSearchItems: ILinkedItems[] = cloneDeep(
      this.state.linkedSearchItems
    );
    let index = linkedSearchItems.length;
    activeResults.map((ar: ICustomSearchResults) => {
      //Check for duplicates
      // const duplicateRecords = linkedSearchItems.filter(item => item.SiteUrl === ar.SitePath && item.DocumentId === ar.DocumentId && item.ChildBaselineId === ar.ItemId);
      const duplicateRecords = linkedSearchItems.filter(
        (item) =>
          (item.ContentType === CONSTANTS.ContentTypeNames.Technical &&
            item.DocumentId === ar.DocumentId) ||
          (item.ContentType === CONSTANTS.ContentTypeNames.Baseline &&
            item.ChildBaselineId === ar.ItemId &&
            item.SiteUrl === ar.SitePath)
      );
      //add only if not duplicates
      if (duplicateRecords.length === 0)
        linkedSearchItems.push({
          // DocumentId: ar.DocumentId,
          // SiteUrl: ar.SitePath,
          // ChildBaselineId: ar.ItemId,
          // ContentType: ar.ContentType,
          // isSelected: false,
          // Title: ar.Title,
          // index: index++
          DocumentId: ar.DocumentId,
          SiteUrl: ar.SitePath,
          ChildBaselineId: ar.ItemId,
          ContentType: ar.ContentType,
          //ProjectReference: ar.ProjectReference,
          ProjectReference: ar.ContentType === CONSTANTS.ContentTypeNames.Technical ? ar.ProjectReference : ar.Code,
          //ProjectRevision: ar.ProjectRevision,
          ProjectRevision: ar.ContentType === CONSTANTS.ContentTypeNames.Technical ? ar.ProjectRevision : ar.BaselineVersion,
          isSelected: false,
          Title: ar.Title,
          index: index++,
          Status: ar.ContentType === CONSTANTS.ContentTypeNames.Technical ? ar.DocumentStatus : ar.BaselineStatus
        });
    });
    if(this.props.displayMode === FormDisplayMode.Edit)
    linkedSearchItems = linkedSearchItems.filter(e=> (e.ContentType ===  CONSTANTS.ContentTypeNames.Baseline && 
                      e.SiteUrl === this.props.context.pageContext.web.absoluteUrl &&
                      e.ChildBaselineId?.toString() !== this.props.itemID.toString())||
                      e.ContentType ===  CONSTANTS.ContentTypeNames.Technical ||
                      e.SiteUrl!==this.props.context.pageContext.web.absoluteUrl)

    this.setState({ linkedSearchItems: linkedSearchItems }, () => {
      this.state.gridAPI?.sizeColumnsToFit();
    });
  }

  public async validateVersions(): Promise<void> {

    const isEmpty = this.state.baselineVersion.trim() === '';
    const hasInvalidChars = !/^[a-zA-Z0-9\-_.]+$/.test(this.state.baselineVersion); // Allow alphanumeric, hyphen, and underscore
    // const hasInvalidVersionCheck = await this._sp.web.lists.getByTitle('Baselines').items.filter(Code eq '${item.Code}').orderBy('BaselineVersion', false)();
    const currentItemId=this.props.itemID;
    const hasInvalidVersionCheck = await BaselineService.getVersionDetail(this.state.code, this.state.baselineVersion,currentItemId);

    let hasInvalidVersion: boolean = false;

    if (hasInvalidVersionCheck && hasInvalidVersionCheck.Row.length > 0) {
      hasInvalidVersion = true;
    }

    const newVersionMessage = isEmpty
      ? 'This field is required'
      : hasInvalidChars
        ? 'Special characters are not allowed'
        : hasInvalidVersion
          ? 'This version already exists'
          : '';
    const requiredFields: IRequiredFields[] = [
      { stateName: "baselineVersion", fieldName: "Baseline Version" }
    ];

if(newVersionMessage.length > 0)
 {   
  this.setState({
    errorControls: requiredFields,
    errorMessage: newVersionMessage,
  })
  // eslint-disable-next-line no-throw-literal
  throw {
      targetFields: requiredFields,
      message: newVersionMessage,
    };   
  } 
    return; // No validation issues.
  }


  public render(): React.ReactElement<IFormContainerProps> {
    const iconClass = mergeStyles({
      margin: "3px 5px 0px 0px",
      fontWeight: "bolder",
      fontSize: "x-large",
      cursor: "pointer",
    });

    const columns = [
      {
        headerName: "TITLE",
        field: "Title",
        headerCheckboxSelection: true,
        checkboxSelection: true,
        tooltipField: "Title",
        flex: 1,
        wrapText: true,
        resizable: true,
        autoHeight: true,
      },
      {
        headerName: "STATUS",
        field: "Status",
        tooltipField: "Status",
        flex: 1,
        resizable: true,
        sortable: true,
      },
      {
        headerName: "REFERENCE",
        field: "ProjectReference",
        tooltipField: "ProjectReference",
        flex: 1,
        resizable: true,
        sortable: true,
      },
      {
        headerName: "REVISION",
        field: "ProjectRevision",

        tooltipField: "ProjectRevision",
        flex: 1,
        resizable: true,
        sortable: true,
      },
      {
        headerName: "SITE URL",
        field: "SiteUrl",
        tooltipField: "SiteUrl",
        flex: 1,
        wrapText: true,
        autoHeight: true,
        minWidth: 300,
        // sortable: true,
      },
    ];

    return (
      <div className={styles.formWrapper}>
        <BaselineCommanBar
          {...this.props}
          linkedSearchItems={this.state.linkedSearchItems}
          checkPendingValue={this.state.checkPendingValue}
          isCheckLikedItem={this.state.isCheckLikedItem}
          onSave={
            this.props.itemID === undefined ? this.submitForm.bind(this) : this.UpdateForm.bind(this)
          }
        />
        <div className={`${styles.formOuterContainer} `}>
          <div className={`${styles.formMidContainer} `}>
            <div>
              {this.state.isPageLoading ? (
                <div>
                  <Spinner
                    size={SpinnerSize.large}
                    defaultValue={"Please wait loading..."}
                  >
                    Please wait loading...
                  </Spinner>
                </div>
              ) : (
                <div className={`${styles.formInnerContainer} `}>
                  {this.state.errorMessage.length > 0 && (
                    <div>
                      <UserMessage
                        type={"error"}
                        message={this.state.errorMessage}
                      />{" "}
                    </div>
                  )}
                  {/* Label & Status*/}
                  <div className={styles.column}>
                    <div className={styles.row}>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost content=" Label" delay={TooltipDelay.long}>
                          <TextField
                            required
                            label="Label"
                            placeholder="Label"
                            value={this.state.label}
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => {
                              this.onTextChange("label", newValue);
                              if (
                                this.props.item &&
                                this.props.item.BaselineStatus === "Ongoing" &&
                                this.props.item.Title === newValue &&
                                this.props.item.CRID === this.state.cRID
                              ) {
                                this.setState({
                                  checkPendingValue: true,
                                });
                              } else {
                                this.setState({
                                  checkPendingValue: false,
                                });
                              }
                            }}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "label"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost content="Code" delay={TooltipDelay.long}>
                          <TextField
                            required
                            value={this.state.code}
                            label="Code"
                            placeholder="Code"
                            disabled={
                              this.state.currentMode === "New" ? false : true
                            }
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => this.onTextChange("code", newValue)}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "code"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost
                          content="Baseline Version"
                          delay={TooltipDelay.long}
                        >
                          <TextField
                            required
                            label="Version"
                            placeholder="Version"
                            value={this.state.baselineVersion}
                            // disabled={
                            //   this.state.currentMode === "New" ? false : true
                            // }
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => this.onTextChange("baselineVersion", newValue)}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "baselineVersion"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost content="CR ID" delay={TooltipDelay.long}>
                          <TextField
                            label="CR ID"
                            placeholder="CR ID"
                            value={this.state.cRID}
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => {
                              this.onTextChange("cRID", newValue);
                              if (
                                this.props.item &&
                                this.props.item.BaselineStatus === "Ongoing" &&
                                this.props.item.CRID === newValue &&
                                this.props.item.Title === this.state.label
                              ) {
                                this.setState({
                                  checkPendingValue: true,
                                });
                              } else {
                                this.setState({
                                  checkPendingValue: false,
                                });
                              }
                            }}
                          />
                        </TooltipHost>
                      </div>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost
                          content="Baseline Status"
                          delay={TooltipDelay.long}
                        >
                          <TextField
                            required
                            value={this.state.baselineStatus}
                            label="Status"
                            placeholder="Baseline Status"
                            disabled={true}
                            //disabled={this.state.currentMode === "New" ? true : false}
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => this.onTextChange("baselineStatus", newValue)}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "baselineStatus"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>
                      <div className={styles.fieldWrapper}>&nbsp;</div>
                    </div>
                  </div>
                  <div className={styles.column}>
                    <div className={styles.row}>
                      <div className={styles.fieldWrapper}>
                        <TooltipHost
                          content="Baseline Comments"
                          delay={TooltipDelay.long}
                        >
                          <TextField
                            value={this.state.baselineComment}
                            label="Comments"
                            placeholder="comments"
                            multiline={true}
                            rows={5}
                            // disabled={
                            //   this.state.currentMode === "New" ? false : true
                            // }
                            onChange={(
                              event: React.FormEvent<
                                HTMLInputElement | HTMLTextAreaElement
                              >,
                              newValue: string
                            ) => this.onTextChange("baselineComment", newValue)}
                          />
                        </TooltipHost>
                      </div>
                    </div>
                  </div>
                  <div className={`${styles.msGridrow} ${styles.margin5}`}>
                    <div className={`${styles.msGridcol}  ${styles.mssm8}  `}>
                      <label className="ms-Label"> Linked Items</label>
                    </div>
                    <div className={`${styles.msGridcol}  ${styles.mssm2}  `}>
                      <TooltipHost content="Add" delay={TooltipDelay.long}>
                        <Icon
                          iconName="Add"
                          className={iconClass}
                          onClick={() => {
                            this.setState({
                              openSearchDialog: true,
                              isCheckLikedItem: false,
                            });
                          }}
                        />
                      </TooltipHost>
                    </div>
                    {this.state.updatedRows.filter((l) => l.isSelected === true)
                      .length > 0 && (
                        <div className={`${styles.msGridcol}  ${styles.mssm2}  `}>
                          <TooltipHost content="Remove" delay={TooltipDelay.long}>
                            <Icon
                              iconName="Remove"
                              className={iconClass}
                              onClick={() => {
                                this.setState({
                                  openUnLinkDialog: true,
                                  isCheckLikedItem: false,
                                });
                              }}
                            />
                            {/*hidden={this.state.updatedRows?.filter(l => l.isSelected === true).length === 0} />*/}
                          </TooltipHost>
                        </div>
                      )}
                  </div>
                  {/* Linked Items Content*/}
                  <div className={`${styles.msGridrow}`}>
                    <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
                      <div
                        className={`ag-theme-alpine rowmsGrid ${styles.agGridContainer}`}
                      >
                        <AgGridReact
                          rowData={this.state.linkedSearchItems}
                          columnDefs={columns}
                          context={this}
                          rowSelection="multiple"
                          rowMultiSelectWithClick={true}
                          domLayout={"autoHeight"}
                          //onSelectionChanged={this.onSelectionChanged}
                          onGridReady={this.gridReadyRowResize.bind(this)}
                          onRowSelected={this.onSelectionChanged.bind(this)}
                        />
                      </div>
                    </div>
                  </div>

                  {/* Footer Button section */}
                  <div className={styles.formFooter}>
                    <DefaultButton text="Cancel" onClick={this.props.onClose} />
                    <PrimaryButton
                      text="Save"
                      onClick={
                        this.props.itemID === undefined
                          ? this.submitForm.bind(this)
                          : this.UpdateForm.bind(this)
                      }
                      disabled={this.state.isSaving}
                    >
                      {this.state.isSaving && (
                        <Spinner
                          size={SpinnerSize.small}
                          className={styles.spinnerIcon}
                        />
                      )}
                    </PrimaryButton>
                  </div>
                </div>
              )}
            </div>

            {this.state.openSearchDialog && (
              <RelatedItems
                webpartContext={this.props}
                isOpen={this.state.openSearchDialog}
                onDismiss={() => {
                  this.setState({ openSearchDialog: false });
                }}
                selectionCallback={this.updatedLinkedItems.bind(this)}
              />
            )}

            {this.state.openUnLinkDialog && (
              <Dialog
                hidden={!this.state.openUnLinkDialog}
                onDismiss={() => {
                  this.setState({ openUnLinkDialog: false });
                }}
                dialogContentProps={{
                  type: DialogType.normal,
                  title: "Unlink",
                  subText: "De-associate the references",
                }}
                modalProps={{
                  isBlocking: true,
                  containerClassName: "ms-dialogMainOverride",
                }}
              >
                Are you sure about removing the references for this baseline? If
                yes, please clcik on the Proceed button, otherwise click on
                Cancel.
                <p />
                <DialogFooter>
                  <PrimaryButton
                    onClick={this.unlinkReferences.bind(this)}
                    text="Proceed"
                  />
                  <DefaultButton
                    onClick={() => {
                      this.setState({ openUnLinkDialog: false });
                    }}
                    text="Cancel"
                  />
                </DialogFooter>
              </Dialog>
            )}

            <Dialog
              hidden={!this.state.currentPresaveOperation}
              dialogContentProps={{
                type: DialogType.normal,
                title: "Pre-save Operation",
                subText:
                  this.state.currentPresaveOperation?.status === "running"
                    ? `${this.state.currentPresaveOperation?.label} (${this.state.currentPresaveOperation?.index}/${this.state.currentPresaveOperation?.count})`
                    : this.state.currentPresaveOperation?.status === "success"
                      ? `${this.state.currentPresaveOperation?.label} ✅`
                      : `${this.state.currentPresaveOperation?.label} ❌`,
              }}
            >
              <DialogFooter>
                {((this.state.currentPresaveOperation?.status === "success" &&
                  this.state.currentPresaveOperation?.index ===
                  this.state.currentPresaveOperation?.count) ||
                  this.state.currentPresaveOperation?.status === "error") && (
                    <DefaultButton
                      text="Close"
                      onClick={() =>
                        this.setState({ currentPresaveOperation: undefined })
                      }
                    />
                  )}
              </DialogFooter>
            </Dialog>
          </div>
        </div>
      </div>
    );
  }
}
