import * as React from "react";
import styles from "../BaselineForms.module.scss";
import { TooltipDelay, TooltipHost } from "@fluentui/react/lib/Tooltip";
import { TextField } from "@fluentui/react/lib/TextField";
import { DefaultButton} from "@fluentui/react/lib/Button";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import BaselineService from "../../../../service/BaselineService";
import {ICustomSearchResults, ILinkedItems } from "../../../../interfaces/IGlobalInterfaces";
import { AgGridReact } from "ag-grid-react";
import { GridApi } from "ag-grid-community";
import CONSTANTS from "../../../../_Constants";
import "ag-grid-community/dist/styles/ag-grid.css";
import "ag-grid-community/dist/styles/ag-theme-alpine.css";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { BaselineCommanBar } from "../FreezeBaseline/BaselineCommanBar";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { AadHttpClient } from "@microsoft/sp-http";
import { INotificationService } from "../../../../service/NotificationService";
import { IPermissionsService } from "../../../../service/PermissionsService";

export interface IViewFormContainerState {
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
  errorControls: any[];
  linkedSearchItems: ILinkedItems[];
  isPageLoading: boolean;
  gridAPI?: GridApi;
  technicalDocLinkedItems: ICustomSearchResults[];
  baselineLinkedItems: ILinkedItems[];
}

export interface IViewFormContainerProps {
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
//

export default class ViewFormContainer extends React.Component<

  IViewFormContainerProps,
  IViewFormContainerState
> {
  
  constructor(props: IViewFormContainerProps) {

    super(props);
    this.state = {
      openSearchDialog: false,
      openUnLinkDialog: false,
      title: "",
      baselineStatus: "",
      baselineVersion: "",
      cRID: "",
      code: "",
      label: "",
      baselineComment: "",
      errorMessage: "",
      errorControls: [],
      gridAPI: undefined,
      linkedSearchItems: [],
      isPageLoading: true,
      technicalDocLinkedItems: [],
      baselineLinkedItems: []
    };
  
  }
  public async componentDidMount() {
    const { item, itemID } = this.props;
    try {
      console.log('Fetching baseline and technical list data...');
      const baslineList = await BaselineService.getReferenceListItem(CONSTANTS.ListNames.BaselineRefernceList, itemID.toString(), this.props.context.pageContext.site.absoluteUrl);
      const technicalList = await BaselineService.getReferenceListItem(CONSTANTS.ListNames.TechnicalReferenceList, itemID.toString(), this.props.context.pageContext.site.absoluteUrl);

      const linkedSearchItems: ILinkedItems[] = [];
      let index = 0;

      const baselineItems: ICustomSearchResults[] = [];
      baslineList.map(b => {
        const baseItem = {
          DocumentId: undefined,
          SiteUrl: b[CONSTANTS.BaselineRefernceListFieldNames.SiteURL],
          ChildBaselineId: b[CONSTANTS.BaselineRefernceListFieldNames.BaselineChildID],
          ContentType: CONSTANTS.ContentTypeNames.Baseline,
          isSelected: false,
          Title: b[CONSTANTS.BaselineRefernceListFieldNames.Title],
          ProjectReference:b[CONSTANTS.BaselineListFieldNames.Code],
          ProjectRevision:b[CONSTANTS.BaselineListFieldNames.BaselineVersion],
          index: index++,
          Status: b[CONSTANTS.BaselineRefernceListFieldNames.Status]
        };
        linkedSearchItems.push(baseItem);
        baselineItems.push(baseItem);
      });

      const techdocs: ICustomSearchResults[] = [];
      await Promise.all(technicalList.map(async (t) => {
        const result = await BaselineService.getSearchResults(`DlcDocId=${t[CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]}`);
        if (result.length > 0) {
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
            ProjectRevision: result[0].ProjectRevision
          });
          techdocs.push(result[0]);
        }
      }));

      console.log('Baseline and technical documents fetched successfully.');
      this.setState({
        baselineLinkedItems: baselineItems,
        technicalDocLinkedItems: techdocs,
        title: item[CONSTANTS.BaselineListFieldNames.Title as keyof ILinkedItems] as string,
        baselineStatus: item[CONSTANTS.BaselineListFieldNames.BaselineStatus as keyof ILinkedItems] as string,
        baselineVersion: item[CONSTANTS.BaselineListFieldNames.BaselineVersion as keyof ILinkedItems] as string,
        cRID: item[CONSTANTS.BaselineListFieldNames.CRID as keyof ILinkedItems] as string,
        code: item[CONSTANTS.BaselineListFieldNames.Code as keyof ILinkedItems] as string,
        label: item[CONSTANTS.BaselineListFieldNames.Title as keyof ILinkedItems] as string,
        baselineComment: item[CONSTANTS.BaselineListFieldNames.BaselineComments as keyof ILinkedItems] as string,
        linkedSearchItems: linkedSearchItems,
      }, () => {
        this.state.gridAPI?.sizeColumnsToFit();
      });
      this.setState({ isPageLoading: false });
    } catch (error) {
      console.error('Error fetching baseline or technical data:', error);
      this.setState({ errorMessage: 'Failed to load data. Please try again later.' });
    }
  }

  public gridReadyRowResize = (event: any) => {
    const api = event.api;
    //api.showLoadingOverlay();
    setTimeout(() => {
      api.sizeColumnsToFit();
    }, 1500);
    this.setState({ gridAPI: api });
  };

  public render(): React.ReactElement<IViewFormContainerProps> {
    const columns: any[] = [

      {
        headerName: "TITLE",
        field: "Title",
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
        minWidth: 300
      },
    ];

    return (
      <div className={styles.formWrapper}>
        <BaselineCommanBar
          {...this.props}
          linkedSearchItems={this.state.linkedSearchItems}
          checkPendingValue={true}
          isCheckLikedItem={true}
          
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
                            disabled={true}
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
                            disabled={true}
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
                          content="Version"
                          delay={TooltipDelay.long}
                        >
                          <TextField
                            required
                            label="Baseline Version"
                            placeholder="Version"
                            value={this.state.baselineVersion}
                            disabled={true}
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
                            required
                            label="CR ID"
                            placeholder="CR ID"
                            value={this.state.cRID}
                            disabled={true}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "cRID"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>

                      <div className={styles.fieldWrapper}>
                        <TooltipHost content="Status" delay={TooltipDelay.long}>
                          <TextField
                            required
                            value={this.state.baselineStatus}
                            label="Baseline Status"
                            placeholder="Baseline Status"
                            disabled={true}
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
                      <div className={styles.fieldWrapper}> &nbsp;</div>
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
                            required
                            value={this.state.baselineComment}
                            label="Comments"
                            placeholder="comments"
                            multiline={true}
                            rows={5}
                            disabled={true}
                            className={
                              this.state.errorControls.some(
                                (e) => e.stateName === "baselineComment"
                              )
                                ? styles.required
                                : ""
                            }
                          />
                        </TooltipHost>
                      </div>
                    </div>
                  </div>

                  {/* Linked Items header*/}
                  <div className={`${styles.msGridrow}`}>
                    <div className={`${styles.msGridcol}  ${styles.mssm8}  `}>
                      <label className="ms-Label"> Linked Items</label>
                    </div>
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
                          onGridReady={this.gridReadyRowResize}
                        />
                      </div>
                    </div>
                  </div>

                  {/* Footer Button section */}
                  <div className={styles.formFooter}>
                    <DefaultButton text="Cancel" onClick={this.props.onClose} />
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
