/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import CONSTANTS from "./../../../_Constants";
import { AgGridReact } from "ag-grid-react";
import "ag-grid-community/dist/styles/ag-grid.css";
import "ag-grid-community/dist/styles/ag-theme-alpine.css";
import { ICustomSearchResults } from "../../../interfaces/IGlobalInterfaces";
import styles from "./AdvancedSearch.module.scss";
import { Open12Regular } from "@fluentui/react-icons";
import { ICellRendererParams } from "ag-grid-community";
import { PrimaryButton } from '@fluentui/react/lib/Button';
export interface IAdvanceSearchResultTableProps {
  rows?: ICustomSearchResults[];
  contentType: string;
  exportCallback: () => void;
}

export default class AdvanceSearchResultTable extends React.Component<IAdvanceSearchResultTableProps> {
  constructor(props: IAdvanceSearchResultTableProps) {
    super(props);
  }

  private handleProjectReferenceClick(rowData: { Path: string }): void {
    const url = `${rowData.Path}`;
    window.open(url, "_blank");
  }
  private handleSiteUrlClick(rowData: { Path: string }): void {
    const url = `${rowData.Path}`;
    window.open(url, "_blank");
  }

  public render(): React.ReactElement<{}> {
    const { rows } = this.props;
    const columns: any[] = [
      {
        headerName: "TITLE",
        field: "Title",
        tooltipField: "Title",
        wrapText: true,
        resizable: true,
        width:150
      },
      {
        headerName:
          this.props.contentType === CONSTANTS.ContentTypeNames.Technical
            ? "DOCUMENT STATUS"
            : "BASELINE STATUS",
        field:
          this.props.contentType === CONSTANTS.ContentTypeNames.Technical
            ? "Document Status"
            : "BaselineStatus",
        tooltipField:
          this.props.contentType === CONSTANTS.ContentTypeNames.Technical
            ? "Document Status"
            : "BaselineStatus",
        wrapText: true,
        resizable: true,
        width:150
      },
    ];

    if (this.props.contentType === CONSTANTS.ContentTypeNames.Technical) {
      columns.push({
        headerName: "PROJECT REFERENCE",
        field: "Reference",
        tooltipField: "Reference",
        wrapText: true,
        flex: 1,
        resizable: true,
        cellRenderer: (params: ICellRendererParams) => {
          return (
            <div style={{ display: "flex", alignItems: "center" }}>
              <Open12Regular
                className={`${styles.openNewRegular}`}
                onClick={() => this.handleProjectReferenceClick(params.data)}
              />
              {params.value}
            </div>
          );
        },
      });
      columns.push({
        headerName: "PROJECT REVISION",
        field: "Revision",
        tooltipField: "Revision",
        wrapText: true,
        flex: 1,
        resizable: true,
      });
      columns.push({
        headerName: "Project Name",
        field: "Prj Name",
        tooltipField: "Prj Name",
        flex: 1,
        wrapText: true,
        autoHeight: true,
        resizable: true,
      });
      columns.push({
        headerName: "Project Code",
        field: "Prj Code",
        tooltipField: "Prj Code",
        flex: 1,
        wrapText: true,
        autoHeight: true,
        resizable: true,
      });
    }
    if (this.props.contentType === CONSTANTS.ContentTypeNames.Baseline) {
      columns.push({
        headerName: "SITE URL",
        field: "SitePath",
        tooltipField: "SitePath",
        flex: 1,
       // wrapText: true,
       // autoHeight: true,
        resizable: true,
        cellRenderer: (params: ICellRendererParams) => {
          return (
            <div style={{ cursor: 'pointer', color: 'blue', textDecoration: 'underline' }} onClick={() => this.handleSiteUrlClick(params.data)}>
              {params.value}
            </div>
          );
        },
      });
    }


    return (
      <React.StrictMode>
        <div>
          <div className={`${styles.formFooter}`}>
            <div className={styles.exportButtonresultTop}>
              <PrimaryButton text='Export' disabled={(this.props.rows?.length === 0)} onClick={this.props.exportCallback}>

              </PrimaryButton>
            </div>
            <div className={`${styles.msGridcol} ${styles.mssm12}`}>
              {/* <label className="ms-Label">Results :</label> */}
              <div
                className={`ag-theme-alpine rowmsGrid`}
                style={{ width: "100%",textAlign:"left" }}
              >
                <AgGridReact
                  rowData={rows}
                  columnDefs={columns}
                  context={this}
                  domLayout={"autoHeight"}
                  pagination={true}
                  onCellClicked={(params) => {
                    if (params.colDef.field === "ProjectReference") {
                      this.handleProjectReferenceClick(params.data);
                    }
                  }}
                />
              </div>
              <div className={styles.exportresultBottam}>
              
              <PrimaryButton text='Export' disabled={(this.props.rows?.length === 0)} onClick={this.props.exportCallback}>

              </PrimaryButton>
            </div>
            </div>
           
          </div>
        </div>
      </React.StrictMode>
    );
  }
}
