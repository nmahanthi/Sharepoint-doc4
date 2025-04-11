/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import styles from '../BaselineForms.module.scss';
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import * as React from "react";
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import { ICustomSearchResults, ISPColumn } from '../../../../interfaces/IGlobalInterfaces';
import CONSTANTS from '../../../../_Constants';

import { AgGridReact } from "ag-grid-react";
import { GridApi } from "ag-grid-community";

import 'ag-grid-community/dist/styles/ag-grid.css';
import 'ag-grid-community/dist/styles/ag-theme-alpine.css';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';

export interface IResultTableProps {
  rows?: ICustomSearchResults[];
  callBackSearchResults: (result?: ICustomSearchResults[]) => void;
  contentType: string;
  onDismiss: () => void;
}
export interface IResultTableState {
  operatorOptions: IDropdownOption[];
  selectedField?: ISPColumn;
  selectedOperator?: string;
  selectedCondition?: string;
  selectedstring?: string;
  errorMessage: string;
  errorControls: any[];
  gridAPI?: GridApi;
  updatedRows?: ICustomSearchResults[];
}
export default class ResultTable extends React.Component<IResultTableProps, IResultTableState> {
  constructor(props: IResultTableProps) {
    super(props);
    this.state = {
      operatorOptions: [],
      selectedField: undefined,
      selectedOperator: undefined,
      selectedCondition: undefined,
      selectedstring: undefined,
      errorMessage: "",
      errorControls: [],
      gridAPI: undefined,
      updatedRows: [],
    };
  }

  /* private deleteRow = (idx: number): void => {
      let rows = cloneDeep(this.props.rows);
      rows?.splice(idx, 1);
      this.props.callBackSearchResults(rows);
    };
 
  private selectRow = (isChecked: boolean, idx: number): void => {
    const rows = cloneDeep(this.props.rows);
    if (rows !== undefined)
      rows[idx].isSelected = isChecked;
    this.props.callBackSearchResults(rows);
  };
*/
  private onSelectionChanged = () => {
    const selectedRows = this.state.gridAPI?.getSelectedRows();
    const rows = cloneDeep(this.props.rows);
    rows?.map(r => {
      const exists = selectedRows?.some((srch) => srch.index === r.index);
      if (exists)
        r.isSelected = true;
      else
        r.isSelected = false;
    })
    console.log(rows)
    this.setState({ updatedRows: rows })
  };
  public gridReadyRowResize = (event: any) => {
    const api = event.api;
    console.log(api);
    console.log(event);
    api.forEachNode((node: any) => {
      console.log(node);
    })
    //api.showLoadingOverlay();
    setTimeout(() => {
      api.sizeColumnsToFit();
    }, 1500);
    this.setState({ gridAPI: api });
  };

  private addSelection = () => {
    this.props.callBackSearchResults(this.state.updatedRows);
  };
  public render(): React.ReactElement<{}> {

    const { rows } = this.props;
    console.log("rows:"+rows);
   
    const columns: any[] = [
      {
        headerName:  "TITLE",
        field:  "Title",
        checkboxSelection: true,
        // sortable: true,
        tooltipField:  "Title",
        headerCheckboxSelection: true,
        flex: 1,
        wrapText: true,
        resizable: true,
      },
      {
        headerName: this.props.contentType ===
CONSTANTS.ContentTypeNames.Technical ?  "DOCUMENT STATUS" :  "BASELINE STATUS",
        field: this.props.contentType === CONSTANTS.ContentTypeNames.Technical ?
 "DocumentStatus" :  "BaselineStatus",
        tooltipField: this.props.contentType ===
CONSTANTS.ContentTypeNames.Technical ?  "DocumentStatus" :  "BaselineStatus",
        flex: 1,
        wrapText: true,
        resizable: true,
      },
    ];
    if (this.props.contentType === CONSTANTS.ContentTypeNames.Technical) {
      columns.push({
 
        headerName:  "PROJECT REVISION",
        field:  "ProjectRevision",
        tooltipField:  "ProjectRevision",
        wrapText: true,
        flex: 1,
        resizable: true,
      });
      columns.push({
        headerName:  "PROJECT REFERENCE",
        field:  "ProjectReference",
        tooltipField:  "ProjectReference",
        wrapText: true,
        flex: 1,
        resizable: true,
      });
    }
    columns.push({
      headerName:  "SITE URL",
      field:  "SitePath",
      tooltipField:  "SitePath",
      flex: 1,
      wrapText: true,
      autoHeight: true,
      resizable: true,
    })

    return <React.StrictMode>
      <div><div className={`${styles.msGridrow}`}>
        <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
          <label className="ms-Label">Results :</label>
          <div className={`ag-theme-alpine rowmsGrid `}
            style={{
              width: "100%"
            }}>
            <AgGridReact
              rowData={rows}
              columnDefs={columns}
              context={this}
              rowSelection="multiple"
              rowMultiSelectWithClick={true}
              domLayout={'autoHeight'}
              onSelectionChanged={this.onSelectionChanged}
              onGridReady={this.gridReadyRowResize}
            />
          </div>
        </div>
      </div >

        <div className={`${styles.msGridrow}`}>
          <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
            <div className={styles.formFooter}>
              <PrimaryButton text='Add Selection' onClick={this.addSelection} disabled={this.state.updatedRows?.length === 0} />
              <DefaultButton text='Cancel' onClick={this.props.onDismiss} />
            </div>
          </div>
        </div>
      </div>
    </React.StrictMode >
  }

}