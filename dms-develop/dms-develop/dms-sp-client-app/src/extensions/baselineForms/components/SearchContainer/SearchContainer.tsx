/* eslint-disable @typescript-eslint/explicit-function-return-type */
import styles from '../BaselineForms.module.scss';

//import * as strings from "BaselineWebPartStrings";
//import Modal from "@fluentui/react/lib/Modal";
//import { IconButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { TooltipDelay, TooltipHost } from "@fluentui/react/lib/Tooltip";
//import { TextField } from "@fluentui/react/lib/TextField";
import BaselineService from "../../../../service/BaselineService";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import * as React from "react";
import SearchParameterRow from "./SearchParameterRow";
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import CONSTANTS from '../../../../_Constants';
import { Icon } from '@fluentui/react/lib/Icon';
import { ICustomSearchResults, ISearchParam, ISPColumn } from '../../../../interfaces/IGlobalInterfaces';
import * as _ from 'lodash';



export interface ISearchContainerProps {
  webpartContext: any;
  contentTypeOptions: IDropdownOption[];
  searchResults: (result: ICustomSearchResults[]) => void;
  selectedType: (result: string) => void;
}
export interface ISearchContainerState {
  searchRows: ISearchParam[];
  contentTypeFields: ISPColumn[];
  searchQuery: string;
  hubsiteID: string;
  selectedContentType?: IDropdownOption;
  isSearching: boolean;
  validationError: string[];
}
export default class SearchContainer extends React.Component<ISearchContainerProps, ISearchContainerState> {
  constructor(props: ISearchContainerProps) {
    super(props);
    this.state = { hubsiteID: "", validationError: [], isSearching: false, selectedContentType: undefined, searchRows: [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }], searchQuery: "", contentTypeFields: [] };

  }
  async componentDidMount() {
    const hubsiteId = await BaselineService.getHubSiteID();
    this.setState({ hubsiteID: hubsiteId?.HubSiteId });
  }

  //ON ADDING A NEW ROW ON SEARCH PATTERN
  public addRow = () => {
    const srch: ISearchParam[] = cloneDeep(this.state.searchRows);
    srch.push({ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" });
    this.setState({ searchRows: srch });
  }

  //ON DELETING A ROW ON SEARCH PATTERN
  public deleteRow = async (idx: number) => {
    const srch: ISearchParam[] = cloneDeep(this.state.searchRows);

    srch.splice(idx, 1);
    this.setState({ searchRows: srch });
    this.processSearchquery(srch);

  }

  //ON CONTENTTYPE/TYPE CHANGE EVENT
  // private typeChanged = async (item: IDropdownOption) => {
  //   let flds: ISPColumn[] = [];
  //   if (item && item.key)
  //     flds = await BaselineService.getContentTypeFields(item.key.toString());

  //   let searchQuery = `ContentType:"${item.text}" `;
  //   if (this.state.hubsiteID && this.state.hubsiteID !== '00000000-0000-0000-0000-000000000000') {
  //     searchQuery += `DepartmentID:${this.state.hubsiteID} `;
  //   }

  //   this.setState({ selectedContentType: item, searchQuery: searchQuery, searchRows: [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }] });

  //   this.setState({ contentTypeFields: flds });
  //   this.props.selectedType(item.text);
  // };

  private typeChanged = async (item: IDropdownOption) => {
    let flds: ISPColumn[] = [];
    if (item && item.key)
      flds = await BaselineService.getContentTypeFields(item.key.toString());
   
    if (item && item.text === CONSTANTS.ContentTypeNames.Baseline) {
      const newFields = this.props.webpartContext.baseLineFields.split(",");
      flds = flds.filter(f => newFields.includes(f.internalName));
    }
    if (item && item.text === CONSTANTS.ContentTypeNames.Technical) {
      const newFields = this.props.webpartContext.technicalFields.split(",");
      flds = flds.filter(f => newFields.includes(f.internalName));
    }

    let searchQuery = `ContentType:"${item.text}" `;
    if (this.state.hubsiteID && this.state.hubsiteID !== '00000000-0000-0000-0000-000000000000') {
      searchQuery += `DepartmentID:${this.state.hubsiteID} `;
    }

    this.setState({ selectedContentType: item, searchQuery: searchQuery, searchRows: [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }] });
 

    this.setState({ contentTypeFields: flds });
    this.props.selectedType(item.text);
  };


  //CALL BACK TO UPDATE QUERY STRING
  private updateSearchParams = (newParam: ISearchParam, index: number) => {
    const currentParams: ISearchParam[] = cloneDeep(this.state.searchRows);
    const changedParam: ISearchParam = currentParams[index];

    changedParam.condition = newParam.condition;
    changedParam.field = newParam.field;
    changedParam.operator = newParam.operator;
    changedParam.searchFieldName = newParam.searchFieldName;
    changedParam.searchText = newParam.searchText;
    this.setState({ searchRows: currentParams, validationError: [] });


    this.processSearchquery(currentParams);

  };

  //FORMAT SEARCH QUERY BASED ON SELECTED DATA
  private processSearchquery(currentParams: ISearchParam[]) {
    // let searchQuery = `ContentType:"${this.state.selectedContentType?.text}" DepartmentID:${this.state.hubsiteID} `;
    let searchQuery = `ContentType:"${this.state.selectedContentType?.text}" `;
    if (this.state.hubsiteID && this.state.hubsiteID !== '00000000-0000-0000-0000-000000000000') {
      searchQuery += `DepartmentID:${this.state.hubsiteID} `;
    }
    currentParams.map(srch => {
      if (srch.operator !== '-') {
        searchQuery += srch.searchFieldName ? `(${srch.searchFieldName}` : "";
        if (srch.operator) {
          if (srch.operator === ':text*' || srch.operator === ':*text' || srch.operator === ':*text*') {
            searchQuery += srch.operator.replace("text", srch.searchText ? `${srch.searchText}` : "") + ")";
          }
          else {
            searchQuery += srch.operator ? srch.operator : "";
            searchQuery += srch.searchText ? (srch.searchText.indexOf(" ") >= 0 ? `"${srch.searchText}")` : `${srch.searchText})`) : "";
          }
        }
      }
      else if (srch.operator === '-') {
        searchQuery += `(-${srch.searchFieldName ? srch.searchFieldName : ""}:${srch.searchText ? srch.searchText : ""})`
      }

      if (currentParams.length > 1)
        searchQuery += srch.condition ? ` ${srch.condition} ` : "";
    })
    this.setState({ searchQuery: searchQuery });
  }


  // private processSearchquery(currentParams: ISearchParam[]) {
  //   let searchQuery = `ContentType:"${this.state.selectedContentType?.text}" DepartmentID:${this.state.hubsiteID} `;
  //   currentParams.map(srch => {
  //     searchQuery += srch.searchFieldName ? `(${srch.searchFieldName}` : "";
  //     if (srch.operator) {
  //       if (srch.operator === ':text*' || srch.operator === ':*text' || srch.operator === ':*text*') {
  //         searchQuery += srch.operator.replace("text", srch.searchText ? `${srch.searchText}` : "") + ")";
  //       }
  //       else {
  //         searchQuery += srch.operator ? srch.operator : "";
  //         searchQuery += srch.searchText ? (srch.searchText.indexOf(" ") >= 0 ? `"${srch.searchText}")` : `${srch.searchText})`) : "";
  //       }
  //     }

  //     if (currentParams.length > 1)
  //       searchQuery += srch.condition ? ` ${srch.condition} ` : "";
  //   })
  //   this.setState({ searchQuery: searchQuery });
  // }

  //SEARCH BUTTON ON CLICK EVEVNT TO SEARCH FROM SHAREPOINT DATA
  private search = async () => {
    const emptySearch = [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }];
    // search if there is no change in parameters, if not do the validations
    if (!_.isEqual(this.state.searchRows, emptySearch)) {
      const errorList = this.validateSearchParam();
      if (errorList.length > 0)
        return
    }
    this.setState({ isSearching: true });
    const result = await BaselineService.getSearchResults(this.state.searchQuery);
    this.props.searchResults(result);
    this.setState({ isSearching: false });
  };

  //VALIDATE ENTERED DATA
  private validateSearchParam = () => {
    const currentParams: ISearchParam[] = cloneDeep(this.state.searchRows);
    const errorText: string[] = [];
    currentParams.map((param, index) => {
      if (index === (currentParams.length - 1)) {
        if (param.field === undefined || param.operator === undefined || param.searchText === undefined
          || param.operator === "" || param.searchText === "")

          errorText.push(`Row ${(index + 1)} should have all the entered.`);
      }
      else {
        if (param.condition === undefined || param.field === undefined || param.operator === undefined || param.searchText === undefined
          || param.condition === "" || param.operator === "" || param.searchText === "")

          errorText.push(`Row ${(index + 1)} should have all the entered.`);
      }
      if (param.field?.typeAsString === CONSTANTS.FieldTypes.Boolean) {
        const regex = /^(0|1)$/; // Matches any string containing only 0s and 1s
        const isValid = regex.test(param.searchText === undefined ? "" : param.searchText);
        if (!isValid)
          errorText.push(`Row ${(index + 1)}: Search text should be 0 or 1 for Boolean type of field.`);
      }
      else if (param.field?.typeAsString === CONSTANTS.FieldTypes.DateTime) {
        const regex = /\d{4}-\d{2}-\d{2}/; // Matches any string containing only 0s and 1s
        const isValid = regex.test(param.searchText === undefined ? "" : param.searchText);
        if (!isValid)
          errorText.push(`Row ${(index + 1)}: Search text should be in the format of YYYY-MM-DD for Date column.`);
      }
    });
    this.setState({ validationError: errorText });
    return errorText;
  };

  public render(): React.ReactElement<ISearchContainerProps> {

    return <Stack className={styles.body}>
      <div className={`${styles.msGrid} ${styles.noPadding} ${styles.width100}`} >
        <div className={`${styles.msGridrow}`}>
          <div className={`${styles.msGridcol}  ${styles.mssm4}  `}>
            <TooltipHost content="Title" delay={TooltipDelay.long} >
              <Dropdown placeholder="Pick a content type"
                label="Type"
                options={this.props.contentTypeOptions}
                onChanged={this.typeChanged}
              />
            </TooltipHost>
          </div>
        </div>
        {this.state.searchRows.map((row, index) => {
          // eslint-disable-next-line react/jsx-key
          return <SearchParameterRow fieldOptions={this.state.contentTypeFields} callbackSearchParamChanged={this.updateSearchParams} searchParam={row} addRow={this.addRow} deleteRow={this.deleteRow} index={index} showAdd={index === 0} />
          // eslint-disable-next-line no-lone-blocks
          {/*return <SearchParameterRow fieldOptions={this.state.contentTypeFields} callbackSearchParamChanged={this.updateSearchParams} searchParam={row} addRow={this.addRow} deleteRow={this.deleteRow} index={index} showAdd={index === (this.state.searchRows.length - 1)}/>*/ }
        })
        }
        {
          this.state.validationError.length > 0 && <div className={`${styles.msGridrow}`}>
            <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
              <div className={`${styles.notice} ${styles.error}`} id="userMessage">
                {
                  this.state.validationError.map(r => {
                    return <>
                      <Icon iconName="Error" />
                      <span>&nbsp;{r}</span><br />
                    </>
                  })
                }
              </div>
            </div>
          </div>
        }
        <div className={`${styles.msGridrow}`}>
          <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
            <div className={styles.searchQueryContainer}>
              {this.state.searchQuery}
            </div>
          </div>
        </div>
        <div className={`${styles.msGridrow}`}>
          <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
            <div className={styles.formFooter}>
              <PrimaryButton text='Search' onClick={this.search} disabled={(this.state.isSearching || this.state.selectedContentType === undefined)}>
                {this.state.isSearching && <Spinner
                  size={SpinnerSize.small}
                  className={styles.spinnerIcon}
                />}
              </PrimaryButton>
            </div>
          </div>
        </div>
      </div>
    </Stack>
  }

}