/* eslint-disable @typescript-eslint/explicit-function-return-type */

import { Stack } from "@fluentui/react/lib/Stack";
import { TooltipDelay, TooltipHost } from "@fluentui/react/lib/Tooltip";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import * as React from "react";
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter, TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import * as _ from 'lodash';
import { ISearchParam, ISPColumn } from "../../../interfaces/IGlobalInterfaces";
import BaselineService from "../../../service/BaselineService";
import styles from "./AdvancedSearch.module.scss";
import SearchParameterRow from "../../../extensions/baselineForms/components/SearchContainer/SearchParameterRow";
import CONSTANTS from '../../../_Constants';
import "@pnp/sp/hubsites";
import { Web } from "@pnp/sp/webs";
// import { getSP } from "../../../pnpjs-config";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/site-users/web";
import { IProjectPropertiesService } from "../../../service/IProjectPropertiesService";




export interface ISearchContainerProps {
  webpartContext: any;
  contentTypeOptions: IDropdownOption[];
  searchResults: (result: any[]) => void;
  selectedType: (result: string) => void;
  projectPropertiesService: IProjectPropertiesService;
  prevsearchQuery: string;
  reloadSearchQuery: () => void;
  mySearch: any[];
  publicSearch: any[];
}
export interface ISearchContainerState {
  searchRows: ISearchParam[];
  contentTypeFields: ISPColumn[];
  searchQuery: string;
  hubsiteID: string;
  selectedContentType?: IDropdownOption;
  isSearching: boolean;
  validationError: string[];
  rawContentTypeFields: ISPColumn[];
  isDialogOpen: boolean;
  selectedOption: string;
  showSuccessQueryDialog: boolean;
  searchName: string;
  projectName?: string;
  disableSave: boolean;
  querySaveMessage: string;
}
export default class AdvanceSearchContainer extends React.Component<ISearchContainerProps, ISearchContainerState> {

  public sp: any | undefined;
  public hubsiteUrl: string;
  constructor(props: ISearchContainerProps) {
    super(props);
    this.state = {

      hubsiteID: "",
      rawContentTypeFields: [],
      validationError: [],
      isSearching: false,
      selectedContentType: undefined,
      searchRows: [
        {
          field: undefined,
          condition: "",
          operator: "",
          searchFieldName: undefined,
          searchText: "",
        },
      ],
      searchQuery: "",
      contentTypeFields: [],
      isDialogOpen: false,
      selectedOption: '',
      showSuccessQueryDialog: false,
      searchName: '',
      projectName: '',
      disableSave: true,
      querySaveMessage: ""
    };
  }

  async componentDidMount() {

    const { projectPropertiesService } = this.props;

    const [propertyBags] = await Promise.all([projectPropertiesService.getProperties()]);

    if (propertyBags && propertyBags.dms_proj_name) {
      this.setState({ projectName: propertyBags.dms_proj_name });
    } else {
      console.warn("Project name not found in the property bag.");
    }

    try {

      const sp = spfi().using(SPFx(this.props.webpartContext.context));
      if (!sp) {
        console.error("PnPJS is not initialized properly.");
        return;
      }

      const hubsiteData = await BaselineService.getHubSiteID();

      if (!hubsiteData || !hubsiteData.HubSiteId || hubsiteData.HubSiteId === '00000000-0000-0000-0000-000000000000') {
        console.warn("This site is not associated with any hub.");
        return;
      }

      this.setState({ hubsiteID: hubsiteData.HubSiteId });

      if (!sp.hubSites) {
        console.error("hubSites API is not available in PnPJS.");
        return;
      }
      const hubSiteInfo = await sp.hubSites.getById(hubsiteData.HubSiteId)();
      if (!hubSiteInfo || !hubSiteInfo.SiteUrl) {
        console.error("Failed to retrieve Hub Site URL.");
        return;
      }
      this.hubsiteUrl = hubSiteInfo.SiteUrl;
      const SPWeb = Web(hubSiteInfo.SiteUrl).using(SPFx(this.props.webpartContext));
      console.log(SPWeb);
    } catch (error) {
      console.error("Error retrieving Hub Site data:", error);
    }
  }

  async componentDidUpdate(prevProps: ISearchContainerProps) {
    if (this.props.prevsearchQuery !== prevProps.prevsearchQuery) {
      if(this.props.prevsearchQuery.length === 0)
      {
        this.forceUpdate();
        this.setState({selectedContentType:undefined});
        return;
      }
      const SearchQuerydata = JSON.parse(this.props.prevsearchQuery)
      // console.log("this is search rowsss",SearchQuerydata.searchRows);
      // console.log("search param in update",SearchQuerydata.searchParam);


      //   this.setState({ 
      //     searchQuery: SearchQuerydata.searchQuery,
      //     selectedContentType:SearchQuerydata.selectedContentType,
      //     searchRows:SearchQuerydata.searchRows});
      // }

      const data = this.props.contentTypeOptions.filter(c => c.text ===
        SearchQuerydata.selectedContentType.text);

      await this.typeChanged(data[0]);
      this.setState({
        searchQuery: SearchQuerydata.searchQuery,
        selectedContentType: data.length > 0 ? data[0] : undefined,
        searchRows: SearchQuerydata.searchRows,
        disableSave: true
      });
    }
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

  private typeChanged = async (item: IDropdownOption) => {

    let flds: ISPColumn[] = [];
    if (item && item.key)
      flds = await BaselineService.getContentTypeFields(item.key.toString());
    //Included this to add it in Export data, since these will be filtered out based on the configurable column names
    flds.push({ group: "_Hidden", hidden: false, internalName: "Created", title: "Created Date", typeAsString: "DateTime", id: 'Created-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "CreatedBy", title: "Created By", typeAsString: "Text", id: 'CreatedBy-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "ModifiedBy", title: "Modified By", typeAsString: "Text", id: 'ModifiedBy-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "LastModifiedTime", title: "Modified Date", typeAsString: "DateTime", id: 'LastModifiedTime-001' });
    this.setState({ rawContentTypeFields: flds });
    if (item && item.text === CONSTANTS.ContentTypeNames.Baseline) {
      const newFields = this.props.webpartContext.baseLineFields.split(",");
      flds = flds.filter(f => newFields.includes(f.internalName));
    }
    if (item && item.text === CONSTANTS.ContentTypeNames.Technical) {
      const newFields = this.props.webpartContext.technicalFields.split(",");
      flds = flds.filter(f => newFields.includes(f.internalName));
    }

    //Adding again to include it in fields drop down
    flds.push({ group: "_Hidden", hidden: false, internalName: "Created", title: "Created Date", typeAsString: "DateTime", id: 'Created-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "CreatedBy", title: "Created By", typeAsString: "Text", id: 'CreatedBy-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "ModifiedBy", title: "Modified By", typeAsString: "Text", id: 'ModifiedBy-001' });
    flds.push({ group: "_Hidden", hidden: false, internalName: "LastModifiedTime", title: "Modified Date", typeAsString: "DateTime", id: 'LastModifiedTime-001' });

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
    this.setState({ searchQuery: searchQuery, disableSave: false });
  }

  private search = async () => {
    const emptySearch = [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }];
    // search if there is no change in parameters, if not do the validations
    if (!_.isEqual(this.state.searchRows, emptySearch)) {
      const errorList = this.validateSearchParam();
      if (errorList.length > 0)
        return
    }
    this.setState({ isSearching: true });
    // const result = await BaselineService.getSearchResults(this.state.searchQuery);
    const searchFields: any[] = [];
    this.state.rawContentTypeFields.map(fld => {
      searchFields.push({ "SearchPropertyName": this.getSearchFieldName(fld), "DisplayName": fld.title });
    });
    searchFields.push({ "SearchPropertyName": "ListItemID", "DisplayName": "ItemID" });
    searchFields.push({ "SearchPropertyName": "SitePath", "DisplayName": "SitePath" });
    searchFields.push({ "SearchPropertyName": "ContentType", "DisplayName": "ContentType" });
    searchFields.push({ "SearchPropertyName": "Path", "DisplayName": "Path" });
    searchFields.push({ "SearchPropertyName": "DlcDocId", "DisplayName": "DocumentId" });
    searchFields.push({ "SearchPropertyName": "DocId", "DisplayName": "DocId" });
    const result = await BaselineService.getSearchResultsbyProperties(this.state.searchQuery, searchFields.map(obj => obj["SearchPropertyName"]));
    const processedResult: any[] = [];
    result.forEach(row => {
      const singleResult: any = {};
      searchFields.forEach(flds => {
        singleResult[flds.DisplayName] = row[flds.SearchPropertyName]
      });
      console.log(singleResult);
      processedResult.push(singleResult);
    });
    console.log(processedResult);
    this.props.searchResults(processedResult);
    this.setState({ isSearching: false });
  };



  private visibilityPopup = () => {
    const emptySearch = [{ field: undefined, condition: "", operator: "", searchFieldName: undefined, searchText: "" }];
    // search if there is no change in parameters, if not do the validations
    if (!_.isEqual(this.state.searchRows, emptySearch)) {
      const errorList = this.validateSearchParam();
      if (errorList.length > 0)
        return
    }
    this.setState({ isDialogOpen: true })
  }


  private saveSelectedOption = (selectedKey: string | number) => {
    console.log('Current selected option:', selectedKey);
    this.setState({ selectedOption: selectedKey.toString() }, () => {
      // This callback is called after the state is updated
      console.log('Updated state.selectedOption:', this.state.selectedOption);
    });
  };


  // Function to save the search query to the "SearchRequest" list
  private saveSearchQuery = async () => {
    try {

      this.setState({ querySaveMessage: "" })
      const sp = spfi(this.hubsiteUrl).using(SPFx(this.props.webpartContext.context));

      const currentUser = await sp.web.currentUser();
      if (!currentUser || !currentUser.Email) {
        console.error("Failed to retrieve current user's email.");
        return;
      }
      // if (this.state.selectedOption === "Public") {
      //   const result = this.props.publicSearch.filter(item => item.Title.toLowerCase() === this.state.searchName.toLowerCase());
      //   if (result.length > 0) {
      //     this.setState({ querySaveMessage: "Name already exists, please pick something new!" });
      //     return;
      //   }
      // }
      // if (this.state.selectedOption === "Personal") {
      //   const result = this.props.mySearch.filter(item => item.Title.toLowerCase() === this.state.searchName.toLowerCase());
      //   if (result.length > 0) {
      //     this.setState({ querySaveMessage: "Name already exists, please pick something new!" });
      //     return;
      //   }
      // }

      const prevQueries =[...this.props.publicSearch, ...this.props.mySearch];
        const result = prevQueries.filter(item => item.Title.toLowerCase() === this.state.searchName.toLowerCase());
        if (result.length > 0) {
          this.setState({ querySaveMessage: "Name already exists, please pick something new!" });
          return;
        }

      this.setState({ isDialogOpen: false })
      const list = sp.web.lists.getByTitle(this.props.webpartContext.SearchRequestListName);
      const fullSearchQuery = {
        selectedContentType: this.state.selectedContentType,
        searchQuery: this.state.searchQuery,
        searchRows: this.state.searchRows
      }
      const newSearchItem = {
        Title: this.state.searchName,
        Project: this.state.projectName,
        Query: JSON.stringify(fullSearchQuery),
        Visibility: this.state.selectedOption,
        User: currentUser.Email
      };

      await list.items.add(newSearchItem);
      this.props.reloadSearchQuery();
      this.setState({
       // searchQuery: '',
        selectedOption: '',
        showSuccessQueryDialog: true,
        disableSave: true
      })
      this.props.reloadSearchQuery();
    } catch (error) {
      console.error("Error saving the search query:", error);
    }
  };


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





  private getSearchFieldName(field: ISPColumn) {

    const fieldName = field.internalName;

    if (field.title === "Title" || field.title === "Name")
      return "Title";
    if (field.title.indexOf("Created") === 0 || field.title.indexOf("Modified") === 0)
      return fieldName;

    switch (field.typeAsString) {

      case CONSTANTS.FieldTypes.Text:

        return fieldName + "OWSTEXT";

      case CONSTANTS.FieldTypes.Note:
        return fieldName + "OWSMTXT";


      case CONSTANTS.FieldTypes.Choice:
      case CONSTANTS.FieldTypes.MultiChoice:
        return fieldName + "OWSCHCS";

      case CONSTANTS.FieldTypes.TaxonomyFieldType:
      case CONSTANTS.FieldTypes.TaxonomyFieldTypeMulti:

        return "owstaxId" + fieldName.replace(" ", "x0020");

      case CONSTANTS.FieldTypes.URL:
        return fieldName + "OWSURLH";

      case CONSTANTS.FieldTypes.User:
      case CONSTANTS.FieldTypes.UserMulti:
        return fieldName + "OWSUSER";

      case CONSTANTS.FieldTypes.Boolean:
        return fieldName + "OWSBOOL";

      case CONSTANTS.FieldTypes.DateTime:
        return fieldName + "OWSTDATE";
    }
    return fieldName;
  }
  public render(): React.ReactElement<ISearchContainerProps> {

    return <Stack className={styles.body}>
      <div className={`${styles.msGrid} ${styles.noPadding} ${styles.width100}`} >
        <div className={`${styles.msGridrow}`}>
          <div className={`${styles.msGridcol}  ${styles.mssm4}  `}>
            <TooltipHost content="Pick a content type" delay={TooltipDelay.long} >
              <Dropdown placeholder="Pick a content type"
                label="Type"
                options={this.props.contentTypeOptions}
                selectedKey={this.state.selectedContentType?.key}
                onChanged={this.typeChanged}
              />
            </TooltipHost>
          </div>
        </div>
        {this.state.searchRows.map((row, index) => {
          return <SearchParameterRow fieldOptions={this.state.contentTypeFields} callbackSearchParamChanged={this.updateSearchParams} searchParam={row} addRow={this.addRow} deleteRow={this.deleteRow} index={index} showAdd={index === 0} />
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
              <PrimaryButton text='Save' onClick={this.visibilityPopup} disabled={this.state.disableSave}>
              </PrimaryButton>
              <PrimaryButton text='Search' onClick={this.search}
                disabled={(this.state.isSearching || this.state.selectedContentType === undefined)}
              >
                {this.state.isSearching && <Spinner
                  size={SpinnerSize.small}
                  className={styles.spinnerIcon}
                />}
              </PrimaryButton>
            </div>
          </div>
        </div>
      </div>

      <div>
        <Dialog
          hidden={!this.state.isDialogOpen}
          onDismiss={() => this.setState({ isDialogOpen: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Please Select Visibility.'
          }}
        >
          {
            this.state.querySaveMessage.length > 0 && <div style={{ padding: "10px 0px", color: "red" }}>{this.state.querySaveMessage}</div>
          }
          <TextField
            label="Please provide name  for search"
            value={this.state.searchName} // Current value from state
            onChange={(e, newValue) => this.setState({ searchName: newValue || '' })} // Update state on change
          />
          <Dropdown
            placeholder="Select an option"
            label="Options"
            selectedKey={this.state.selectedOption}
            options={[
              { key: 'Public', text: 'Public' },
              { key: 'Personal', text: 'Personal' }
            ]}
            onChange={(event, option) => {
              if (option) {
                this.saveSelectedOption(option.key);
              }
            }}
          />
          <DialogFooter><PrimaryButton text="OK" onClick={this.saveSearchQuery} disabled={this.state.searchName.length === 0 || this.state.selectedOption.length === 0} /><PrimaryButton text="Cancel" onClick={() => this.setState({ isDialogOpen: false })} /></DialogFooter></Dialog>
      </div>
      <div>
        <Dialog
          hidden={!this.state.showSuccessQueryDialog}
          onDismiss={() => this.setState({ showSuccessQueryDialog: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Success',
            subText: 'Search query has been saved successfully!'
          }}>
          <DialogFooter><PrimaryButton text="OK" onClick={() => this.setState({ showSuccessQueryDialog: false })} /></DialogFooter>
        </Dialog>
      </div>

    </Stack>
  }

}