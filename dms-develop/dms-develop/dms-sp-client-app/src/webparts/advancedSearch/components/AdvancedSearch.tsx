/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import {
  //ICustomSearchResults,
  ISPType,
} from "../../../interfaces/IGlobalInterfaces";
import { IAdvancedSearchProps } from "./IAdvancedSearchProps";
import styles from "./AdvancedSearch.module.scss";
import AdvanceSearchContainer from "./AdvanceSearchContainer";
import { Dialog, DialogFooter, DialogType, Icon, IDropdownOption, PrimaryButton, ProgressIndicator, Spinner, SpinnerSize, Stack } from "@fluentui/react";
import { Log } from "@microsoft/sp-core-library";
import BaselineService from "../../../service/BaselineService";
import AdvanceSearchResultTable from "./AdvanceSearchResultTable";
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import { DmsRole } from "../../../service/PermissionsService";
import ExcelJS from 'exceljs';


const LOG_SOURCE: string = "SearchRelatedItems";


export interface IAdvanceSearchState {
  contentTypeOptions: IDropdownOption[];
  searchResult?: any[];
  selectedCType: string;
  hubSiteUrl: string;
  searchRequestlistData: any[];
  mySearches: any[];
  publicSearches: any[];
  searchQuery: string;
  IsValidPrjAdminUserL: boolean;
  selectedContentType: any;
  searchProcessed: boolean;
  showExportDialog: boolean;
  dialogMessage: string
}


export default class AdvancedSearch extends React.Component<
  IAdvancedSearchProps,
  IAdvanceSearchState
> {

  constructor(props: IAdvancedSearchProps) {
    super(props);
    this.state = {
      contentTypeOptions: [],
      searchResult: [],
      selectedCType: "",
      hubSiteUrl: "",
      searchRequestlistData: [],
      mySearches: [],
      publicSearches: [],
      searchQuery: "",
      IsValidPrjAdminUserL: false,
      selectedContentType: null,
      searchProcessed: false,
      dialogMessage: "Please wait... preparing the content to export..",
      showExportDialog: false,

    };
  }



  public async componentDidMount() {

    const prjAdmin = await this.props.permissionsService.currentUserHasRole(DmsRole.ProjectAdmin);
    this.setState({
      IsValidPrjAdminUserL: prjAdmin
    })

    await this.fetchHubSiteUrl();
    await this.getSearchRequests();

    // await this.loadPanel();
    Log.info(LOG_SOURCE, "React Element: RelatedItems mounted");
    const cTypes: ISPType[] = await BaselineService.getSiteContentTypes();
    this.setState({
      contentTypeOptions: cTypes.map((c) => {
        return { key: c.id, text: c.title };
      }),
    });

  }



  private async fetchHubSiteUrl() {
    try {
      const hubsiteData = await BaselineService.getHubSiteID();
      if (hubsiteData && hubsiteData.HubSiteId) {
        const sp = spfi().using(SPFx(this.props.context));

        const hubSiteInfo = await sp.hubSites.getById(hubsiteData.HubSiteId)();
        if (hubSiteInfo && hubSiteInfo.SiteUrl) {
          this.setState({ hubSiteUrl: hubSiteInfo.SiteUrl });
        }
      }
    } catch (error) {
      console.error("Error retrieving Hub Site URL:", error);
    }
  }

  // Fetch data from the SearchRequests list
  private async getSearchRequests() {
    const { hubSiteUrl } = this.state;
    if (hubSiteUrl) {
      try {
        const spWeb = Web(hubSiteUrl).using(SPFx(this.props.context));

        let finalItems: any[] = [];
        for await (const items of spWeb.lists.getByTitle(this.props.SearchRequestListName).items.top(2000)) {
          finalItems = finalItems.concat(items);
        }

        const currentUser = await spWeb.currentUser();

        const mySearches = finalItems.filter(item => item.Visibility === 'Personal' && item.User === currentUser.Email);

        const publicSearches = finalItems.filter(item => item.Visibility === 'Public');

        this.setState({
          mySearches: mySearches,
          publicSearches: publicSearches,
          searchProcessed: true
        });

      } catch (error) {
        console.error("Error retrieving search requests:", error);
      }
    }
  }


  private searchResult = (searchResult: any[]) => {
    this.setState({ searchResult: searchResult });
  };

  private exportData = async () => {
    const searchResult = this.state.searchResult ? this.state.searchResult : [];
    this.setState({ showExportDialog: true });
    if (searchResult.length === 0) {
      this.setState({ dialogMessage: "No content to export!" });
      return;
    }
    const keys = Object.keys(searchResult[0]);
    const columns: { header: string; key: string; }[] = [];
    keys.forEach(k => {
      columns.push({ header: k, key: k });
    })
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(`Export`);
    //const columns = searchResult.map(res=> res.Id)
    sheet.columns = columns;
    searchResult.forEach(obj => {
      const eachRow: any = {};
      keys.forEach(k => {
        eachRow[k] = obj[k];
      })
      sheet.addRow(eachRow);
    });
    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = "SearchResults.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
    this.setState({ showExportDialog: false });
  };

  private updateSelectedContentType = (cType: string) => {
    this.setState({ selectedCType: cType, searchResult: [] });
  };

  private handleSavedSearchClick = async (search: any) => {
    try {
      const selectedSearchQuery = search.Query;
      const searchObject = JSON.parse(selectedSearchQuery);

      this.setState({ searchQuery: selectedSearchQuery, selectedContentType: searchObject.selectedContentType });

    } catch (error) {
      console.error("Error loading the saved search query: ", error);
    }
  };

  private async reloadSearchQuery() {
    //this.forceUpdate();
   // this.setState({ searchQuery: "", selectedContentType: null });
    await this.getSearchRequests();
  }


  private handleDeleteSearch = async (search: any) => {

    try {

      const spWeb = Web(this.state.hubSiteUrl).using(SPFx(this.props.context));
      const searchRequestsList = spWeb.lists.getByTitle(this.props.SearchRequestListName);
      await searchRequestsList.items.getById(search.Id).delete();

      this.setState(prevState => ({
        mySearches: prevState.mySearches.filter(item => item.Id !== search.Id),
        publicSearches: prevState.publicSearches.filter(item => item.Id !== search.Id)
      }));

    } catch (error) {
      console.error("Error deleting search request:", error);
    }

  };


  public render(): React.ReactElement<IAdvancedSearchProps> {

    return (
      <div>
        <Stack className={styles.body}>
          <div className={styles.parentSearchblock}>
            <div
              className={`${styles.msGrid} ${styles.noPadding} ${styles.width100} ${styles.searchConfigblock}`}
            >
              <div className={styles.searchConfigHeader}>
                <h4 className={styles.headerTextSearchConfig}>Search Configuration</h4>
              </div>

              <div className={`${styles.msGridrow} ${styles.marginTop30}`}>
                <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
                  <AdvanceSearchContainer
                    contentTypeOptions={this.state.contentTypeOptions}
                    searchResults={this.searchResult}
                    selectedType={this.updateSelectedContentType}
                    webpartContext={this.props}
                    projectPropertiesService={this.props.projectPropertiesService}
                    prevsearchQuery={this.state.searchQuery}
                    reloadSearchQuery={this.reloadSearchQuery.bind(this)}
                    mySearch={this.state.mySearches}
                    publicSearch={this.state.publicSearches}
                  />
                </div>
              </div>

              <div className={styles.searchConfigHeader}>
                <h4 className={styles.headerTextSearchConfig}>Search Results</h4>
              </div>

              <AdvanceSearchResultTable
                rows={this.state.searchResult}
                contentType={this.state.selectedCType}
                exportCallback={this.exportData}
              />
            </div>
            <div className={styles.verticalbar}>
            </div>
            <div
              className={`${styles.msGrid} ${styles.noPadding} ${styles.width100} ${styles.saveSearchblock}`}>
              <div className={styles.searchConfigHeader}>
                <h4 className={styles.headerTextSearchConfig}>Saved Searches</h4>
              </div>
              <div className={styles.mySearchheader}>
                <Stack horizontal verticalAlign="center">
                  <Icon iconName="Search" /><h5>My Searches</h5>
                </Stack>
                {!this.state.searchProcessed && <Spinner label="loading..." size={SpinnerSize.small}></Spinner>}
                <ul>
                  {this.state.mySearches.map((search) => (

                    <Stack horizontal verticalAlign="center">
                      <li className={styles.clickable} onClick={() => this.handleSavedSearchClick(search)}>{search.Title}

                        <Icon className={styles.deleteIcon} iconName="Delete" onClick={() => this.handleDeleteSearch(search)} /></li>
                    </Stack>
                  ))}
                </ul>
              </div>

              <div className={styles.publicSearchheader}>
                <Stack horizontal verticalAlign="center">
                  <Icon iconName="Search" />
                  <h5>Public Searches</h5>
                </Stack>
                {!this.state.searchProcessed && <Spinner label="loading..." size={SpinnerSize.small}></Spinner>}
                <ul>
                  {this.state.publicSearches.map((itm) => (
                    <Stack horizontal verticalAlign="center">
                      <li className={styles.clickable} onClick={() => this.handleSavedSearchClick(itm)}>{itm.Title}</li>
                      {
                        this.state.IsValidPrjAdminUserL && <Icon className={styles.deleteIcon} iconName="Delete" onClick={() => this.handleDeleteSearch(itm)} />
                      }

                    </Stack>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        </Stack>
        {this.state.showExportDialog &&
          <Dialog
            hidden={!this.state.showExportDialog}
            onDismiss={() => { this.setState({ showExportDialog: false }) }}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Export Search Results",
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
                onClick={() => { this.setState({ showExportDialog: false }) }}
                text="OK"
              />
            </DialogFooter>
          </Dialog>        }
      </div>
    );
  }
}
