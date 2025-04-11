/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { Log } from "@microsoft/sp-core-library";
import styles from '../BaselineForms.module.scss';
//import * as strings from "BaselineWebPartStrings";
import Modal from "@fluentui/react/lib/Modal";
import { IconButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
//import { TextField } from "@fluentui/react/lib/TextField";
import BaselineService from "../../../../service/BaselineService";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import * as React from "react";
import SearchContainer from "../SearchContainer/SearchContainer";
import ResultTable from "./ResultTable";
import { ICustomSearchResults, ISPType } from "../../../../interfaces/IGlobalInterfaces";
//import BaselineService from "../../../../services/BaselineService";

const LOG_SOURCE: string = 'RelatedItems';


export interface IRelatedItemsProps {
  webpartContext: any;
  isOpen: boolean;
  onDismiss: () => void;
  selectionCallback: (searchResul: ICustomSearchResults[]) => void;
}
export interface IRelatedItemsState {
  contentTypeOptions: IDropdownOption[];
  searchResult?: ICustomSearchResults[];
  selectedCType: string;
}
export default class RelatedItems extends React.Component<IRelatedItemsProps, IRelatedItemsState> {

  constructor(props: IRelatedItemsProps) {
    super(props);
    this.state = {
      contentTypeOptions: [],
      searchResult: [],
      selectedCType: ""
    };
  }

  public async componentDidMount() {
    Log.info(LOG_SOURCE, 'React Element: RelatedItems mounted');
    const cTypes: ISPType[] = await BaselineService.getSiteContentTypes();
    this.setState({ contentTypeOptions: cTypes.map(c => { return { key: c.id, text: c.title } }) });
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: RelatedItems unmounted');
  }

  //CALL BACK SEARCH RESULT
  private searchResult = (searchResult: ICustomSearchResults[]) => {
    this.setState({ searchResult: searchResult });
  }

  //CALL BACK TO UPDATE SELECTED CONTENT TYPE
  private updateSelectedContentType = (cType: string) => {
    this.setState({ selectedCType: cType, searchResult: [] });
  }

  private finalSearchResult = (searchResult: ICustomSearchResults[]) => {
    this.setState({ searchResult: searchResult });
    this.props.selectionCallback(searchResult !== undefined ? searchResult : []);
    this.props.onDismiss();
  }


  public render(): React.ReactElement<{}> {
    const { isOpen, onDismiss } = this.props;

    return <Modal
      isBlocking={true}
      isOpen={isOpen}
      onDismiss={onDismiss}
      containerClassName={styles.DialogContainer}
      layerProps={{eventBubblingEnabled:true}}
    >
      <div className={styles.header}>
        <h2 className={styles.heading}>
          Search
        </h2>
        <IconButton
          className={styles.CloseButton}
          iconProps={{ iconName: 'Cancel' }}
          onClick={onDismiss}
        />
      </div>
      <Stack className={styles.body}>
        <div className={`${styles.msGrid} ${styles.noPadding} ${styles.width100}`} >
          <div className={`${styles.msGridrow} ${styles.marginTop30}`}>
            <div className={`${styles.msGridcol}  ${styles.mssm12}  `}>
              <SearchContainer contentTypeOptions={this.state.contentTypeOptions} searchResults={this.searchResult} selectedType={this.updateSelectedContentType} webpartContext={this.props.webpartContext}/>
            </div>
          </div>

          <ResultTable callBackSearchResults={this.finalSearchResult} rows={this.state.searchResult} contentType={this.state.selectedCType} onDismiss={this.props.onDismiss} />

        </div>
      </Stack>
    </Modal>;
  }

}