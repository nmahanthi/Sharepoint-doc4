/* eslint-disable no-throw-literal */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import styles from '../BaselineForms.module.scss';
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import * as React from "react";
import { ComboBox, PrimaryButton, TextField } from "@fluentui/react";
import { ISPColumn,ISearchParam ,IRequiredFields} from "../../../../interfaces/IGlobalInterfaces";
import CONSTANTS from '../../../../_Constants';
import { BaselineHelper } from "../../../../service/BaselineHelper";



export interface ISearchParameterRowProps {
  fieldOptions: ISPColumn[];
  addRow: () => void;
  deleteRow: (id: number) => void;
  callbackSearchParamChanged: (srchParam: ISearchParam, index: number) => void;
  index: number;
  showAdd: boolean;
  searchParam: ISearchParam;
}
export interface ISearchParameterRowState {
  operatorOptions: IDropdownOption[];
  selectedField?: ISPColumn;
  selectedOperator?: string;
  selectedCondition?: string;
  selectedstring?: string;
  errorMessage: string;
  errorControls: any[];
}
export default class SearchParameterRow extends React.Component<ISearchParameterRowProps, ISearchParameterRowState> {
  constructor(props: ISearchParameterRowProps) {
    super(props);
    this.state = {
      operatorOptions: [],
      selectedField: this.props.searchParam?.field,
      selectedOperator: this.props.searchParam?.operator,
      selectedCondition: this.props.searchParam?.condition,
      selectedstring: this.props.searchParam?.searchText,
      errorMessage: "",
      errorControls: []
    };
  }

  private deleteRow = (evt: any): void => {
    this.props.deleteRow(this.props.index);
  };


  //ON FIELD CHANGE EVENT
  private onFieldChanged = async (item: IDropdownOption) => {
    console.log('Field Changed:', item);
    const selectedField = this.props.fieldOptions.filter(f => f.id === item.key);


    this.resetError("field");

    if (selectedField.length > 0) {
      this.chooseOperator(selectedField[0].typeAsString);
      this.setState({
        selectedField: selectedField[0],
        selectedOperator: undefined,
        selectedCondition: undefined,
        selectedstring: undefined
      });

      const search: ISearchParam = this.props.searchParam;

      search.field = selectedField[0];
      search.searchFieldName = this.processSearchFieldName(selectedField[0]);

      this.props.callbackSearchParamChanged(search, this.props.index);
    }
    else {
      const search: ISearchParam = this.props.searchParam;

      search.field = undefined;
      search.condition = undefined;
      search.operator = undefined;
      search.searchText = undefined;

      this.setState({
        selectedField: undefined,
        selectedOperator: undefined,
        selectedCondition: undefined
      });

      this.props.callbackSearchParamChanged(search, this.props.index);
    }
  };

  private processSearchFieldName(field: ISPColumn) {

    const fieldName = field.internalName;

    if(field.title === "Title" || field.title === "Name")
      return "Title";
    if ( field.title.indexOf("Created") ===0  || field.title.indexOf("Modified") ===0)
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


  public validateFields = async (): Promise<any> => {

    const requiredFields: IRequiredFields[] = [
      { stateName: 'condition', fieldName: 'Condition' },
      { stateName: 'field', fieldName: 'Field' },
      { stateName: 'operator', fieldName: 'Operator' },
      { stateName: 'searchText', fieldName: 'Search Text' }

    ];
    // Required field validation
    try { BaselineHelper.checkRequiredFields(this.props.searchParam, requiredFields); }
    catch (e) {
      this.setState({
        errorControls: e.targetFields,
        errorMessage: e.message
      });
      throw 'Error';
    }

    return; // No validation issues.
  }
  // PICK OPERATORS BASED ON THE SELECTED COLUMN TYPE
  private chooseOperator(fieldType: string) {
    switch (fieldType) {
      case CONSTANTS.FieldTypes.Text:
      case CONSTANTS.FieldTypes.Choice:
      case CONSTANTS.FieldTypes.MultiChoice:
      case CONSTANTS.FieldTypes.Note:
      case CONSTANTS.FieldTypes.TaxonomyFieldType:
      case CONSTANTS.FieldTypes.TaxonomyFieldTypeMulti:
      case CONSTANTS.FieldTypes.URL:
      case CONSTANTS.FieldTypes.User:
      case CONSTANTS.FieldTypes.UserMulti:
        {
          const options: IDropdownOption[] = [
            // { key: '=', text: 'equals' },
            { key: '<>', text: 'not equals' },
            { key: ':', text: 'contains' },
            { key: '-', text: 'does not contains' },
            { key: ':text*', text: 'begins with' },
            { key: ':*text', text: 'ends with' }
          ];
          this.setState({ operatorOptions: options })
        }
        break;

      case CONSTANTS.FieldTypes.Boolean:
      case CONSTANTS.FieldTypes.DateTime:
        {
          const options: IDropdownOption[] = [
            // { key: ':', text: 'equals' },
            { key: "<>", text: "<>" },
            { key: "<", text: "<" },
            { key: "<=", text: "≤" },
            { key: ">", text: ">" },
            { key: ">=", text: "≥" },
          ];
          this.setState({ operatorOptions: options })
        }
        break;
      default:
        {
          const options: IDropdownOption[] = [
            // { key: '=', text: 'equals' },
            { key: '<>', text: 'not equals' },
            { key: ':', text: 'contains' },
            { key: '-', text: 'does not contains' },
            { key: ':text*', text: 'begins with' },
            { key: ':*text', text: 'ends with' }
          ];
          this.setState({ operatorOptions: options })
        }
        break;
    }
  }


  //ON CONDITION CHANGED
  private conditionChanged = (option: IDropdownOption, index: number): void => {

    const search: ISearchParam = this.props.searchParam;

    search.condition = option.key.toString();

    this.setState({ selectedCondition: option?.key.toString() });

    this.resetError("condition");

    this.props.callbackSearchParamChanged(search, this.props.index);

  }

  //ON OPERATOR CHANGED
  private operatorChanged = (option: IDropdownOption, index: number): void => {
    console.log('operator Field Changed:', option);
    const search: ISearchParam = this.props.searchParam;

    search.operator = option.key.toString();

    this.setState({ selectedOperator: option?.key.toString() });

    this.resetError("operator");

    this.props.callbackSearchParamChanged(search, this.props.index);

  }


  // TEXTFIELDS ON CHANGE
  private onTextChange = (value: string) => {

    const search: ISearchParam = this.props.searchParam;

    search.searchText = value;

    this.setState({ selectedstring: value });

    this.resetError("searchText");

    this.props.callbackSearchParamChanged(search, this.props.index);

  }

  private addNewRow = async () => {
    try {
      await this.validateFields();
    } catch (e) {
      console.log(e);
      return;
    }
    this.props.addRow();
  }

  //RESET ERROR CONTROLS
  public resetError(field: any) {
    const errors = this.state.errorControls;
    const index = errors.findIndex((obj) => obj.stateName === field);
    if (index >= 0) {
      errors.splice(index, 1)
      this.setState({ errorControls: errors, errorMessage: errors.length > 0 ? this.state.errorMessage : "" });
    }
  }

  componentDidMount(): void {
    if (this.props.searchParam.field) {
      this.chooseOperator(this.props.searchParam.field.typeAsString);
    }
  }
  async componentDidUpdate(prevProps: ISearchParameterRowProps) {
    if (this.props.searchParam !== prevProps.searchParam) {
      if (this.props.searchParam.field) {
     //   if (this.state.operatorOptions.length === 0)
          this.chooseOperator(this.props.searchParam.field.typeAsString);
      }
    }
  }

  public render(): React.ReactElement<{}> {
  
    const { searchParam } = this.props;
    const conditionOptions: IDropdownOption[] = [
      { key: 'AND', text: 'AND' },
      { key: 'OR', text: 'OR' }
    ];
    return <>
      < div className={`${styles.msGridrow}`}>
        <div className={`${styles.msGridcol}  ${styles.mssm1}  `}>
          <label className={styles.label}>(</label>
        </div>
        <div className={`${styles.msGridcol}  ${styles.mssm6}  `}>
       
<ComboBox placeholder="Select a field" allowFreeform
            autoComplete="on"
            options={this.props.fieldOptions.sort((a, b) => a.title.localeCompare(b.title)).map((r) => {
              return { key: r.id, text: r.title };
            })}
            className={this.state.errorControls.some(e => e.stateName === 'field') ? styles.required : ''}
            onChange={(event,option:IDropdownOption) => this.onFieldChanged(option)} selectedKey={searchParam?.field?.id ? searchParam?.field?.id : ""}
          />
        
        </div>
        {/*  <div className={`${styles.msGridcol}  ${styles.mssm1}  `}>
        <label className={styles.label}>-</label>

      </div>*/}
        <div className={`${styles.msGridcol}  ${styles.mssm3}  `}>
          <Dropdown placeholder="Operator"
            options={this.state.operatorOptions} 
            selectedKey={searchParam.operator}
            onChanged={this.operatorChanged}
            className={this.state.errorControls.some(e => e.stateName === 'operator') ? styles.required : ''}

          />

        </div>

        <div className={`${styles.msGridcol}  ${styles.mssm3}  `}>
          <TextField placeholder='search string' value={searchParam.searchText}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => this.onTextChange(newValue)}
            className={this.state.errorControls.some(e => e.stateName === 'searchText') ? styles.required : ''}

          />
        </div>

        <div className={`${styles.msGridcol}  ${styles.mssm1}  `}>
          <label className={styles.label}>)</label>
        </div>

        <div className={`${styles.msGridcol}  ${styles.mssm3}  `}>
          <Dropdown placeholder="and / or"
            options={conditionOptions} selectedKey={searchParam.condition}
            onChanged={this.conditionChanged}
            className={this.state.errorControls.some(e => e.stateName === 'condition') ? styles.required : ''}
          />
        </div>

        <div className={`${styles.msGridcol}  ${styles.mssm2}  `}>
          {this.props.showAdd ?
            <PrimaryButton text="Add" onClick={this.addNewRow} />
            :
            <PrimaryButton text="Del" onClick={this.deleteRow} />
          }
        </div>
      </div >
    </>
  }

}