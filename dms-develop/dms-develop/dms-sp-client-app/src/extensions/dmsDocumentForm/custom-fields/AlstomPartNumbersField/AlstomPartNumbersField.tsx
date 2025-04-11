
/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from 'react';
import DmsDynamicFieldBase, { IDmsDynamicFieldProps, IDmsDynamicFieldState } from '../../components/DmsDynamicFieldBase';
import { Logger } from '@pnp/logging';
import {Button} from "@fluentui/react-components";
import { AddRegular,SubtractRegular } from '@fluentui/react-icons';
import styles from './AlstomPartNumbersField.module.scss';
import { Label } from "@fluentui/react-components";
interface IPartNumber {
  value?: string;
  reference: string;
  revision: string;
}
interface IAlstomPartNumbersFieldState extends IDmsDynamicFieldState {
  partNumbers: IPartNumber[]; 
  // value: string; 
}

const LOG_SOURCE: string = 'AlstomPartNumbersField';
export default class AlstomPartNumbersFieldextends extends DmsDynamicFieldBase< IAlstomPartNumbersFieldState> {

  constructor(props: IDmsDynamicFieldProps) {
    super(props);
    this.state = {
      ...this.state,
      partNumbers: this.parseValueToPartNumbers(props.data) 

    };
  }

  public componentDidMount(): void {
    const { fieldInfo, onError } = this.props;
    if (super.componentDidMount) super.componentDidMount();
    this.initializeField().catch(error => {
      Logger.error(new Error(`${LOG_SOURCE}: Error initializing field ${fieldInfo.Title}, ${error}`));
      console.error(error);
      this.setState({ isErrored: true, isLoading: false });
      if (onError) onError(new Error(`Error initializing field ${fieldInfo.Title}: ${error}`));
    });
  }

  parseValueToPartNumbers = (value: any): IPartNumber[] => {
    if (!value) return [{ reference: '', revision: '' }]; 
    return value.split(';').map((part: { split: (arg0: string) => [any, any]; }) => {
      const [reference, revision] = part.split('/');
      return { reference: reference.trim(), revision: revision?.trim() || '' }; 
    });
  };

  handleInputChange = (index: number, field: keyof IPartNumber, value: string) => {
    const updatedParts = [...this.state.partNumbers];
    updatedParts[index][field] = value;
    this.setState({ partNumbers: updatedParts }, this.notifyParent); 
  };

  notifyParent = () => {
    if (this.props.onChange) {
      const singleLineValue = this.partNumbersToSingleLine();
      this.props.onChange(singleLineValue)
      this.setState({value:singleLineValue})
    }
   
  };

  partNumbersToSingleLine = (): string => {
    return this.state.partNumbers
      .filter(part => part.reference && part.revision) 
      .map(part => `${part.reference}/${part.revision}`) 
      .join('; '); 
  };

  // Add new row for part numbers
  handleAddRow = () => {
    this.setState(prevState => ({
      partNumbers: [...prevState.partNumbers, { reference: '', revision: '' }]
    }));
  };

  handleRemoveRow = (index: number) => {
    this.setState(prevState => ({
      partNumbers: prevState.partNumbers.filter((_, i) => i !== index)
    }));
  };

    // Concatenate part numbers and revisions into a single string
    getConcatenatedPartNumbers = (): string => {
      return this.state.partNumbers
        .filter(part => part.reference && part.revision) 
        .map(part => `${part.reference}/${part.revision}`)
        .join('; ');
    };

    isRowFilled = (part: IPartNumber): boolean => {
      return part.reference !== '' && part.revision !== '';
    };
  public render(): React.ReactElement<{}> {
    const { partNumbers} = this.state;
    const { fieldInfo, disabled: propsDisabled  } = this.props;
    const disabled = propsDisabled || fieldInfo.ReadOnlyField;
    return (
      <div className={styles.AlstomPartNumbersField}>
         <Label size="large" className={styles.AlstomLabel}>{fieldInfo.Title}</Label> 
        <div className={styles.PartNumbersContainer}>
        {partNumbers.map((part, index) => (
          <div key={index} className="part-row" >
            <input
            className={styles.ReferenceInputBox}
              type="text"
              placeholder="Reference"
              value={part.reference}
              onChange={(e) => this.handleInputChange(index, 'reference', e.target.value)}
              required
              disabled={disabled}
            />
            <input
              type="text"
              placeholder="Revision"
              value={part.revision}
              onChange={(e) => this.handleInputChange(index, 'revision', e.target.value)}
              required
              disabled={disabled}
            />
            
            {partNumbers.length > 1 && !disabled && (
              <>{ <Button size="small" className={styles.CusButtonStyles} icon={<SubtractRegular className={styles.CusbtnIcon}/>} onClick={() => this.handleRemoveRow(index)} />}</>
            )}
             {this.isRowFilled(part) && index === partNumbers.length - 1 && !disabled && (
              <><Button size="small"  className={styles.CusButtonStyles}   icon={<AddRegular className={styles.CusbtnIcon} />} onClick={this.handleAddRow} /></>
            )}
          </div>
        ))}
        </div>

      </div>
    );
  } 
}