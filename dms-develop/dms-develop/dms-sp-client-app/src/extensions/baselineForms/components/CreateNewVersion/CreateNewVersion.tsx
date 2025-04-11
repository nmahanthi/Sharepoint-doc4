import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar'; // Import MessageBar
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../../pnpjs-config';
import { RenderListDataOptions } from '@pnp/sp/lists';

export interface IUnderstandFunctionComponentProps {
  item: any;
  context: any;
  isNewVersion: boolean;
  updateVersion: () => void
}

interface ICreateNewVersionState {
  isDialogVisible: boolean;
  formData: {
    newVersion: string;
  };
  errors: {
    newVersion: boolean;
    newVersionMessage: string;
  };
  successMessage: string;
  errorMessage: string;
}

export class CreateNewVersion extends React.Component<IUnderstandFunctionComponentProps, ICreateNewVersionState> {
  private _sp: SPFI;

  constructor(props: IUnderstandFunctionComponentProps) {
    super(props);
    this.state = {
      isDialogVisible: this.props.isNewVersion,
      formData: {
        newVersion: '',
      },
      errors: {
        newVersion: false,
        newVersionMessage: '',
      },
      successMessage: '',
      errorMessage: '',
    };
    this._sp = getSP(this.props.context);
  }

  openDialog = () => {
    this.setState({ isDialogVisible: true, successMessage: '', errorMessage: '' });
  };

  closeDialog = () => {
    this.setState({ isDialogVisible: false });
    this.props.updateVersion();

  };



  handleInputChange = (value: string) => {

    this.setState({
      formData: { newVersion: value },
      errors: { newVersion: value.trim() === '', newVersionMessage: '' },
    });
  };

  private handleFormSubmit = async () => {
    const { formData } = this.state;
    const { item } = this.props;


    // Validation: Check if the new version field is empty or contains special characters
    const isEmpty = formData.newVersion.trim() === '';
    const hasInvalidChars = !/^[a-zA-Z0-9\-_.]+$/.test(formData.newVersion); // Allow alphanumeric, hyphen, and underscore
    //const hasInvalidVersionCheck = await this._sp.web.lists.getByTitle('Baselines').items.filter(`Code eq '${item.Code}' and BaselineVersion eq '${formData.newVersion.trim()}'`)();
    //Promise.all(hasInvalidVersionCheck);

    const hasInvalidVersionCheck = await this._sp.web.lists.getByTitle('Baselines').renderListDataAsStream({
        ViewXml: `<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='Code' /><Value Type='Text'>${item.Code}</Value></Eq><Eq><FieldRef Name='BaselineVersion' /><Value Type='Text'>${formData.newVersion.trim()}</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields><RowLimit>1</RowLimit></View>`,
        RenderOptions: RenderListDataOptions.ListData,
    });

    let hasInvalidVersion: boolean = false;
    debugger;
    if (hasInvalidVersionCheck && hasInvalidVersionCheck.Row.length > 0) {
      hasInvalidVersion = true;
    }
    const newErrors = {
      newVersion: isEmpty || hasInvalidChars || hasInvalidVersion,
      newVersionMessage: isEmpty
        ? 'This field is required'
        : hasInvalidChars
          ? 'Special characters are not allowed'
          : hasInvalidVersion
            ? 'This version already exists'
            : '',
    };

    this.setState({ errors: newErrors });


    if (!newErrors.newVersion) {
      try {

        const payload = {
          Title: item.Title || '',
          CRID: item.CRID || '',
          Code: item.Code || '',
          BaselineStatus: 'Ongoing',
          BaselineComments: item.BaselineComments || '',
          BaselineVersion: formData.newVersion,
        };


        // Create a new item in the 'Baselines' list
        const addItemResponse = await this._sp.web.lists.getByTitle('Baselines').items.add(payload);


      
        this.setState({ successMessage: 'New Baseline version successfully created.', errorMessage: '' }); // Set success message


        await this.handleBaselineDocuments(item.Id, addItemResponse.ID);
        // Call the method to handle fetching and creating items in the "Baseline Baseline" list
        await this.handleBaselineBaselines(item.Id, addItemResponse.ID);

        this.closeDialog();
        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/Lists/Baselines/AllItems.aspx`;
      } catch (error) {
        console.error('Error creating new item:', error);
        const errorMessage = `An error occurred while creating the item: ${error.message || 'Unknown error'}`;
      
        this.setState({ errorMessage });

        //alert(`An error occurred while creating the item: ${error.message || 'Unknown error'}`);
      }
    }
  };
  // Method to handle fetching and creating items in the "Baseline documents" list
  handleBaselineDocuments = async (oldBaselineId: number, newBaselineId: number) => {
    try {

      const documents = await this._sp.web.lists.getByTitle('Baseline documents').items
        .filter(`BaselineParentID eq ${oldBaselineId}`)
        .select('DocumentID', 'Title', 'SiteURL') // Chaining methods correctly
        ();
      // Loop through each fetched document and create a new item with the new BaselineParentID
      for (const doc of documents) {
        const { DocumentID, Title, SiteURL } = doc;
        



        const newBaselineParentID = String(newBaselineId);
        // Create new item in 'Baseline Documents' list with fetched values and new BaselineParentID

        await this._sp.web.lists.getByTitle('Baseline Documents').items.add({
          DocumentID: DocumentID,
          Title: Title,
          SiteURL: SiteURL,
          BaselineParentID: newBaselineParentID, // Set new BaselineParentID to the newly created Baseline item
        });

        
      }
    } catch (error) {
      console.error('Error handling Baseline Documents:', error);
      alert(`An error occurred while processing Baseline Documents: ${error.message || 'Unknown error'}`);
    }
  };
  // Method to handle fetching and creating items in the "Baseline Baseline" list
  handleBaselineBaselines = async (oldBaselineId: number, newBaselineId: number) => {
    try {
      // Fetch data from 'Baseline Baseline' list where 'BaselineParentID' equals the old item's ID
      const baselines = await this._sp.web.lists.getByTitle('Baseline Baseline').items
        .filter(`BaselineParentID eq '${oldBaselineId}'`) // Use string interpolation for filtering
        .select('BaselineChildID', 'Title', 'SiteURL') // Only select needed fields
        ();

      // Loop through each fetched baseline and create a new item with the new BaselineParentID
      for (const baseline of baselines) {
        const { BaselineChildID, Title, SiteURL } = baseline;


        const newBaselineParentID = String(newBaselineId);

        // Create new item in 'Baseline Baseline' list with fetched values and new BaselineParentID
        await this._sp.web.lists.getByTitle('Baseline Baseline').items.add({
          BaselineChildID: BaselineChildID,
          Title: Title,
          SiteURL: SiteURL,
          BaselineParentID: newBaselineParentID
        });

        
      }
    } catch (error) {
      console.error('Error handling Baseline Baselines:', error);
      alert(`An error occurred while processing Baseline Baselines: ${error.message || 'Unknown error'}`);
    }
  };



  render() {
    const { errors, successMessage, errorMessage } = this.state;
    
    return (
      <div>

        {/* <PrimaryButton text="Create New Version" onClick={this.openDialog} style={{ marginBottom: '20px' }} /> */}

        <Dialog
          hidden={!this.props.isNewVersion}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Create New Version',
          }}
          modalProps={{
            isBlocking: true,
          }}
        >
          {/* Show success message */}
          {successMessage && (
            <MessageBar
              messageBarType={MessageBarType.success}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >
              {successMessage}
            </MessageBar>
          )}
          {/* Show error message */}
          {errorMessage && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >
              {errorMessage}
            </MessageBar>
          )}
          {/*New Version" field*/}
          <TextField
            label="New Version"
            required
            value={this.state.formData.newVersion}
            errorMessage={errors.newVersion ? errors.newVersionMessage : undefined}
            onChange={(event, newValue) => this.handleInputChange(newValue || '')}
          />

          <DialogFooter>
            <PrimaryButton text="Save" onClick={this.handleFormSubmit} />
            <DefaultButton text="Cancel" onClick={this.closeDialog} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}