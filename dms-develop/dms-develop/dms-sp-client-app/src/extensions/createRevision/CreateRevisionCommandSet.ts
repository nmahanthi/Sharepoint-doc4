import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import _CONSTANTS from '../../_Constants';
import CreateRevisionDialogContainer from '../../components/CreateRevision/createRevisionDialogContainer';
import { getSP } from '../../pnpjs-config';

export interface ICreateRevisionCommandSetProperties {
  ApplicableListTitle: string;
  DraftListTitle: string;
  PermissionListTitle: string;
  ContentTypeName: string;
}

const LOG_SOURCE: string = 'CreateRevisionCommandSet';

export default class CreateRevisionCommandSet extends BaseListViewCommandSet<ICreateRevisionCommandSetProperties> {

  private _sp: SPFI;
  private _isContributor: boolean;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CreateRevisionCommandSet');

    // initial state of the command's visibility
    const requestRevisionCommand: Command = this.tryGetCommand('REQ_Revision');
    requestRevisionCommand.visible = false;
    this._sp = getSP(this.context);
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'REQ_Revision':
        if(this.context.listView.selectedRows){
          const projReference = this.context.listView.selectedRows[0].getValueByName(_CONSTANTS.TechnicalDocumentsFields.ProjectReference);
          const docSetServerRelativeUrl = this.context.listView.selectedRows[0].getValueByName(_CONSTANTS.CommonFields.FileRef);
          const dialog = new CreateRevisionDialogContainer(projReference, this._sp, docSetServerRelativeUrl);
          dialog.show().then(() => {
            console.log("Dialog opened");
          }).catch((err) => {
            console.error("Dialog error: ", err);
          });
          
        } else {
          throw new Error("Select an item");
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async checkIfContributor(sp: SPFI): Promise<boolean> {
    const { PermissionListTitle } = this.properties;
    const items = await sp.web.lists.getByTitle(PermissionListTitle).items.select("Title").filter("Title eq '"+ _CONSTANTS.PermissionListValues.Contributor +"'")();
    return (items && items.length>0);
  }

  private _onListViewStateChanged = async (args: ListViewStateChangedEventArgs): Promise<void> => {
    Log.info(LOG_SOURCE, 'List view state changed');
    const {ContentTypeName, ApplicableListTitle } = this.properties;
    let isTechnicalDocument:boolean = false;
    let isApplicableLibrary: boolean = false;
    try{
      if(this.context.listView.selectedRows?.length === 1) {
        const ctValue = this.context.listView.selectedRows[0].getValueByName(_CONSTANTS.CommonFields.ContentTypeFieldName);
        if(ctValue && ctValue === ContentTypeName){
          isTechnicalDocument = true;
        }
      }

      if(this.context.listView.list && this.context.listView.list.title == ApplicableListTitle){
        isApplicableLibrary = true;
      }

      if(!this._isContributor){
          this._isContributor = await this.checkIfContributor(this._sp);
      }
    } catch(err){
      console.log(err);
    }
    const requestRevisionCommand: Command = this.tryGetCommand('REQ_Revision');
    if (requestRevisionCommand && this.context.listView.selectedRows?.length === 1 && isTechnicalDocument && this._isContributor && isApplicableLibrary) {
      requestRevisionCommand.visible = true;
    } else {
      requestRevisionCommand.visible = false;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
