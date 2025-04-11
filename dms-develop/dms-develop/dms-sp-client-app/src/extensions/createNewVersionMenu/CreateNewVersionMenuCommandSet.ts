import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import ReactDOM from 'react-dom';
import { CreateNewVersion } from '../baselineForms/components/CreateNewVersion/CreateNewVersion';
import React from 'react';
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import { SPFI } from '@pnp/sp/fi';
import { getSP } from '../../pnpjs-config';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICreateNewVersionMenuCommandSetProperties {
  sampletextone: string;
  samppletexttwo: string;
}

const LOG_SOURCE: string = 'CreateNewVersionMenuCommandSet';

export default class CreateNewVersionMenuCommandSet extends BaseListViewCommandSet<ICreateNewVersionMenuCommandSetProperties> {

  private dialogContainer: HTMLElement | null = null;
  private isNewVersion: boolean = true;
  private _sp: SPFI;

  private _newVersion(): void {
    this.isNewVersion = !this.isNewVersion
  }
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CreateNewVersionMenuCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('CreateVersionCommand');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'CreateVersionCommand':

        this._sp = getSP(this.context);
        try {

          const selectedRowId = event.selectedRows[0].getValueByName("ID");

          const itemId = Number(selectedRowId);
          const item: any = await this._sp.web.lists.getByTitle('Baselines').items.getById(itemId)();
          this.showCreateNewVersionDialog(item);

        } catch (error: any) {
          console.error("Error fetching item:", error);
        }

        break;

      default:
        throw new Error('Unknown command');
    }
  }

  private showCreateNewVersionDialog(item: any): void {
    this.dialogContainer = document.createElement('div');
    document.body.appendChild(this.dialogContainer);

    const element = React.createElement(CreateNewVersion, {
      item: item,
      context: this.context,
      isNewVersion: this.isNewVersion,
      updateVersion: this._newVersion.bind(this)
    });

    ReactDOM.render(element, this.dialogContainer);
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('CreateVersionCommand');

    if (compareOneCommand) {
      const selectedRows = this.context.listView.selectedRows;

      if (Array.isArray(selectedRows) && selectedRows.length === 1) {
        const documentStatus = selectedRows[0]._values.get("BaselineStatus");

        compareOneCommand.visible = documentStatus === "Frozen";
      } else {
        compareOneCommand.visible = false;
      }
    }
    // TODO: Add your logic here

    this.raiseOnChange();
  }
}
