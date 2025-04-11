import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { FieldNames } from "../constants";
import { AssignFrom } from "@pnp/core";
import { spGet, SPQueryable } from "@pnp/sp";

// import { getSPFromPageContext } from "../pnpjs-config";
import { getSPFromPageContextCacheNever } from "../pnpjs-config";
import { BaseListViewCommandSet, ListViewStateChangedEventArgs } from "@microsoft/sp-listview-extensibility";
export interface ISelectedItemService {
    onListViewStateChanged(event: ListViewStateChangedEventArgs): void;
    getSelectedItemData(): Promise<{
        selectedItemId: number;
        userName: string;
        userEmail: string;
        docIDextracted: string;
        createdByobj: { EMail: string; Title: string; UserName: string; ID: number };
        docTitle: string;
        itemLink: string;
        projReference: string;
        projRevision: string;
        createdBy: string;
        docName: string;
        docStatus: string;
        contentTypeId: string;
        childItemsCount: number;
    } | undefined>;
    init(commandSet: BaseListViewCommandSet<unknown>): void;
}
export class SelectedItemService implements ISelectedItemService {

    public static readonly serviceKey: ServiceKey<ISelectedItemService> =
        ServiceKey.create<ISelectedItemService>('dms:ISelectedItemService', SelectedItemService);
    private _sp: SPFI;
    private _context: PageContext;
    private _commandSet: BaseListViewCommandSet<unknown>;
    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._context = serviceScope.consume(PageContext.serviceKey);
            this._sp = getSPFromPageContextCacheNever(this._context);
        });
    }
    public init(commandSet: BaseListViewCommandSet<unknown>): void {
        this._commandSet = commandSet;
        commandSet.context.listView.listViewStateChangedEvent.add(commandSet, this.onListViewStateChanged.bind(this));
    }
    public onListViewStateChanged(event: ListViewStateChangedEventArgs): void {
        if (this._commandSet.context.listView.selectedRows?.length === 1 &&
            this._commandSet.context.pageContext.list
        ) {
            this._getData(this._commandSet.context.listView.selectedRows[0].getValueByName("ID"), this._commandSet.context.pageContext.list.id.toString())
                .catch(error => {
                    console.error('Error fetching item data:', error);
                });
        }
    }
    public getSelectedItemData(): Promise<{
        selectedItemId: number;
        userName: string;
        userEmail: string;
        docIDextracted: string;
        createdByobj: { EMail: string; Title: string; UserName: string; ID: number; };
        docTitle: string;
        itemLink: string;
        projReference: string;
        projRevision: string;
        createdBy: string;
        docName: string;
        docStatus: string;
        contentTypeId: string;
        childItemsCount: number;
    } | undefined> {
        const selectedRows = this._commandSet.context.listView.selectedRows;
        if (selectedRows?.length !== 1 || !this._commandSet.context.pageContext.list) {
            return Promise.resolve(undefined);
        }
        return this._getData(selectedRows[0].getValueByName('ID'), this._commandSet.context.pageContext.list.id.toString());
    }

    private _data: {
        [id: number]: {
            selectedItemId: number;
            userName: string;
            userEmail: string;
            docIDextracted: string;
            createdByobj: { EMail: string; Title: string; UserName: string; ID: number };
            docTitle: string;
            itemLink: string;
            projReference: string;
            projRevision: string;
            createdBy: string;
            docName: string;
            docStatus: string;
            contentTypeId: string;
            childItemsCount: number;
        }
    } = {};

    private async _getData(id: number, listId: string): Promise<{
        selectedItemId: number;
        userName: string;
        userEmail: string;
        docIDextracted: string;
        createdByobj: { EMail: string; Title: string; UserName: string; ID: number };
        docTitle: string;
        itemLink: string;
        projReference: string;
        projRevision: string;
        createdBy: string;
        docName: string;
        docStatus: string;
        contentTypeId: string;
        childItemsCount: number;
    }> {
        const serverUrl = this._commandSet.context.pageContext.site.absoluteUrl.replace(this._commandSet.context.pageContext.site.serverRelativeUrl, '');
        //  if (!this._data[id]) {
        const item = await this._sp.web.lists.getById(listId).items.getById(id)
            .select('OData__dlc_DocId', 'ContentTypeId', 'Id', 'Author/EMail', 'Author/Title', 'Folder/ItemCount',
                'Author/UserName', 'Author/ID', 'Title', 'FileLeafRef', 'FileRef', FieldNames.ProjectReference,FieldNames.Revision, FieldNames.DocumentStatus)
            .expand('Author', 'Folder')();

        const querry = `${this._commandSet.context.pageContext.site.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${item.FileRef}')/files`
        const spQueryable = SPQueryable(querry).using(AssignFrom(this._sp.web));
        const fileItems: unknown[] = await spGet(spQueryable);
        const itemsCount = fileItems.length;

        this._data[id] = {
            userEmail: this._commandSet.context.pageContext.user.email,
            userName: this._commandSet.context.pageContext.user.displayName,
            docIDextracted: item.OData__dlc_DocId,
            createdByobj: item.Author,
            docTitle: item.Title,
            projReference: item[FieldNames.ProjectReference],
            projRevision: item[FieldNames.Revision],
            selectedItemId: item.Id,
            itemLink: `${serverUrl}/${item.FileRef}`,
            createdBy: item.Author.Title,
            docName: item.FileLeafRef,
            docStatus: item[FieldNames.DocumentStatus],
            contentTypeId: item.ContentTypeId,
            childItemsCount: itemsCount
        };
        //  }
        return this._data[id];
    }
}