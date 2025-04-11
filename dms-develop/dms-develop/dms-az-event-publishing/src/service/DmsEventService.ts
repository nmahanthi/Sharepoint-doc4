import { pnpjs } from "az-common/dist/index.mjs";
import { IDmsEvent, ISPChange, ISpEvent, SPChangeType } from "../model/index.js";
import { IChangeTokenService } from "./ChangeTokenService.js";
import { SPFI } from "@pnp/sp";
import { RenderListDataOptions } from "@pnp/sp/lists/types.js";
import { AppSettings } from "../AppSettings.js";

const LOG_SOURCE = "DmsEventService";
export class DmsEventService {

    constructor(private _changeTokenService: IChangeTokenService) { }

    public async getDmsEventsFromSpEvent(spEvent: ISpEvent, additionalFields?: string[]): Promise<{ events: IDmsEvent[], lastToken: string, listId: string; }> {
        const sp: SPFI = pnpjs.sp(spEvent.siteUrl);
        const changeToken = await this._changeTokenService.GetChangeToken(spEvent.resource);
        const changes: ISPChange[] = await sp.web.lists.getById(spEvent.resource).getChanges({
            Add: true,
            Update: true,
            Rename: true,
            Item: true,
            RoleAssignmentAdd: true,
            RoleAssignmentDelete: true,
            ChangeTokenStart: changeToken ? {
                StringValue: changeToken
            } : undefined
        });
        const additionalData = await this._getAdditionalData(sp, changes.map(c => c.ItemId), spEvent.resource);
        const itemsData = await this._getItemsDataAsStream(sp, changes.map(c => c.ItemId), spEvent.resource, additionalFields);
        const siteServerRelariveUrl = await sp.web.select("ServerRelativeUrl")();
        const messageIncludedFields = [...AppSettings.MessageIncludedFields, ...(additionalFields || [])];
        const events: IDmsEvent[] = changes.map(c => ({
            siteUrl: spEvent.siteUrl,
            siteServerRelativeUrl: siteServerRelariveUrl.ServerRelativeUrl,
            libraryId: spEvent.resource,
            itemId: c.ItemId,
            user: additionalData[c.ItemId].editor,
            modified: additionalData[c.ItemId].modified,
            spContentType: additionalData[c.ItemId].contentType,
            action: this._mapChangeType(c.ChangeType),
            fields: itemsData[c.ItemId] ? messageIncludedFields.reduce((acc, cur) => {
                acc[cur] = itemsData[c.ItemId][cur];
                return acc;
            }, {}) : undefined
        })).filter(e => !!e.action && !!e.user);
        return { events, lastToken: changes.length > 0 ? changes[changes.length - 1].ChangeToken.StringValue : changeToken, listId: spEvent.resource };
    }
    private async _getAdditionalData(sp: SPFI, itemIds: number[], listId: string): Promise<{ [itemId: number]: { editor: string; modified: string, contentType: string; } }> {
        const results: { [itemId: number]: { editor: string; modified: string, contentType: string; } } = {};
        //split in chunks of 100 items
        const chunks = itemIds.reduce((acc, cur, idx) => {
            const chunkIndex = Math.floor(idx / 100);
            if (!acc[chunkIndex]) {
                acc[chunkIndex] = [];
            }
            acc[chunkIndex].push(cur);
            return acc;
        }, [] as number[][]);
        const select = ['Editor/EMail', 'Id', 'Modified', 'ContentType/Name'];
        for (const chunk of chunks) {
            const [batch, execute] = sp.batched();
            Promise.all(
                chunk.map(id => {
                    return batch.web.lists.getById(listId).items
                        .getById(id).select(...select).expand("Editor", "ContentType")()
                        .then(i => ({ id: i.Id, editor: i.Editor.EMail, modified: i.Modified, contentType: i.ContentType.Name }))
                        .catch(err => {
                            if (err.message.indexOf('404') > -1) {
                                return { id, editor: null, modified: null, contentType: null };
                            }
                            throw new Error(err);
                        })
                })
            ).then(res => res.forEach(r => results[r.id] = r));
            await execute();
        }
        return results;
    }
    private _getItemsDataAsStream(sp: SPFI, itemIds: number[], listId: string, additionalFields?: string[]): Promise<{ [itemId: number]: { [field: string]: string } }> {
        const [batch, execute] = sp.batched();
        //Split the array in chunks of 100 items
        const chunks: number[][] = itemIds.reduce((acc, cur, idx) => {
            const chunkIndex = Math.floor(idx / 100);
            if (!acc[chunkIndex]) {
                acc[chunkIndex] = [];
            }
            acc[chunkIndex].push(cur);
            return acc;
        }, [] as number[][]);
        const results = Promise.all(chunks.map(chunk => {
            const xmlFilter = chunk.slice(1).reduce((acc, cur) => {
                return `<Or>${acc}<Eq><FieldRef Name='ID' /><Value Type='Counter'>${cur}</Value></Eq></Or>`;
            }, `<Eq><FieldRef Name='ID' /><Value Type='Counter'>${chunk[0]}</Value></Eq>`);
            const messageIncludedFields = [...AppSettings.MessageIncludedFields, ...(additionalFields || [])];
            const viewFieldsXml = messageIncludedFields.reduce((acc, cur) => {
                return `${acc}<FieldRef Name='${cur}' />`;
            }, '');
            return batch.web.lists.getById(listId).renderListDataAsStream({
                ViewXml: `<View Scope='RecursiveAll'><Query><Where>${xmlFilter}</Where></Query><ViewFields>${viewFieldsXml}<FieldRef Name='ID' /></ViewFields></View>`,
                RenderOptions: RenderListDataOptions.ListData
            });
        })).then(allResults => {
            return allResults.flatMap(r => r.Row).flatMap(r => r).reduce((acc, cur) => ({ ...acc, [cur.ID]: cur }), {} as { [itemId: number]: { [field: string]: string } });
        });
        execute().catch(err => {
            throw new Error(err);
        });
        return results;
    }
    private _mapChangeType(ChangeType: SPChangeType): 'create' | 'update' | 'delete' {
        switch (ChangeType) {
            case SPChangeType.Add:
                return 'create';
            case SPChangeType.Update:
                return 'update';
            case SPChangeType.DeleteObject:
                return 'delete';
            default:
                return null;
        }
    }
}