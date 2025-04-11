import { SPFI } from "@pnp/sp/index.js";
import { pnpjs } from "./PnpjsService.js";

export const ProjectPropertiesList = {
    RelativeUrl : "/Lists/ProjectProperties",
    NameField: "Title",
    ValueField: "Value"
  };
export interface IProjectProperties {
    dms_create_date?: string;
    dms_proj_type?: string;
    dms_modi_date?: string;
    dms_proj_lang?: string;
    dms_prod_line_display?: string;
    dms_proj_alias?: string;
    dms_pl_owner_display?: string;
    dms_modi_display?: string;
    dms_create_email?: string;
    dms_proj_name?: string;
    dms_proj_code?: string;
    dms_end_cust_display?: string;
    dms_end_cust_guid?: string;
    dms_create_display?: string;
    dms_dir_cust_display?: string;
    dms_dir_cust_guid?: string;
    dms_is_active?: string;
    dms_prod_display?: string;
    dms_geo_display?: string;
    dms_pl_owner_guid?: string;
    dms_geo_guid?: string;
    dms_is_old_proj?: string;
    dms_modi_email?: string;
    dms_prod_guid?: string;
    MassImportStatusFieldCustomizerProperties?: string;
}

export class ProjectPropertiesService {
    private _sp: SPFI;
    private _listIdValue: string;
    private get _listId(): Promise<string> {
        if (this._listIdValue) {
            return Promise.resolve(this._listIdValue);
        }
        return this._sp.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${this._serverRelativeUrl}${ProjectPropertiesList.RelativeUrl}'`)().then((list) => {
            if (list.length === 0) {
                return Promise.reject(new Error('List not found'));
            }
            this._listIdValue = list[0].Id;
            return this._listIdValue;
        });
    }
    constructor(private _serverRelativeUrl: string) {
        this._sp = pnpjs.sp(_serverRelativeUrl);
    }

    public async getProperties(): Promise<IProjectProperties> {
        const listId = await this._listId;
        const items = await this._sp.web.lists.getById(listId).items.select('Title', 'Value').orderBy('Modified', false)();
        const names = items.map((item) => item.Title);
        const uniqueItems = items.filter((item, index) => names.indexOf(item.Title) === index);
        return uniqueItems.reduce((properties, item) => {
            properties[item.Title] = item.Value || undefined;
            return properties;
        }, {} as IProjectProperties);
    }
    public async getProperty(propertyName: keyof IProjectProperties): Promise<string | undefined> {
        const listId = await this._listId;
        const items = await this._sp.web.lists.getById(listId).items
            .select('Value')
            .filter(`Title eq '${propertyName}'`)
            .orderBy('Modified', false)
            .top(1)();
        if (items.length === 0) {
            return undefined;
        }
        return items[0].Value;
    }
    public async setProperty(propertyName: keyof IProjectProperties, value: string): Promise<void> {
        const listId = await this._listId;
        const items = await this._sp.web.lists.getById(listId).items
            .select('Id')
            .filter(`Title eq '${propertyName}'`)
            .orderBy('Modified', false)
            .top(1)();
        if (items.length === 0) {
            await this._sp.web.lists.getById(listId).items.add({
                Title: propertyName,
                Value: value
            });
        } else {
            await this._sp.web.lists.getById(listId).items.getById(items[0].Id).update({ Value: value });
        }
    }
    public async setProperties(properties: Partial<Record<keyof IProjectProperties, string>>): Promise<void> {
        const listId = await this._listId;
        const items = await this._sp.web.lists.getById(listId).items.select('Id', 'Title').orderBy('Modified', false)();
        const names = items.map((item) => item.Title);
        const uniqueItems = items.filter((item, index) => names.indexOf(item.Title) === index);
        const itemsToUpdate = uniqueItems.filter((item) => Object.prototype.hasOwnProperty.call(properties, item.Title));
        const itemsToAdd = Object.entries(properties).filter(([key]) => !names.includes(key as string)).map(([key, value]) => ({ Title: key, Value: value }));
        const [batch, execute] = this._sp.batched();
        const updatePromises = itemsToUpdate.map((item) =>
            batch.web.lists.getById(listId).items.getById(item.Id)
                .update({ Value: properties[item.Title as keyof IProjectProperties] })
        );
        const addPromises = itemsToAdd.map((item) =>
            batch.web.lists.getById(listId).items.add(item)
        );
        await execute();
        await Promise.all([...updatePromises, ...addPromises]);
    }
}