import { SPFI } from "@pnp/sp";
import { ILogHistoryItem } from "../interfaces/ILogHistoryItem";
import { ILogHistoryService } from "../interfaces/ILogHistoryService";
import { AadHttpClient } from "@microsoft/sp-http";
import { FieldCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { getSP } from "../pnpjs-config";
import { ListUrls } from "../constants";
export class LogHistoryService implements ILogHistoryService {

    private _dmsClient: AadHttpClient;
    private _sp: SPFI;
    private context: ListViewCommandSetContext | FieldCustomizerContext;
    constructor(dmsClient: AadHttpClient) {
        this._sp = getSP(this.context);
        this._dmsClient = dmsClient;

    }
    async getLogHistoryItems(httpTriggerEndPoint: string): Promise<ILogHistoryItem[]> {
        try {
            const [response, draftFields] = await Promise.all([
                this._dmsClient.get(httpTriggerEndPoint, AadHttpClient.configurations.v1),
               // this._sp.web.lists.getByTitle(ListUrls.Draft).fields.select('InternalName', 'Title', 'TypeAsString')()
                this._sp.web.lists.getByTitle(ListUrls.Draft).fields.select('InternalName', 'Title', 'TypeAsString','Hidden',"ReadOnlyField")()
            ]);



            if (!response.ok) {
                console.error("Function App error:", response.statusText);
                return [];
            }

            const data: ILogHistoryItem[] = await response.json();

            const fieldLookup = new Map(draftFields.map(f => [f.InternalName, f]));
            const cleanMembershipValue = (value: string) =>
                value?.startsWith("i:0#.f|membership|") ? value.replace("i:0#.f|membership|", "") : value;

            const processUserField = (value: any) =>
                Array.isArray(value) && value.length > 0 ? value[0].title : '';

            const processUserMultiField = (value: any) =>
                Array.isArray(value) ? value.map(v => v.title).join(", ") : '';

            const processField = (field: any, draftField: any) => {
                const baseResult = { ...field, fieldTitle: draftField.Title, fieldType: draftField.TypeAsString , isHidden:draftField.Hidden, isReadOnly:draftField.ReadOnlyField};


            // const processField = (field: any, draftField: any) => {
            //     const baseResult = { ...field, fieldTitle: draftField.Title, fieldType: draftField.TypeAsString };

                switch (draftField.TypeAsString) {
                    case "TaxonomyFieldType":
                        return {
                            ...baseResult,
                            newValue: field.newValue?.Label,
                            oldValue: field.oldValue?.Label
                        };
                    case "Text":
                        return (field.newValue?.includes("i:0#.f|membership|") || field.oldValue?.includes("i:0#.f|membership|"))
                            ? { ...baseResult, newValue: cleanMembershipValue(field.newValue), oldValue: cleanMembershipValue(field.oldValue) }
                            : baseResult;
                    case "User":
                        return { ...baseResult, newValue: processUserField(field.newValue), oldValue: processUserField(field.oldValue) };
                    case "UserMulti":
                        return { ...baseResult, newValue: processUserMultiField(field.newValue), oldValue: processUserMultiField(field.oldValue) };
                    default:
                        return baseResult;
                }
            };

            const logHistoryData = data.map(item => {
                if (item.actionType === "create") {
                    const { changedFields, ...rest } = item;
                    return rest;
                }

                const changedFieldsWithTitles = item.changedFields?.map(field => {
                    const draftField = fieldLookup.get(field.fieldName);
                    return draftField ? processField(field, draftField) : field;
                });

                return { ...item, changedFields: changedFieldsWithTitles };
            });
            return logHistoryData;
        } catch (error) {
            console.error("Error processing log history:", error);
            return [];
        }
    }
}