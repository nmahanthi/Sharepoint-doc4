import { SPFI } from "@pnp/sp";
import * as strings from "DmsDocumentFormFormCustomizerStrings";
import { FieldNames } from "../../../constants";
import DmsDocumentForm from "../components/DmsDocumentForm";
import { RenderListDataOptions } from "@pnp/sp/lists";

export async function revisionFieldValidator(value: unknown, sp: SPFI, key?: string): Promise<string | undefined> {
    const revision = value as string;
    if (!this) {
        throw new Error('This method must be bound to a DmsDocumentForm instance');
    }
    const dmsDocumentForm = this as unknown as DmsDocumentForm;
    const { fields, item, libraries } = dmsDocumentForm.state;
    const { context } = dmsDocumentForm.props;
    const fieldInfo = fields.find(f => f.InternalName === FieldNames.Revision);
    if (!fieldInfo) {
        throw new Error('Revision FieldInfo is required');
    }
    if (fieldInfo.Required && !value) {
        return strings.validationErrorRequiredMissingValue;
    }
    if (!value) {
        return;
    }
    if (value) {
        const specialCharaRegex = /[^a-zA-Z0-9.]/g;
        if (specialCharaRegex.test(revision)) {
            return strings.validationErrorSpecialCharacters;
        }
    }

    const projectRef = item ? (item as { [key: string]: unknown })[FieldNames.ProjectReference] as string : undefined;
    if (!projectRef) {
        return strings.validationErrorProjectReferenceMissing;
    }

    const camlQuery = `<View Scope="RecursiveAll">
        <Query>
            <Where>
                <And>
                    <Eq>
                        <FieldRef Name="${FieldNames.ProjectReference}" />
                        <Value Type="Text">${projectRef}</Value>
                    </Eq>
                    <Eq>
                        <FieldRef Name="${FieldNames.Revision}" />
                        <Value Type="Text">${revision}</Value>
                    </Eq>
                </And>
            </Where>
        </Query>
    </View>`;

    let camlQueryCurrentList = camlQuery;
    if(item?.ID){
        camlQueryCurrentList = `<View Scope="RecursiveAll">
            <Query>
                <Where>
                    <And>
                        <And>
                            <Eq>
                                <FieldRef Name="${FieldNames.ProjectReference}" />
                                <Value Type="Text">${projectRef}</Value>
                            </Eq>
                            <Eq>
                                <FieldRef Name="${FieldNames.Revision}" />
                                <Value Type="Text">${revision}</Value>
                            </Eq>
                        </And>
                        <Neq>
                            <FieldRef Name="ID" />
                            <Value Type="Counter">${item.ID}</Value>
                        </Neq>
                    </And>
                </Where>
            </Query>
        </View>`;
    }

    const duplicateItems = await sp.web.lists
        .getById(context.list.guid.toString())
        .renderListDataAsStream({
            ViewXml: camlQueryCurrentList,
            RenderOptions: RenderListDataOptions.ListData,
        });

    if (duplicateItems?.Row?.length > 0) {
        return strings.validationErrorDuplicateRevision;
    }

    let totalCount = 0;
    if (!libraries) {
        throw new Error('Libraries are undefined');
    }
    for (const libraryId of libraries) {
        if (totalCount === 0) {
            const listItems = await sp.web.lists
                .getById(libraryId)
                .renderListDataAsStream({
                    ViewXml: camlQuery,
                    RenderOptions: RenderListDataOptions.ListData,
                });
            const itemCount = listItems?.Row?.length || 0;
            totalCount += itemCount;

            if (totalCount > 0) {
                break;
            }
        }
    }
    if (totalCount > 0) {
        return strings.validationErrorDuplicateRevision;
    }

    return;
}