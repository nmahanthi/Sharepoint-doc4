import { SPFI } from "@pnp/sp";
import { FieldNames } from "../../../constants";
import { revisionFieldValidator } from "./RevisionFieldValidator";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { referenceFieldValidator } from "./ReferenceFieldValidator";

export interface IFieldPreSaveAction {
    label: string;
    execute: (context:FormCustomizerContext, value: unknown, sp: SPFI) => Promise<unknown>;
    isBlocking?: boolean;
}

export const customValidators = {
    [FieldNames.Revision]: revisionFieldValidator,
    [FieldNames.ProjectReference]: referenceFieldValidator
    
};