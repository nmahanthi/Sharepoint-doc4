import { SPFI } from "@pnp/sp";
import * as strings from "DmsDocumentFormFormCustomizerStrings";
import { FieldNames } from "../../../constants";
import DmsDocumentForm from "../components/DmsDocumentForm";

export async function referenceFieldValidator(value: unknown, sp: SPFI, key?: string): Promise<string | undefined> {
    const reference = value as string;
    
    if (!this) {
        throw new Error('This method must be bound to a DmsDocumentForm instance');
    }
    const dmsDocumentForm = this as unknown as DmsDocumentForm;
    const { fields } = dmsDocumentForm.state;
   
    const fieldInfo = fields.find(f => f.InternalName === FieldNames.ProjectReference);
    if (!fieldInfo) {
        throw new Error('Reference FieldInfo is required');
    }
    if (fieldInfo.Required && !value) {
        return strings.validationErrorRequiredMissingValue;
    }
    if (!value) {
        return
    }
    if (dmsDocumentForm.props.displayMode===8) {
        const specialCharaRegex = /[^a-zA-Z0-9._-]/g;
        if (specialCharaRegex.test(reference)) {
            return strings.ValidationErrorSpecialCharctersInReference;
        }
    }
    
    return;
}