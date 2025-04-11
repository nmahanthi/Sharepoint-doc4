export interface IUpdateApplicableDocRequest {
    properties?: {
        FieldName?: string;
        FieldValue?: string;
    }[];
    siteUrl: string;
    docSetId: number;
}