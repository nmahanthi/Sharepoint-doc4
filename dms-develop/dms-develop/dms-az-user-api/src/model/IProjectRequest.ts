export interface IProjectRequest {
    properties?: {
        key: string
        value: string;
        spFieldName?: string;
        spValue?: string;
    }[];
    fields?: {
        fieldName: string;
        hidden?: boolean;
        displayName?: string;
        choices?: string[];
    }[];
    url: string;
}