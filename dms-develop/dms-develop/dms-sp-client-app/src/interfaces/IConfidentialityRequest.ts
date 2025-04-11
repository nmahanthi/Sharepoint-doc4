export interface IConfidentialityRequest {
    confidentiality: 'Public' | 'Restricted' | 'Confidential';
    users?:string[]
    siteUrl:string;
    itemId:number;
    libraryId:string;
}