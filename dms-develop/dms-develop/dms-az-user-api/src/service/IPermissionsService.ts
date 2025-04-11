export enum DmsRole{
    DocumentController = 'Document Controller',
    Contributor = 'Contributor',
    Viewer = 'Viewer',
    ProjectAdmin = 'Project Administrator'
}
export interface IPermissionsService {
    getDmsRoles(url: string): Promise<string[]>;
}