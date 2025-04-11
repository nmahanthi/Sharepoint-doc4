import { PermissionKind } from "@pnp/sp/security/index.js";

export const AppSettings = {
    CustomPermissionListName: process.env['PERMISSIONS_LIST_NAME'],
    ProjectUpdateRequiredWebPermision: parseInt(process.env['PROJECT_UPDATE_REQ_WEB_PERM']) as PermissionKind,
    ConfidentialValue: process.env['CONFIDENTIAL_VALUE'],
    ServiceBusTopicName: process.env['SERVICE_BUS_TOPIC_NAME'],
    ServiceBusConnectionString: process.env['SERVICE_BUS_CONNECTION_STRING'],
    MassImportTopicName: process.env['MASS_IMPORT_TOPIC_NAME'],
    NewMassImportSubject: process.env['NEW_MASS_IMPORT_SUBJECT'],
    DocChangesTableConnectionString: process.env['DOCUMENT_CHANGES_TABLE_CONNECTION_STRING'],
    DocumentChangesTableName:process.env['DOCUMENT_CHANGES_TABLE_NAME'],
    IsLocalEnvironment: true,
    TENANT_NAME:process.env['TENANT_NAME'],
   
};