import { logger } from "./service/Logger";

export const AppSettings = {
    get TenantName() {
        if (!process.env['TENANT_NAME']) {
            throw new Error('TENANT_NAME is not set');
        }
        return process.env['TENANT_NAME'];
    },
    get IsLocalEnvironment() {
        if (process.env['IS_LOCAL_ENVIRONMENT'] === 'true')
            logger.trackTrace({
                message: `IS_LOCAL_ENVIRONMENT is set to ${process.env['IS_LOCAL_ENVIRONMENT']}`,
                severity: 'Verbose',
                properties: { source: 'AppSettings' }
            });
        return process.env['IS_LOCAL_ENVIRONMENT'] === 'true';
    },
    get KeyVaultUrl() {
        if (!process.env['KEY_VAULT_URL']) {
            throw new Error('KEY_VAULT_URL is not set');
        }
        return process.env['KEY_VAULT_URL'];
    },
    get CertificateName() {
        if (!process.env['CERTIFICATE_NAME']) {
            throw new Error('CERTIFICATE_NAME is not set');
        }
        return process.env['CERTIFICATE_NAME'];
    },
    get TenantId() {
        if (!process.env['TENANT_ID']) {
            throw new Error('TENANT_ID is not set');
        }
        return process.env['TENANT_ID'];
    },
    get ClientId() {
        if (!process.env['CLIENT_ID']) {
            throw new Error('CLIENT_ID is not set');
        }
        return process.env['CLIENT_ID'];
    },
    get CertificatePath() {
        return process.env['CLIENT_CERTIFICATE_PATH'];
    },
    get CustomPermissionListName() {
        if (!process.env['CUSTOM_PERMISSION_LIST_NAME']) {
            logger.trackTrace({
                message: 'CUSTOM_PERMISSION_LIST_NAME is not set',
                severity: 'Warning',
                properties: { source: 'AppSettings' }
            });
        }
        return process.env['CUSTOM_PERMISSION_LIST_NAME'];
    },
    get AllowedDmsRoles() {
        if (!process.env['ALLOWED_DMS_ROLES']) {
            logger.trackTrace({
                message: 'ALLOWED_DMS_ROLES is not set',
                severity: 'Warning',
                properties: { source: 'AppSettings' }
            });
        }
        return process.env['ALLOWED_DMS_ROLES']?.split(',');
    },
    get PnPLogLevel() {
        return +(process.env['PNP_LOG_LEVEL'] || '1');
    },
    get ImportJobTableName() {
        if (!process.env['IMPORT_JOB_TABLE_NAME']) {
            throw new Error('IMPORT_JOB_TABLE_NAME is not set');
        }
        return process.env['IMPORT_JOB_TABLE_NAME'];
    },
    get MassImportTableConnectionString() {
        if (!process.env['MASS_IMPORT_TABLE_CONNECTION_STRING']) {
            throw new Error('MASS_IMPORT_TABLE_CONNECTION_STRING is not set');
        }
        return process.env['MASS_IMPORT_TABLE_CONNECTION_STRING'];
    },
    get ServiceBusConnectionString() {
        if (!process.env['SERVICE_BUS_CONNECTION_STRING']) {
            throw new Error('SERVICE_BUS_CONNECTION_STRING is not set');
        }
        return process.env['SERVICE_BUS_CONNECTION_STRING'];
    },
    get MassImportTopicName() {
        if (!process.env['MASS_IMPORT_TOPIC_NAME']) {
            throw new Error('MASS_IMPORT_TOPIC_NAME is not set');
        }
        return process.env['MASS_IMPORT_TOPIC_NAME'];
    },
    get ApplicableDocsLibName() {
        return process.env['APPLICABLE_DOCS_LIB_NAME'];
    },
    get DraftDocsLibName() {
        return process.env['DRAFT_DOCS_LIB_NAME'];
    },
    get ImportFieldsMap() {
        if (!process.env['IMPORT_FIELDS_MAP']) {
            throw new Error('IMPORT_FIELDS_MAP is not set');
        }
        try {
            return JSON.parse(process.env['IMPORT_FIELDS_MAP'] || '{}') as { [key: string]: { type: string; name: string, required?: { [importType: string]: boolean }, transform?: { regex: string; replace: string } } };
        } catch (e) {
            logger.trackException({
                exception: new Error(`IMPORT_FIELDS_MAP is not a valid JSON: ${process.env['IMPORT_FIELDS_MAP']}. Details: ${e}`),
                properties: { source: 'AppSettings' }
            });
            return {};
        }
    },
    get ImportDocumentKeyFields() {
        if (!process.env['IMPORT_DOCUMENT_KEY_FIELDS']) {
            throw new Error('IMPORT_DOCUMENT_KEY_FIELDS is not set');
        }
        try {
            return JSON.parse(process.env['IMPORT_DOCUMENT_KEY_FIELDS'] || '{}') as { [key: string]: string[] };
        } catch (e) {
            logger.trackException({
                exception: new Error(`IMPORT_DOCUMENT_KEY_FIELDS is not a valid JSON: ${process.env['IMPORT_DOCUMENT_KEY_FIELDS']}. Details: ${e}`),
                properties: { source: 'AppSettings' }
            });
            return {};
        }
    },
    get DocSetContentTypeId() {
        return process.env['DOC_SET_CONTENT_TYPE_ID'];
    },
    get ImportDocumentAttachmentField() {
        if (!process.env['IMPORT_DOCUMENT_ATTACHMENT_FIELD']) {
            logger.trackTrace({
                message: 'IMPORT_DOCUMENT_ATTACHMENT_FIELD is not set',
                severity: 'Warning',
                properties: { source: 'AppSettings' }
            });
        }
        return process.env['IMPORT_DOCUMENT_ATTACHMENT_FIELD']?.split(';').map(field => field.trim());
    },
    get AutoClassifyTopicName() {
        console.log('AUTOCLASSIFY_TOPIC_NAME HAS BEEN ACCESSED');
        if (!process.env['SERVICE_BUS_TOPIC_NAME_AUTO_CLASSIFY']) {
            throw new Error('SERVICE_BUS_TOPIC_NAME_AUTO_CLASSIFY is not set');
        }
        return process.env['SERVICE_BUS_TOPIC_NAME_AUTO_CLASSIFY'];
    },
    get MessageIncludedFields() {
        return process.env['MESSAGE_INCLUDED_FIELDS']?.split(',');
    },
    get MessagePromotedTaxonomyFields() {
        return process.env['MESSAGE_PROMOTED_TAXONOMY_FIELDS']?.split(',');
    },
    get MessagePromotedFields() {
        return process.env['MESSAGE_PROMOTED_FIELDS']?.split(',');
    },
    get DocStatusExcelColumn() {
        return process.env['DOC_STATUS_EXCEL_COLUMN'];
    },
    get NativeContentType() {
        return process.env['NATIVE_DOCUMENT_CONTENT_TYPE_ID'];
    },
    get NativeContentTypeName() {
        return process.env['NATIVE_DOCUMENT_CONTENT_TYPE_NAME'];
    },
    get ContentTypeTemplateDocumentLibrary() {
        return process.env['CONTENT_TYPE_TEMPLATE_DOCUMENT_LIBRARY'];
    }
};