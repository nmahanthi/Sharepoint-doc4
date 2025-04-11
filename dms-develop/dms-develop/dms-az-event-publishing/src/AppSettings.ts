export const AppSettings = {
    TenantName: process.env['TENANT_NAME'],
    ChangeTokenTableName: process.env['CHANGE_TOKEN_TABLE_NAME'],
    ChangeTokenStorageConnectionString: process.env['CHANGE_TOKEN_STORAGE_CONNECTION_STRING'],
    ServiceBusTopicName: process.env['SERVICE_BUS_TOPIC_NAME'],
    ServiceBusConnectionString: process.env['SERVICE_BUS_CONNECTION_STRING'],
    MessageIncludedFields: process.env['MESSAGE_INCLUDED_FIELDS']?.split(','),
    MessagePromotedTaxonomyFields: process.env['MESSAGE_PROMOTED_TAXONOMY_FIELDS']?.split(','),
    MessagePromotedFields: process.env['MESSAGE_PROMOTED_FIELDS']?.split(','),
    IsLocalEnvironment: process.env['IS_LOCAL_ENVIRONMENT'] === 'true',
    KeyVaultUrl: process.env['KEY_VAULT_URL'],
    CertificateName: process.env['CERTIFICATE_NAME'],
    TenantId: process.env['TENANT_ID'],
    ClientId: process.env['CLIENT_ID'],
    CertificatePath: process.env['CLIENT_CERTIFICATE_PATH'],
    ServiceBusDmlTopicName: process.env['SERVICE_BUS_DML_TOPIC_NAME']
};