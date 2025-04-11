export const AppSettings = {
    ReferenceIdTableName: process.env['REFERENCE_ID_TABLE_NAME'],
    MassImportTableConnectionString: process.env['MASS_IMPORT_TABLE_CONNECTION_STRING'],
    ServiceBusConnectionString: process.env['SERVICE_BUS_CONNECTION_STRING'],
    TenantName: process.env['TENANT_NAME'],
    IsLocalEnvironment: process.env['IS_LOCAL_ENVIRONMENT'] === 'true',
    KeyVaultUrl: process.env['KEY_VAULT_URL'],
    CertificateName: process.env['CERTIFICATE_NAME'],
    TenantId: process.env['TENANT_ID'],
    ClientId: process.env['CLIENT_ID'],
    CertificatePath: process.env['CLIENT_CERTIFICATE_PATH'],
    PropertyBagData:process.env['PROPERTY_BAG_DATA'],
    InitialReference:process.env['INITIAL_REFERENCE']
};