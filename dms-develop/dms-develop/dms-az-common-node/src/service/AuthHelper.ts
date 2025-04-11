import { ClientCertificateCredential, DefaultAzureCredential, ManagedIdentityCredential, TokenCredential } from "@azure/identity";
import { AppSettings } from "../AppSettings";
import { CertificateClient } from "@azure/keyvault-certificates";
import { SecretClient } from "@azure/keyvault-secrets";

export class AuthHelper {
    public async getCredential(): Promise<TokenCredential> {
        let credential: TokenCredential;
        if (AppSettings.IsLocalEnvironment) {
            credential = new ClientCertificateCredential(AppSettings.TenantId, AppSettings.ClientId, AppSettings.CertificatePath)
        } else {
            const kvAccessCredential = new ManagedIdentityCredential();
            const secretClient = new SecretClient(AppSettings.KeyVaultUrl, kvAccessCredential);
            const secretResponse = await secretClient.getSecret(AppSettings.CertificateName);
            credential = new ClientCertificateCredential(AppSettings.TenantId, AppSettings.ClientId, { certificate: secretResponse.value })
        }
        return credential;
    }
}
export const authHelper = new AuthHelper();