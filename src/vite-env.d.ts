/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_AZURE_TENANT_ID: string;
  readonly VITE_MICROSOFT_CLIENT_ID: string;
  readonly VITE_EMAIL_API_SCOPE: string;
  readonly VITE_EMAIL_API_URL: string;
}
