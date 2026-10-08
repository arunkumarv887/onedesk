import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
} from "@azure/msal-browser";
import type { AppUser } from "./types";

const tenantId = (
  import.meta.env.VITE_AZURE_TENANT_ID
  || import.meta.env.VITE_MICROSOFT_TENANT_ID
)?.trim();
const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID?.trim();
const emailApiScope = import.meta.env.VITE_EMAIL_API_SCOPE?.trim();
const emailApiUrl = import.meta.env.VITE_EMAIL_API_URL?.trim();

export const microsoftConfigured = Boolean(tenantId && clientId);
export const emailNotificationsConfigured = Boolean(emailApiScope && emailApiUrl);

const client = microsoftConfigured
  ? new PublicClientApplication({
    auth: {
      clientId: clientId || "",
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: { cacheLocation: "sessionStorage" },
  })
  : undefined;

const graphScopes = {
  profile: ["User.Read"],
  workbook: ["User.Read", "Files.ReadWrite.All"],
  email: emailApiScope ? [emailApiScope] : [],
};

let initialization: Promise<void> | undefined;
let interactiveQueue: Promise<void> = Promise.resolve();

function queueInteractive<T>(operation: () => Promise<T>): Promise<T> {
  const request = interactiveQueue.then(operation, operation);
  interactiveQueue = request.then(() => undefined, () => undefined);
  return request;
}

async function ready(): Promise<PublicClientApplication> {
  if (!client) {
    throw new Error("Configure VITE_AZURE_TENANT_ID and VITE_MICROSOFT_CLIENT_ID to enable Microsoft sign-in.");
  }
  initialization ??= client.initialize().then(async () => {
    let redirectAccount: AccountInfo | null = null;
    try {
      redirectAccount = (await client.handleRedirectPromise())?.account || null;
    } catch (error) {
      if (!isMissingTokenRequestCache(error) || hasMicrosoftRedirectResponse()) {
        throw error;
      }
    }
    if (redirectAccount) {
      client.setActiveAccount(redirectAccount);
    } else if (!client.getActiveAccount()) {
      const account = client.getAllAccounts()[0];
      if (account) client.setActiveAccount(account);
    }
  });
  await initialization;
  return client;
}

function hasMicrosoftRedirectResponse(): boolean {
  const response = new URLSearchParams(window.location.hash.slice(1));
  return ["code", "error", "state", "access_token", "id_token"]
    .some((parameter) => response.has(parameter));
}

function activeAccount(msal: PublicClientApplication): AccountInfo | null {
  return msal.getActiveAccount() || msal.getAllAccounts()[0] || null;
}

function stringClaim(account: AccountInfo, name: string): string | undefined {
  const claims = account.idTokenClaims;
  if (!claims || !(name in claims)) return undefined;
  const value = (claims as Record<string, unknown>)[name];
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function userFromAccount(account: AccountInfo): AppUser {
  const email = stringClaim(account, "preferred_username")
    || stringClaim(account, "email")
    || account.username;
  if (!email) throw new Error("Microsoft did not provide an email address for this account.");
  return {
    uid: stringClaim(account, "oid") || account.localAccountId || account.homeAccountId,
    displayName: stringClaim(account, "name") || account.name || email,
    email,
  };
}

async function acquireToken(scopes: string[]): Promise<string> {
  const msal = await ready();
  const account = activeAccount(msal);
  if (!account) throw new Error("Sign in with your Microsoft account first.");
  try {
    const result = await msal.acquireTokenSilent({ account, scopes });
    return result.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError) && !isMissingTokenRequestCache(error)) {
      throw error;
    }
    const result = await queueInteractive(async () => {
      try {
        return await msal.acquireTokenSilent({ account, scopes });
      } catch (retryError) {
        if (!(retryError instanceof InteractionRequiredAuthError) && !isMissingTokenRequestCache(retryError)) {
          throw retryError;
        }
        return msal.acquireTokenPopup({ account, scopes });
      }
    });
    return result.accessToken;
  }
}

function isMissingTokenRequestCache(error: unknown): boolean {
  return typeof error === "object"
    && error !== null
    && "errorCode" in error
    && error.errorCode === "no_token_request_cache_error";
}

export async function getCurrentMicrosoftUser(): Promise<AppUser | null> {
  const msal = await ready();
  const account = activeAccount(msal);
  return account ? userFromAccount(account) : null;
}

export async function signInWithMicrosoft(): Promise<void> {
  const msal = await ready();
  await queueInteractive(() =>
    msal.loginRedirect({ scopes: graphScopes.profile, prompt: "select_account" }),
  );
}
export async function signOutMicrosoft(): Promise<void> {
  const msal = await ready();
  const account = activeAccount(msal);
  if (account) {
    await queueInteractive(() => msal.logoutPopup({ account, mainWindowRedirectUri: window.location.origin }));
  }
}

export async function getMicrosoftAccessToken(
  purpose: keyof typeof graphScopes,
): Promise<string> {
  return acquireToken(graphScopes[purpose]);
}

export async function getMicrosoftAccessTokenForWorkbook(): Promise<string> {
  return acquireToken(graphScopes.workbook);
}

export async function sendMicrosoftMail(
  recipient: string,
  subject: string,
  body: string,
): Promise<void> {
  if (!emailApiScope || !emailApiUrl) {
    throw new Error("Configure VITE_EMAIL_API_SCOPE and VITE_EMAIL_API_URL to enable email notifications.");
  }
  const token = await acquireToken(graphScopes.email);
  const response = await fetch(emailApiUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ recipient, subject, body }),
  });
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`Email notification service could not send the message (${response.status}): ${details}`);
  }
}
