import { app, type HttpRequest, type InvocationContext, type HttpResponseInit } from "@azure/functions";
import { DefaultAzureCredential } from "@azure/identity";
import { createRemoteJWKSet, jwtVerify } from "jose";

const tenantId = requiredSetting("ENTRA_TENANT_ID");
const apiApplicationId = requiredSetting("ENTRA_API_CLIENT_ID");
const allowedClientId = requiredSetting("ENTRA_SPA_CLIENT_ID");
const senderEmail = requiredSetting("NOTIFICATION_SENDER_EMAIL");
const requiredScope = process.env.ENTRA_API_SCOPE_NAME?.trim() || "Notifications.Send";
const jwks = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`),
);
const credential = new DefaultAzureCredential();

type NotificationMessage = {
  recipient: string;
  subject: string;
  body: string;
};

function requiredSetting(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing required application setting: ${name}`);
  return value;
}

function response(status: number, message: string): HttpResponseInit {
  return { status, jsonBody: { error: message } };
}

function isNotificationMessage(value: unknown): value is NotificationMessage {
  if (typeof value !== "object" || value === null) return false;
  const message = value as Record<string, unknown>;
  return typeof message.recipient === "string"
    && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(message.recipient)
    && message.recipient.length <= 254
    && typeof message.subject === "string"
    && message.subject.trim().length > 0
    && message.subject.length <= 250
    && typeof message.body === "string"
    && message.body.trim().length > 0
    && message.body.length <= 10_000;
}

async function authorize(request: HttpRequest): Promise<boolean> {
  const authorization = request.headers.get("authorization");
  if (!authorization?.startsWith("Bearer ")) return false;
  try {
    const { payload } = await jwtVerify(authorization.slice("Bearer ".length), jwks, {
      audience: apiApplicationId,
      issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
    });
    const authorizedClient = payload.azp === allowedClientId || payload.appid === allowedClientId;
    const scopes = typeof payload.scp === "string" ? payload.scp.split(" ") : [];
    return authorizedClient && scopes.includes(requiredScope);
  } catch {
    return false;
  }
}

app.http("sendNotification", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "notifications",
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    if (!await authorize(request)) return response(401, "A valid Onedesk notification authorization token is required.");

    let message: unknown;
    try {
      message = await request.json();
    } catch {
      return response(400, "Request body must be valid JSON.");
    }
    if (!isNotificationMessage(message)) {
      return response(400, "Provide a valid recipient email, a subject up to 250 characters, and a message up to 10,000 characters.");
    }

    try {
      const graphToken = await credential.getToken("https://graph.microsoft.com/.default");
      if (!graphToken) throw new Error("Azure managed identity did not provide a Microsoft Graph access token.");
      const graphResponse = await fetch(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}/sendMail`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${graphToken.token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            message: {
              subject: message.subject,
              body: { contentType: "Text", content: message.body },
              toRecipients: [{ emailAddress: { address: message.recipient } }],
            },
            saveToSentItems: true,
          }),
        },
      );
      if (!graphResponse.ok) {
        const details = await graphResponse.text();
        context.error(`Microsoft Graph sendMail failed (${graphResponse.status}): ${details}`);
        return response(502, `Microsoft Graph rejected the email (${graphResponse.status}). Check the Function's Graph permission and sender mailbox settings.`);
      }
      return { status: 202 };
    } catch (error) {
      context.error("Email notification delivery failed.", error);
      return response(502, "Email notification delivery failed. Check the Azure Function logs for details.");
    }
  },
});
