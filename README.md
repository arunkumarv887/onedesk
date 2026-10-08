# Onedesk

Onedesk is the IT service desk for Topin Technologies.

A responsive IT service desk using Microsoft Entra ID for sign-in, Microsoft Graph for Outlook notifications, and **two separate Excel workbooks** for its database.

## Excel database design

### Incidents workbook

The configured incidents workbook contains ITSM operational data only:

| Sheet / table | Purpose | Columns |
|---|---|---|
| `Tickets` | Incidents and service requests | TicketID, TicketNo, Subject, Description, Kind, Category, Status, Priority, RequesterID, RequesterName, RequesterEmail, AssigneeName, AssigneeEmail, CreatedAt, UpdatedAt |
| `Assets` | Managed hardware/resources | AssetID, AssetTag, Name, Category, AssignedTo, Location, Status, UpdatedAt |
| `Notifications` | In-app alerts and Outlook delivery outcome | NotificationID, RecipientID, Title, Message, TicketID, CreatedAt, Read, EmailStatus, EmailError |
| `AuditLog` | ITSM and member-role change history | AuditID, RecordType, RecordID, TicketNo, ActorID, ActorName, Action, Changes, CreatedAt |
| `Lookups` | Supported statuses, priorities, categories, and roles | LookupType, Value, SortOrder, Active |

Ticket statuses include New (open), In Progress, Pending, Under Maintenance, Resolved, and Closed. Closure events are recorded in `AuditLog` and attributed to the requester or service desk account.

### User details workbook

Keep personal/user details and application roles in a **different** `.xlsx` file. This workbook has one structured table:

| Sheet / table | Purpose | Columns |
|---|---|---|
| `Members` | Microsoft identity-to-ITSM role mapping | UserID, Email, Name, Role, CreatedAt, UpdatedAt |

The account configured in `VITE_ONEDESK_ADMIN_EMAIL` (default `murthy@topin.co.in`) receives the administrator role. Every other account is a requester; the app does not allow other members to be promoted to agent or admin. On connection, Onedesk corrects existing workbook roles to this policy. Workbook-stored roles only control the app interface, not security: SharePoint permissions are the actual access boundary. Restrict both files to authorized users, and enable SharePoint version history/backups.

New users can choose **Register here** on the sign-in page and continue with their work or school Microsoft account after an administrator invites them as a guest (Entra B2B) in the Topin tenant. On first connection to the workbooks, Onedesk creates their member profile from their verified Microsoft identity; no separate Onedesk password is created. Registration does not create the guest invitation or grant access to SharePoint: invite the user in Entra, then grant the guest the required permissions to both workbooks. All accounts except the configured administrator receive the requester role, including when they are the first account in an empty `Members` table.

The app adds missing sheets/tables to selected workbooks, but rejects sheets or tables with incompatible existing columns rather than silently discarding records. You can select both `.xlsx` files at startup or create and download a blank pair of workbooks.

## Configure Microsoft and SharePoint

1. Configure the SPA app registration as a **single-tenant** app for the Topin directory. External users sign in as guests in this resource tenant; do not use the `organizations` authority for this guest-only setup.
2. Add a **Single-page application** redirect URI for each app origin, including exactly `http://localhost:5173` for local development and the deployed HTTPS origin. Do not create or embed a client secret in this browser app.
3. Add Microsoft Graph delegated permissions:
   - `User.Read` for the signed-in user's profile
   - `Files.ReadWrite.All` for both SharePoint workbooks
   The browser app does not need Graph `Mail.Send`; notifications are sent by the server-side Azure Function.
4. Grant any administrator consent required by the tenant. Invite each external user as an Entra B2B guest into Topin, have them redeem the invitation, then grant the guest permission to both workbooks and confirm they can open and edit the files in SharePoint. Signing in to Onedesk by itself does not provision a guest or grant workbook access.
   Sign-in uses a full-page Microsoft redirect (not a popup). On return, the app processes the redirect response and restores the signed-in account.
5. Copy `.env.example` to `.env.local`, set `VITE_AZURE_TENANT_ID` to the **Topin directory (tenant) ID**, and set `VITE_MICROSOFT_CLIENT_ID` to the SPA app registration's **Application (client) ID**. Set `VITE_INCIDENTS_WORKBOOK_URL` and `VITE_USERS_WORKBOOK_URL` to the SharePoint URLs for the incidents and user-details workbooks.
6. Run `npm install` and `npm run dev`, then open `http://localhost:5173` (not `127.0.0.1` or another port). Vite is pinned to port 5173 and will report an error if that port is occupied rather than silently choosing a different redirect URI.

The SPA uses the Topin tenant ID as its sign-in authority. This lets invited external users authenticate in the directory where the SPA and SharePoint workbooks exist; an external user's home-tenant authority can instead produce `AADSTS700016` because the SPA is not registered there. Browser environment variables are public; never put secrets in `VITE_` variables. When both workbook URLs are set, Onedesk automatically connects after sign-in. If either URL is missing, enter the workbook URLs on the connection screen. Microsoft Graph still checks the signed-in account's permissions for both files.

## Configure the server-side email sender

Email is sent from the fixed admin mailbox `murthy@topin.co.in` by the Azure Function under `api/`. The mailbox password or app password is not used by the app and must never be stored in source code, `.env.local`, or browser settings.

1. Create a separate single-tenant Entra app registration for the Notifications API. Under **Expose an API**, set its Application ID URI (for example `api://<API_CLIENT_ID>`) and add the delegated scope `Notifications.Send`. Set `api.requestedAccessTokenVersion` to `2` in this API registration's manifest so the Function can validate v2 tokens.
2. In the SPA app registration, add the Notifications API's `Notifications.Send` delegated permission and grant tenant consent. Keep the SPA's existing `User.Read` and `Files.ReadWrite.All` permissions. Do not give the SPA Graph `Mail.Send`.
3. Set `VITE_EMAIL_API_SCOPE` in `.env.local` to `api://<API_CLIENT_ID>/Notifications.Send` and `VITE_EMAIL_API_URL` to the Function URL, such as `https://<FUNCTION_APP>.azurewebsites.net/api/notifications`. These are identifiers/URLs, not secrets.
4. Deploy `api/` as an Azure Functions Node.js 22 app. Enable its system-assigned managed identity and grant that identity Microsoft Graph **application** permission `Mail.Send` with administrator consent. Restrict the permission to the sender mailbox using Exchange Application RBAC where available.
5. Add these Function application settings: `ENTRA_TENANT_ID`, `ENTRA_API_CLIENT_ID` (Notifications API registration client ID), `ENTRA_SPA_CLIENT_ID` (SPA registration client ID), `ENTRA_API_SCOPE_NAME` (`Notifications.Send`), and `NOTIFICATION_SENDER_EMAIL` (`murthy@topin.co.in`). Do not configure an app password.
6. Configure the Function's CORS allowlist for the exact app origins (including `http://localhost:5173` for local development). The Function independently validates the token signature, tenant, SPA client, audience, and scope; do not remove this authorization check.
7. Build and deploy from the `api/` directory using the Azure Functions Core Tools. Rebuild/redeploy the frontend after setting its two `VITE_EMAIL_*` variables.

If an app password has already been shared, revoke it immediately. This Graph/managed-identity integration does not need that password.

## Saving, refresh, and concurrency

- SharePoint workbooks are saved back through Microsoft Graph. Each file uses its own eTag conflict check; if a workbook changed since it was opened, reload it before retrying. Close the files in Excel for the web before retrying a save if SharePoint reports that a workbook is locked.
- A local `.xlsx` pair is an offline copy. Save downloads two updated files that must be put back in their shared location.
- Excel is a file database, not a transactional multi-user database. Saving two separate files is not atomic; if one save succeeds and the other fails, the app identifies the failed workbook. Avoid editing either file manually while agents are working.

## Notifications

Ticket creation, status/assignment changes, and user-role changes are recorded in the incident workbook's `Notifications`/`AuditLog` tables as applicable. Ticket notification email is sent through Microsoft Graph by the Azure Function from the configured admin mailbox. Email delivery is attempted even if saving the in-app notification fails; delivery and logging errors are reported separately, and a failed email does not discard the ticket change.

## Development

```sh
npm install
npm run dev
npm run build
```
