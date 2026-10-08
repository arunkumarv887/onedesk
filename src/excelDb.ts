import type ExcelJS from "exceljs";
import { roleForEmail } from "./types";
import type {
  AppNotification,
  AppUser,
  Asset,
  Member,
  Ticket,
  UserRole,
} from "./types";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
type WorkbookKind = "incidents" | "users";

export const WORKBOOK_SCHEMA = {
  Tickets: [
    "TicketID", "TicketNo", "Subject", "Description", "Kind", "Category", "Status",
    "Priority", "RequesterID", "RequesterName", "RequesterEmail", "AssigneeName",
    "AssigneeEmail", "CreatedAt", "UpdatedAt",
  ],
  Assets: [
    "AssetID", "AssetTag", "Name", "Category", "AssignedTo", "Location", "Status", "UpdatedAt",
  ],
  Members: ["UserID", "Email", "Name", "Role", "CreatedAt", "UpdatedAt"],
  Notifications: [
    "NotificationID", "RecipientID", "Title", "Message", "TicketID", "CreatedAt",
    "Read", "EmailStatus", "EmailError",
  ],
  AuditLog: [
    "AuditID", "RecordType", "RecordID", "TicketNo", "ActorID", "ActorName",
    "Action", "Changes", "CreatedAt",
  ],
  Lookups: ["LookupType", "Value", "SortOrder", "Active"],
} as const;

type CollectionName = keyof typeof WORKBOOK_SCHEMA;
type Row = Record<string, unknown>;
type ImportedTable = ExcelJS.Table & {
  worksheet?: ExcelJS.Worksheet;
  table?: { ref?: string; rows?: ExcelJS.CellValue[][] };
};
type FileHandle = {
  getFile(): Promise<File>;
  queryPermission?(options: { mode: "readwrite" }): Promise<"granted" | "denied" | "prompt">;
  requestPermission?(options: { mode: "readwrite" }): Promise<"granted" | "denied" | "prompt">;
  createWritable?(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>;
};

interface RemoteFile {
  driveId: string;
  itemId: string;
  etag: string;
  url: string;
  token: () => Promise<string>;
}

function dateValue(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Math.round((value - 25569) * 86400000));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  if (typeof value === "string" && value.trim()) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  return null;
}

function cellValue(value: ExcelJS.CellValue): unknown {
  if (value instanceof Date || typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value == null) {
    return value;
  }
  if ("result" in value && value.result !== undefined) return value.result;
  if ("text" in value && typeof value.text === "string") return value.text;
  if ("richText" in value) return value.richText.map((part) => part.text).join("");
  if ("hyperlink" in value && typeof value.hyperlink === "string") return value.hyperlink;
  return String(value);
}

function stringValue(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

function boolValue(value: unknown): boolean {
  return value === true || value === 1 || String(value).toLowerCase() === "true";
}

function headerValues(row: ExcelJS.Row): ExcelJS.CellValue[] {
  return Array.from({ length: row.cellCount }, (_, index) => row.getCell(index + 1).value);
}

function makeId(prefix: string): string {
  return `${prefix}-${crypto.randomUUID()}`;
}

function shareId(url: string): string {
  const bytes = new TextEncoder().encode(url);
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return `u!${btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "")}`;
}

async function graphJson<T>(url: string, token: string): Promise<T> {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    throw new Error(`The SharePoint workbook could not be read (${response.status}): ${await response.text()}`);
  }
  return response.json() as Promise<T>;
}

function downloadWorkbook(buffer: ExcelJS.Buffer, kind: WorkbookKind): void {
  const bytes = new Uint8Array(buffer);
  const blob = new Blob([bytes], { type: XLSX_MIME });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `onedesk-${kind}-${new Date().toISOString().slice(0, 10)}.xlsx`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export class ExcelWorkbookDatabase {
  readonly workbook: ExcelJS.Workbook;
  readonly sourceName: string;
  schemaChanged = false;
  private readonly kind: WorkbookKind;
  private readonly fileHandle?: FileHandle;
  private readonly originalFile?: File;
  private readonly remoteFile?: RemoteFile;
  private saveQueue: Promise<void> = Promise.resolve();

  private constructor(workbook: ExcelJS.Workbook, kind: WorkbookKind, sourceName: string, fileHandle?: FileHandle, remoteFile?: RemoteFile, originalFile?: File) {
    this.workbook = workbook;
    this.kind = kind;
    this.sourceName = sourceName;
    this.fileHandle = fileHandle;
    this.remoteFile = remoteFile;
    this.originalFile = originalFile;
    this.ensureSchema();
  }

  static async fromFile(file: File, kind: WorkbookKind, fileHandle?: FileHandle): Promise<ExcelWorkbookDatabase> {
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      throw new Error("Select an .xlsx Excel workbook.");
    }
    const { default: ExcelJSRuntime } = await import("exceljs");
    const workbook = new ExcelJSRuntime.Workbook();
    try {
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch (error) {
      throw new Error(`The selected Excel workbook could not be opened: ${error instanceof Error ? error.message : String(error)}`);
    }
    return new ExcelWorkbookDatabase(workbook, kind, file.name, fileHandle, undefined, file);
  }

  static async fromSharePoint(url: string, kind: WorkbookKind, token: () => Promise<string>): Promise<ExcelWorkbookDatabase> {
    const initialToken = await token();
    const shareKey = shareId(url);
    const item = await graphJson<{ id: string; name: string; eTag?: string; parentReference?: { driveId?: string } }>(
      `https://graph.microsoft.com/v1.0/shares/${encodeURIComponent(shareKey)}/driveItem`,
      initialToken,
    );
    const driveId = item.parentReference?.driveId;
    if (!driveId || !item.id) {
      throw new Error("Microsoft Graph did not return the SharePoint workbook location.");
    }
    if (!item.name.toLowerCase().endsWith(".xlsx")) {
      throw new Error("The selected SharePoint file is not an .xlsx workbook.");
    }
    const fileToken = await token();
    const content = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}/content`,
      { headers: { Authorization: `Bearer ${fileToken}` } },
    );
    if (!content.ok) {
      throw new Error(`The SharePoint workbook could not be downloaded (${content.status}): ${await content.text()}`);
    }
    const { default: ExcelJSRuntime } = await import("exceljs");
    const workbook = new ExcelJSRuntime.Workbook();
    try {
      await workbook.xlsx.load(await content.arrayBuffer());
    } catch (error) {
      throw new Error(`The SharePoint file is not a readable .xlsx workbook: ${error instanceof Error ? error.message : String(error)}`);
    }
    return new ExcelWorkbookDatabase(workbook, kind, item.name, undefined, {
      driveId,
      itemId: item.id,
      etag: item.eTag || "",
      url,
      token,
    });
  }

  static async create(kind: WorkbookKind): Promise<ExcelWorkbookDatabase> {
    const { default: ExcelJSRuntime } = await import("exceljs");
    const workbook = new ExcelJSRuntime.Workbook();
    workbook.creator = "Onedesk - Topin Technologies";
    workbook.created = new Date();
    workbook.modified = new Date();
    return new ExcelWorkbookDatabase(workbook, kind, `New ITSM ${kind} workbook`);
  }

  private ensureSchema(): void {
    const misplacedSheets = this.kind === "incidents"
      ? ["Members"]
      : ["Tickets", "Assets", "Notifications", "AuditLog", "Lookups"];
    const misplaced = misplacedSheets.filter((name) => this.workbook.getWorksheet(name));
    if (misplaced.length) {
      throw new Error(`The ${this.kind} workbook contains tables for the other database (${misplaced.join(", ")}). Select two separate workbooks so user data and incident data stay split.`);
    }
    const collections = Object.entries(WORKBOOK_SCHEMA).filter(([name]) =>
      this.kind === "incidents" ? name !== "Members" : name === "Members",
    );
    for (const [name, columns] of collections) {
      const worksheet = this.workbook.getWorksheet(name) || this.workbook.addWorksheet(name);
      const table = worksheet.getTable(name);
      if (table) {
        const tableHeaders = headerValues(worksheet.getRow(1)).map((value) => stringValue(value));
        if (tableHeaders.join("|") !== columns.join("|")) {
          throw new Error(`Table "${name}" has incompatible columns. Expected: ${columns.join(", ")}. Found: ${tableHeaders.join(", ")}.`);
        }
        if (!(table as ImportedTable).worksheet) {
          const existingRows = (worksheet.getRows(2, Math.max(worksheet.rowCount - 1, 0)) || [])
            .map((row) => columns.map((_, index) => cellValue(row.getCell(index + 1).value) as ExcelJS.CellValue))
            .filter((row) => row.some((value) => value != null && value !== ""));
          worksheet.removeTable(name);
          worksheet.addTable({
            name,
            ref: "A1",
            headerRow: true,
            totalsRow: false,
            style: { theme: "TableStyleMedium4", showRowStripes: true },
            columns: columns.map((column) => ({ name: column, filterButton: true })),
            rows: existingRows,
          });
        }
        continue;
      }
      this.schemaChanged = true;
      const existingHeaders = headerValues(worksheet.getRow(1)).map((value) => stringValue(value));
      if (existingHeaders.some(Boolean)) {
        const missingHeaders = columns.filter((column) => !existingHeaders.includes(column));
        if (missingHeaders.length) {
          throw new Error(`Sheet "${name}" has incompatible columns. Expected: ${columns.join(", ")}. Found: ${existingHeaders.join(", ")}.`);
        }
      } else {
        worksheet.getRow(1).values = [...columns];
      }
      const actualHeaders = headerValues(worksheet.getRow(1)).map((value) => stringValue(value));
      const existingData = worksheet.rowCount > 1
        ? worksheet.getRows(2, worksheet.rowCount - 1)?.map((row) => actualHeaders.map((_, index) => row.getCell(index + 1).value)) || []
        : [];
      const normalizedHeaders = columns.map((column) => ({ name: column, filterButton: true }));
      const normalizedRows = existingData.map((row) => columns.map((column) => {
        const index = actualHeaders.indexOf(column);
        return index >= 0 ? cellValue(row[index] as ExcelJS.CellValue) : null;
      }));
      if (actualHeaders.join("|") !== columns.join("|")) {
        worksheet.spliceRows(1, worksheet.rowCount);
        worksheet.getRow(1).values = [...columns];
        worksheet.addRows(normalizedRows);
      }
      worksheet.addTable({
        name,
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        style: { theme: "TableStyleMedium4", showRowStripes: true },
        columns: normalizedHeaders,
        rows: normalizedRows,
      });
      if (name !== "Lookups") {
        worksheet.views = [{ state: "frozen", ySplit: 1 }];
        worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
      }
    }

    if (this.kind === "incidents") this.seedLookups();
  }

  private seedLookups(): void {
    const table = this.workbook.getWorksheet("Lookups")?.getTable("Lookups");
    if (!table) return;
    const values = [
      ["TicketStatus", "New"], ["TicketStatus", "In Progress"], ["TicketStatus", "Pending"],
      ["TicketStatus", "Under Maintenance"], ["TicketStatus", "Resolved"], ["TicketStatus", "Closed"],
      ["Priority", "Low"], ["Priority", "Medium"], ["Priority", "High"], ["Priority", "Urgent"],
      ["TicketKind", "Incident"], ["TicketKind", "Service request"],
      ["Category", "Hardware"], ["Category", "Software"], ["Category", "Network"], ["Category", "Access"], ["Category", "Email"], ["Category", "Other"],
      ["AssetStatus", "In use"], ["AssetStatus", "In repair"], ["AssetStatus", "Available"], ["AssetStatus", "Retired"],
      ["UserRole", "requester"], ["UserRole", "agent"], ["UserRole", "admin"],
    ];
    const existing = new Set(this.readRows("Lookups").map((row) => `${stringValue(row.LookupType)}:${stringValue(row.Value)}`));
    const missingValues = values.filter(([type, value]) => !existing.has(`${type}:${value}`));
    if (!missingValues.length) return;
    this.schemaChanged = true;
    missingValues.forEach(([type, value], index) => table.addRow([type, value, existing.size + index + 1, true]));
    table.commit();
  }

  private readRows(name: CollectionName): Row[] {
    const worksheet = this.workbook.getWorksheet(name);
    if (!worksheet || worksheet.rowCount < 2) return [];
    const headers = WORKBOOK_SCHEMA[name];
    const result: Row[] = [];
    const rows = worksheet.getRows(2, worksheet.rowCount - 1) || [];
    for (const row of rows) {
      const values = headers.map((_, index) => cellValue(row.getCell(index + 1).value));
      if (values.every((value) => value == null || value === "")) continue;
      result.push(Object.fromEntries(headers.map((header, index) => [header, values[index]])));
    }
    return result;
  }

  getTickets(): Ticket[] {
    return this.readRows("Tickets").map((row) => ({
      id: stringValue(row.TicketID),
      ticketNo: stringValue(row.TicketNo),
      subject: stringValue(row.Subject),
      description: stringValue(row.Description),
      kind: stringValue(row.Kind) as Ticket["kind"],
      category: stringValue(row.Category),
      status: stringValue(row.Status) as Ticket["status"],
      priority: stringValue(row.Priority) as Ticket["priority"],
      requesterUid: stringValue(row.RequesterID),
      requesterName: stringValue(row.RequesterName),
      requesterEmail: stringValue(row.RequesterEmail),
      assigneeName: stringValue(row.AssigneeName),
      assigneeEmail: stringValue(row.AssigneeEmail),
      createdAt: dateValue(row.CreatedAt),
      updatedAt: dateValue(row.UpdatedAt),
    })).sort((a, b) => (b.createdAt?.getTime() || 0) - (a.createdAt?.getTime() || 0));
  }

  getAssets(): Asset[] {
    return this.readRows("Assets").map((row) => ({
      id: stringValue(row.AssetID),
      name: stringValue(row.Name),
      assetTag: stringValue(row.AssetTag),
      category: stringValue(row.Category),
      assignedTo: stringValue(row.AssignedTo),
      location: stringValue(row.Location),
      status: stringValue(row.Status) as Asset["status"],
      updatedAt: dateValue(row.UpdatedAt),
    })).sort((a, b) => (b.updatedAt?.getTime() || 0) - (a.updatedAt?.getTime() || 0));
  }

  getMembers(): Member[] {
    return this.readRows("Members").map((row) => ({
      uid: stringValue(row.UserID),
      email: stringValue(row.Email),
      name: stringValue(row.Name),
      role: stringValue(row.Role) as UserRole,
      createdAt: dateValue(row.CreatedAt),
      updatedAt: dateValue(row.UpdatedAt),
    }));
  }

  getNotifications(user: AppUser): AppNotification[] {
    return this.readRows("Notifications")
      .filter((row) => stringValue(row.RecipientID).toLowerCase() === user.email.toLowerCase())
      .map((row) => ({
        id: stringValue(row.NotificationID),
        recipientUid: stringValue(row.RecipientID),
        title: stringValue(row.Title),
        message: stringValue(row.Message),
        ticketId: stringValue(row.TicketID),
        createdAt: dateValue(row.CreatedAt),
        read: boolValue(row.Read),
        emailStatus: stringValue(row.EmailStatus),
        emailError: stringValue(row.EmailError),
      })).sort((a, b) => (b.createdAt?.getTime() || 0) - (a.createdAt?.getTime() || 0));
  }

  getRole(user: AppUser): UserRole {
    return roleForEmail(user.email);
  }

  getTicketClosureCounts(tickets: Ticket[]): { byAdmin: number; byUser: number } {
    const closedTickets = new Map(tickets.filter((ticket) => ticket.status === "Closed").map((ticket) => [ticket.id, ticket]));
    const latestClosure = new Map<string, { actorId: string; createdAt: number; byUser: boolean }>();
    for (const row of this.readRows("AuditLog")) {
      const ticket = closedTickets.get(stringValue(row.RecordID));
      if (stringValue(row.RecordType) !== "Ticket" || !ticket) continue;
      const action = stringValue(row.Action);
      const changes = stringValue(row.Changes);
      if (action !== "closed-by-admin" && action !== "closed-by-user" && !/"status"\s*:\s*"Closed"/.test(changes)) continue;
      const createdAt = dateValue(row.CreatedAt)?.getTime() || 0;
      const previous = latestClosure.get(ticket.id);
      if (previous && previous.createdAt > createdAt) continue;
      const actorId = stringValue(row.ActorID);
      latestClosure.set(ticket.id, {
        actorId,
        createdAt,
        byUser: action === "closed-by-user" || (action !== "closed-by-admin" && actorId === ticket.requesterUid),
      });
    }
    let byAdmin = 0;
    let byUser = 0;
    for (const closure of latestClosure.values()) {
      if (closure.byUser) byUser += 1;
      else byAdmin += 1;
    }
    return { byAdmin, byUser };
  }

  get isRemote(): boolean {
    return Boolean(this.remoteFile);
  }

  get identityKey(): string {
    if (this.remoteFile) return `sharepoint:${this.remoteFile.driveId}:${this.remoteFile.itemId}`;
    if (this.originalFile) return `local:${this.originalFile.name}:${this.originalFile.size}:${this.originalFile.lastModified}`;
    return `new:${this.kind}`;
  }

  async reload(): Promise<ExcelWorkbookDatabase> {
    if (this.remoteFile) {
      return ExcelWorkbookDatabase.fromSharePoint(this.remoteFile.url, this.kind, this.remoteFile.token);
    }
    const file = this.fileHandle ? await this.fileHandle.getFile() : this.originalFile;
    if (!file) throw new Error("The original workbook file is unavailable. Select the workbook again.");
    return ExcelWorkbookDatabase.fromFile(file, this.kind, this.fileHandle);
  }

  async addTicket(ticket: Ticket, actor: AppUser): Promise<void> {
    await this.addRow("Tickets", {
      TicketID: ticket.id, TicketNo: ticket.ticketNo, Subject: ticket.subject,
      Description: ticket.description, Kind: ticket.kind, Category: ticket.category,
      Status: ticket.status, Priority: ticket.priority, RequesterID: ticket.requesterUid,
      RequesterName: ticket.requesterName, RequesterEmail: ticket.requesterEmail,
      AssigneeName: ticket.assigneeName, AssigneeEmail: ticket.assigneeEmail,
      CreatedAt: ticket.createdAt, UpdatedAt: ticket.updatedAt,
    });
    await this.addAudit("Ticket", ticket.id, ticket.ticketNo, actor, "created", "");
    this.schemaChanged = true;
  }

  async updateTicket(ticket: Ticket, changes: Partial<Pick<Ticket, "status" | "priority" | "assigneeName" | "assigneeEmail">>, actor: AppUser): Promise<void> {
    const values: Row = { ...changes, updatedAt: new Date() };
    const columnNames: Record<string, string> = {
      status: "Status", priority: "Priority", assigneeName: "AssigneeName",
      assigneeEmail: "AssigneeEmail", updatedAt: "UpdatedAt",
    };
    const translated = Object.fromEntries(Object.entries(values).map(([key, value]) => [columnNames[key] || key, value]));
    await this.updateRow("Tickets", "TicketID", ticket.id, translated);
    const action = changes.status === "Closed"
      ? actor.uid === ticket.requesterUid ? "closed-by-user" : "closed-by-admin"
      : "updated";
    await this.addAudit("Ticket", ticket.id, ticket.ticketNo, actor, action, JSON.stringify(changes));
    this.schemaChanged = true;
  }

  async addAsset(asset: Asset, actor: AppUser): Promise<void> {
    await this.addRow("Assets", {
      AssetID: asset.id, AssetTag: asset.assetTag, Name: asset.name,
      Category: asset.category, AssignedTo: asset.assignedTo, Location: asset.location,
      Status: asset.status, UpdatedAt: asset.updatedAt,
    });
    await this.addAudit("Asset", asset.id, asset.assetTag, actor, "created", "");
    this.schemaChanged = true;
  }

  async updateAsset(asset: Asset, status: Asset["status"], actor: AppUser): Promise<void> {
    await this.updateRow("Assets", "AssetID", asset.id, { Status: status, UpdatedAt: new Date() });
    await this.addAudit("Asset", asset.id, asset.assetTag, actor, "status-updated", JSON.stringify({ status }));
    this.schemaChanged = true;
  }

  async addMember(email: string, name: string, role: UserRole): Promise<string> {
    const existing = this.getMembers().find((member) => member.email.toLowerCase() === email.toLowerCase());
    const now = new Date();
    const memberId = existing?.uid || makeId("USER");
    await this.upsertRow("Members", "Email", email, {
      UserID: memberId, Email: email, Name: name, Role: role,
      CreatedAt: existing?.createdAt || now, UpdatedAt: now,
    });
    this.schemaChanged = true;
    return memberId;
  }

  async updateMember(member: Member, role: UserRole): Promise<void> {
    await this.updateRow("Members", "UserID", member.uid, { Role: role, UpdatedAt: new Date() });
    this.schemaChanged = true;
  }

  async addNotification(notification: AppNotification): Promise<void> {
    await this.addRow("Notifications", {
      NotificationID: notification.id, RecipientID: notification.recipientUid,
      Title: notification.title, Message: notification.message, TicketID: notification.ticketId,
      CreatedAt: notification.createdAt, Read: notification.read,
      EmailStatus: notification.emailStatus, EmailError: notification.emailError,
    });
    this.schemaChanged = true;
  }

  async updateNotification(notification: AppNotification, changes: Partial<AppNotification>): Promise<void> {
    const columns: Record<string, string> = {
      read: "Read", emailStatus: "EmailStatus", emailError: "EmailError",
    };
    const translated = Object.fromEntries(
      Object.entries(changes).map(([key, value]) => [columns[key] || key, value]),
    );
    await this.updateRow("Notifications", "NotificationID", notification.id, translated);
    this.schemaChanged = true;
  }

  async save(): Promise<void> {
    this.saveQueue = this.saveQueue.catch(() => undefined).then(() => this.writeWorkbook());
    return this.saveQueue;
  }

  async download(): Promise<void> {
    const buffer = await this.workbook.xlsx.writeBuffer();
    downloadWorkbook(buffer, this.kind);
  }

  private async writeWorkbook(): Promise<void> {
    this.workbook.modified = new Date();
    const buffer = await this.workbook.xlsx.writeBuffer();
    if (this.remoteFile) {
      const token = await this.remoteFile.token();
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(this.remoteFile.driveId)}/items/${encodeURIComponent(this.remoteFile.itemId)}/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": XLSX_MIME,
            ...(this.remoteFile.etag ? { "If-Match": this.remoteFile.etag } : {}),
          },
          body: new Blob([new Uint8Array(buffer)], { type: XLSX_MIME }),
        },
      );
      if (response.status === 412) {
        throw new Error("This workbook changed in SharePoint since it was opened. Reload the workbook before saving to avoid overwriting someone else's changes.");
      }
      if (!response.ok) {
        throw new Error(`Changes could not be saved to SharePoint (${response.status}): ${await response.text()}`);
      }
      const updated = await response.json() as { eTag?: string };
      this.remoteFile.etag = updated.eTag || this.remoteFile.etag;
      this.schemaChanged = false;
      return;
    }

    if (this.fileHandle?.createWritable) {
      const permission = await this.fileHandle.queryPermission?.({ mode: "readwrite" });
      const granted = permission === "granted"
        || (await this.fileHandle.requestPermission?.({ mode: "readwrite" })) === "granted";
      if (!granted) {
        throw new Error("Write access to the selected workbook was not granted. Reopen it with edit access or use Download workbook.");
      }
      const writable = await this.fileHandle.createWritable();
      try {
        await writable.write(new Blob([new Uint8Array(buffer)], { type: XLSX_MIME }));
        await writable.close();
      } catch (error) {
        throw new Error(`The Excel workbook could not be saved: ${error instanceof Error ? error.message : String(error)}`);
      }
      this.schemaChanged = false;
      return;
    }
    downloadWorkbook(buffer, this.kind);
    this.schemaChanged = false;
  }

  async addAudit(recordType: string, recordId: string, ticketNo: string, actor: AppUser, action: string, changes: string): Promise<void> {
    await this.addRow("AuditLog", {
      AuditID: makeId("AUD"), RecordType: recordType, RecordID: recordId,
      TicketNo: ticketNo, ActorID: actor.uid, ActorName: actor.displayName,
      Action: action, Changes: changes, CreatedAt: new Date(),
    });
    this.schemaChanged = true;
  }

  private async addRow(name: CollectionName, row: Row): Promise<void> {
    const worksheet = this.workbook.getWorksheet(name);
    const table = worksheet?.getTable(name);
    if (!table) throw new Error(`The "${name}" table is missing from this workbook.`);
    const values = WORKBOOK_SCHEMA[name].map((column) => row[column] ?? null);
    table.addRow(values as ExcelJS.CellValue[]);
    table.commit();
  }

  private async updateRow(name: CollectionName, key: string, id: string, changes: Row): Promise<void> {
    const worksheet = this.workbook.getWorksheet(name);
    const headers = WORKBOOK_SCHEMA[name] as readonly string[];
    const keyIndex = headers.indexOf(key);
    if (!worksheet || keyIndex < 0) throw new Error(`The "${name}" table is missing its ${key} column.`);
    const rows = worksheet.getRows(2, Math.max(worksheet.rowCount - 1, 0)) || [];
    const row = rows.find((item) => stringValue(cellValue(item.getCell(keyIndex + 1).value)) === id);
    if (!row) throw new Error(`The ${name} record "${id}" no longer exists in this workbook.`);
    for (const [column, value] of Object.entries(changes)) {
      const index = headers.indexOf(column);
      if (index >= 0) row.getCell(index + 1).value = value as ExcelJS.CellValue;
    }
    worksheet.getTable(name)?.commit();
  }

  private async upsertRow(name: CollectionName, key: string, id: string, values: Row): Promise<void> {
    const worksheet = this.workbook.getWorksheet(name);
    const headers = WORKBOOK_SCHEMA[name] as readonly string[];
    const keyIndex = headers.indexOf(key);
    const existingRows = worksheet?.getRows(2, Math.max((worksheet?.rowCount || 1) - 1, 0)) || [];
    const existing = existingRows.find((row) => stringValue(cellValue(row.getCell(keyIndex + 1).value)).toLowerCase() === id.toLowerCase());
    if (existing) {
      for (const [column, value] of Object.entries(values)) {
        const index = headers.indexOf(column);
        if (index >= 0) existing.getCell(index + 1).value = value as ExcelJS.CellValue;
      }
      worksheet?.getTable(name)?.commit();
      return;
    }
    await this.addRow(name, values);
  }
}
