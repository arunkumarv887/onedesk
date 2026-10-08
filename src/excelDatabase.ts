import { ExcelWorkbookDatabase } from "./excelDb";
import { roleForEmail } from "./types";
import type {
  AppNotification,
  AppUser,
  Asset,
  Member,
  Ticket,
  UserRole,
} from "./types";

export class ExcelDatabase {
  readonly incidents: ExcelWorkbookDatabase;
  readonly users: ExcelWorkbookDatabase;

  private constructor(incidents: ExcelWorkbookDatabase, users: ExcelWorkbookDatabase) {
    if (incidents.identityKey === users.identityKey) {
      throw new Error("The incidents and user details databases must be two different Excel workbooks.");
    }
    this.incidents = incidents;
    this.users = users;
  }

  static async fromFiles(incidentsFile: File, usersFile: File): Promise<ExcelDatabase> {
    const [incidents, users] = await Promise.all([
      ExcelWorkbookDatabase.fromFile(incidentsFile, "incidents"),
      ExcelWorkbookDatabase.fromFile(usersFile, "users"),
    ]);
    return new ExcelDatabase(incidents, users);
  }

  static async fromSharePoint(
    incidentsUrl: string,
    usersUrl: string,
    token: () => Promise<string>,
  ): Promise<ExcelDatabase> {
    if (!incidentsUrl || !usersUrl) {
      throw new Error("Configure both the incidents workbook URL and the user details workbook URL.");
    }
    const [incidents, users] = await Promise.all([
      ExcelWorkbookDatabase.fromSharePoint(incidentsUrl, "incidents", token),
      ExcelWorkbookDatabase.fromSharePoint(usersUrl, "users", token),
    ]);
    return new ExcelDatabase(incidents, users);
  }

  static async create(): Promise<ExcelDatabase> {
    const [incidents, users] = await Promise.all([
      ExcelWorkbookDatabase.create("incidents"),
      ExcelWorkbookDatabase.create("users"),
    ]);
    return new ExcelDatabase(incidents, users);
  }

  get sourceName(): string {
    return `Incidents: ${this.incidents.sourceName} | User details: ${this.users.sourceName}`;
  }

  get schemaChanged(): boolean {
    return this.incidents.schemaChanged || this.users.schemaChanged;
  }

  get isRemote(): boolean {
    return this.incidents.isRemote && this.users.isRemote;
  }

  get storageMode(): string {
    if (this.isRemote) return "SharePoint";
    if (!this.incidents.isRemote && !this.users.isRemote) return "Local files";
    return "Mixed";
  }

  getTickets(): Ticket[] {
    return this.incidents.getTickets();
  }

  getAssets(): Asset[] {
    return this.incidents.getAssets();
  }

  getMembers(): Member[] {
    return this.users.getMembers().map((member) => ({
      ...member,
      role: roleForEmail(member.email),
    }));
  }

  getNotifications(user: AppUser): AppNotification[] {
    return this.incidents.getNotifications(user);
  }

  getRole(user: AppUser): UserRole {
    return this.users.getRole(user);
  }

  getTicketClosureCounts(tickets: Ticket[]): { byAdmin: number; byUser: number } {
    return this.incidents.getTicketClosureCounts(tickets);
  }

  async ensureUser(user: AppUser): Promise<void> {
    const members = this.users.getMembers();
    for (const member of members) {
      const role = roleForEmail(member.email);
      if (member.role === role) continue;
      await this.users.updateMember(member, role);
      await this.incidents.addAudit(
        "Member",
        member.uid,
        "",
        user,
        "role-policy-enforced",
        JSON.stringify({ previousRole: member.role, role }),
      );
    }
    const existing = this.users.getMembers().find(
      (member) => member.email.trim().toLowerCase() === user.email.trim().toLowerCase(),
    );
    const role = roleForEmail(user.email);
    if (existing?.role === role) return;
    const memberId = await this.users.addMember(user.email, user.displayName, role);
    await this.incidents.addAudit(
      "Member",
      memberId,
      "",
      user,
      existing ? "role-policy-enforced" : "user-registered",
      JSON.stringify(existing ? { previousRole: existing.role, role } : { role }),
    );
  }

  async addTicket(ticket: Ticket, actor: AppUser): Promise<void> {
    await this.incidents.addTicket(ticket, actor);
  }

  async updateTicket(ticket: Ticket, changes: Partial<Pick<Ticket, "status" | "priority" | "assigneeName" | "assigneeEmail">>, actor: AppUser): Promise<void> {
    await this.incidents.updateTicket(ticket, changes, actor);
  }

  async addAsset(asset: Asset, actor: AppUser): Promise<void> {
    await this.incidents.addAsset(asset, actor);
  }

  async updateAsset(asset: Asset, status: Asset["status"], actor: AppUser): Promise<void> {
    await this.incidents.updateAsset(asset, status, actor);
  }

  async addMember(email: string, name: string, actor: AppUser): Promise<void> {
    const enforcedRole = roleForEmail(email);
    const memberId = await this.users.addMember(email, name, enforcedRole);
    await this.incidents.addAudit("Member", memberId, "", actor, "role-updated", JSON.stringify({ role: enforcedRole }));
  }

  async addNotification(notification: AppNotification): Promise<void> {
    await this.incidents.addNotification(notification);
  }

  async updateNotification(notification: AppNotification, changes: Partial<AppNotification>): Promise<void> {
    await this.incidents.updateNotification(notification, changes);
  }

  async save(): Promise<void> {
    const changed = [
      ...(this.incidents.schemaChanged ? [{ label: "Incidents", workbook: this.incidents }] : []),
      ...(this.users.schemaChanged ? [{ label: "User details", workbook: this.users }] : []),
    ];
    const results = await Promise.allSettled(changed.map(({ workbook }) => workbook.save()));
    const failures = results.flatMap((result, index) => result.status === "rejected"
      ? [`${changed[index].label} workbook: ${friendlySaveError(result.reason)}`]
      : []);
    if (failures.length) {
      throw new Error(`${failures.join("; ")}. A separate workbook may already have saved successfully; check both files before retrying.`);
    }
  }

  async download(): Promise<void> {
    await Promise.all([this.incidents.download(), this.users.download()]);
  }

  async reload(): Promise<ExcelDatabase> {
    const [incidents, users] = await Promise.all([this.incidents.reload(), this.users.reload()]);
    return new ExcelDatabase(incidents, users);
  }
}

function friendlySaveError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (message.includes("(423)")) {
    return "SharePoint has locked this workbook for editing. Close it in Excel for the web, then retry.";
  }
  return message;
}
