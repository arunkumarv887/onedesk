import { useEffect, useMemo, useRef, useState, type FormEvent, type ReactNode } from "react";
import {
  Activity,
  ArrowDownRight,
  ArrowUpRight,
  Bell,
  BookOpen,
  BriefcaseBusiness,
  CalendarDays,
  Check,
  CheckCheck,
  ChevronDown,
  ChevronRight,
  CircleHelp,
  Clock3,
  Command,
  Download,
  FileText,
  Filter,
  HardDrive,
  Headset,
  Laptop,
  LayoutDashboard,
  LifeBuoy,
  LogOut,
  Mail,
  Menu,
  MoreHorizontal,
  Plus,
  Search,
  Send,
  Settings,
  ShieldCheck,
  Sparkles,
  Tag,
  TicketCheck,
  TicketPlus,
  UserRound,
  UsersRound,
  X,
  Zap,
} from "lucide-react";
import { ExcelDatabase } from "./excelDatabase";
import {
  getCurrentMicrosoftUser,
  getMicrosoftAccessTokenForWorkbook,
  emailNotificationsConfigured,
  microsoftConfigured,
  sendMicrosoftMail,
  signInWithMicrosoft,
  signOutMicrosoft,
} from "./microsoft";
import type {
  AppNotification,
  AppUser,
  Asset,
  Member,
  Ticket,
  TicketKind,
  TicketPriority,
  TicketStatus,
  UserRole,
} from "./types";

type Page = "Dashboard" | "Incidents" | "Requests" | "Assets" | "Knowledge" | "Reports" | "Members" | "Settings";
type Toast = { kind: "success" | "error" | "info"; message: string };

const incidentsWorkbookUrl = import.meta.env.VITE_INCIDENTS_WORKBOOK_URL?.trim() || "";
const usersWorkbookUrl = import.meta.env.VITE_USERS_WORKBOOK_URL?.trim() || "";
const tenantIdConfigured = Boolean(
  import.meta.env.VITE_AZURE_TENANT_ID?.trim()
  || import.meta.env.VITE_MICROSOFT_TENANT_ID?.trim(),
);
const clientIdConfigured = Boolean(import.meta.env.VITE_MICROSOFT_CLIENT_ID?.trim());

const navigation: { label: Page; icon: typeof LayoutDashboard; group: string }[] = [
  { label: "Dashboard", icon: LayoutDashboard, group: "WORKSPACE" },
  { label: "Incidents", icon: LifeBuoy, group: "WORKSPACE" },
  { label: "Requests", icon: TicketCheck, group: "WORKSPACE" },
  { label: "Assets", icon: HardDrive, group: "MANAGE" },
  { label: "Knowledge", icon: BookOpen, group: "MANAGE" },
  { label: "Reports", icon: Activity, group: "INSIGHTS" },
  { label: "Members", icon: UsersRound, group: "INSIGHTS" },
  { label: "Settings", icon: Settings, group: "INSIGHTS" },
];

const articles = [
  { title: "Getting started with your workstation", category: "Hardware", time: "4 min read", icon: Laptop, summary: "Set up your Topin Technologies workstation and get the essentials ready for your first day.", steps: ["Connect your laptop to power and sign in with your company Microsoft account.", "Join the approved office Wi-Fi network; use the VPN when working remotely.", "Open Company Portal to install approved business applications.", "If your device is missing or damaged, submit an incident and include the asset tag."] },
  { title: "Connect securely to the company VPN", category: "Network", time: "3 min read", icon: ShieldCheck, summary: "Use the company VPN to securely reach internal tools while away from the office.", steps: ["Connect to a trusted internet network.", "Open the VPN application installed by IT and choose the Topin Technologies profile.", "Sign in with your company Microsoft account and complete any verification prompt.", "If connection fails, capture the error message and contact the service desk."] },
  { title: "Requesting access to business applications", category: "Access", time: "2 min read", icon: UsersRound, summary: "Request access through the service desk so your manager and application owner can review it.", steps: ["Create a service request and select the Access category.", "Include the application name, access level, and business reason.", "Name your manager or project owner in the request description.", "IT will follow up if an approval or additional information is required."] },
  { title: "Troubleshooting Outlook sync issues", category: "Email", time: "5 min read", icon: Mail, summary: "Try these quick checks if Outlook is not syncing new mail or calendar changes.", steps: ["Confirm that you are online and Outlook is not in Work Offline mode.", "Restart Outlook and allow a minute for your mailbox to reconnect.", "Check Outlook on the web to see whether the issue is limited to one device.", "If it continues, report an incident with the affected device and approximate start time."] },
  { title: "Keeping your account secure", category: "Security", time: "3 min read", icon: ShieldCheck, summary: "Protect your work account with safe sign-in habits and prompt reporting.", steps: ["Never share your password or approve an unexpected sign-in prompt.", "Use the organization-approved verification methods when signing in.", "Report suspicious email using your mail client's reporting feature.", "Contact IT promptly if you think your password or device may be compromised."] },
  { title: "Connect a printer to the office network", category: "Hardware", time: "4 min read", icon: BriefcaseBusiness, summary: "Get help connecting to a shared office printer.", steps: ["Confirm the printer location and name with your office administrator.", "Connect your workstation to the office network or company VPN.", "Install the approved printer from your managed device settings.", "Submit a service request with the printer name if it is not listed."] },
];

function friendlyError(error: unknown): string {
  if (error instanceof Error) return error.message;
  return "Something went wrong. Please try again.";
}

function firstName(user: AppUser): string {
  return user.displayName?.split(" ")[0] || user.email?.split("@")[0] || "there";
}

function dateLabel(value: Date | null): string {
  if (!value) return "Just now";
  return value.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}

function timeAgo(value: AppNotification["createdAt"]): string {
  if (!value) return "Just now";
  const minutes = Math.max(1, Math.round((Date.now() - value.getTime()) / 60_000));
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  return `${Math.floor(hours / 24)}d ago`;
}

function initials(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() || "")
    .join("");
}

function StatusPill({ status }: { status: TicketStatus | Asset["status"] }) {
  const className = status.toLowerCase().replaceAll(" ", "-");
  return <span className={`status-pill ${className}`}><span />{status}</span>;
}

function PriorityPill({ priority }: { priority: TicketPriority }) {
  return <span className={`priority-pill ${priority.toLowerCase()}`}>{priority}</span>;
}

function App() {
  const [user, setUser] = useState<AppUser | null>(null);
  const [database, setDatabase] = useState<ExcelDatabase | null>(null);
  const [workbookDirty, setWorkbookDirty] = useState(false);
  const [workbookBusy, setWorkbookBusy] = useState(false);
  const [dataVersion, setDataVersion] = useState(0);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState("");
  const [authMode, setAuthMode] = useState<"login" | "register">("login");
  const [page, setPage] = useState<Page>("Dashboard");
  const [tickets, setTickets] = useState<Ticket[]>([]);
  const [assets, setAssets] = useState<Asset[]>([]);
  const [members, setMembers] = useState<Member[]>([]);
  const [role, setRole] = useState<UserRole>("requester");
  const [notifications, setNotifications] = useState<AppNotification[]>([]);
  const [dataError, setDataError] = useState("");
  const [toast, setToast] = useState<Toast | null>(null);
  const [search, setSearch] = useState("");
  const [ticketKindFilter, setTicketKindFilter] = useState<"All" | TicketStatus>("All");
  const [modal, setModal] = useState<"ticket" | "asset" | null>(null);
  const [selectedTicket, setSelectedTicket] = useState<Ticket | null>(null);
  const [notificationOpen, setNotificationOpen] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const loginInProgress = useRef(false);
  const autoConnectAttempted = useRef(false);
  const [articleSearch, setArticleSearch] = useState("");
  const [selectedArticle, setSelectedArticle] = useState<(typeof articles)[number] | null>(null);
  const [ticketForm, setTicketForm] = useState({
    kind: "Incident" as TicketKind,
    subject: "",
    category: "Hardware",
    priority: "Medium" as TicketPriority,
    description: "",
    assigneeName: "",
    assigneeEmail: "",
  });
  const [memberForm, setMemberForm] = useState({
    email: "",
    name: "",
  });
  const [assetForm, setAssetForm] = useState({
    name: "",
    assetTag: "",
    category: "Laptop",
    assignedTo: "",
    location: "",
  });
  const isAgent = role === "agent" || role === "admin";
  const isAdmin = role === "admin";

  useEffect(() => {
    if (!microsoftConfigured) {
      setAuthLoading(false);
      return;
    }
    let active = true;
    void getCurrentMicrosoftUser().then((nextUser) => {
      if (!active) return;
      setUser(nextUser);
      setAuthError("");
      setAuthLoading(false);
    }).catch((error: unknown) => {
      if (!active) return;
      setAuthError(friendlyError(error));
      setAuthLoading(false);
    });
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!user || database || autoConnectAttempted.current) return;
    if (!incidentsWorkbookUrl || !usersWorkbookUrl) return;
    autoConnectAttempted.current = true;
    void connectSharePointWorkbook(incidentsWorkbookUrl, usersWorkbookUrl);
  }, [user, database]);

  useEffect(() => {
    if (!user || !database) {
      setRole("requester");
      setTickets([]);
      setAssets([]);
      setNotifications([]);
      setMembers([]);
      return;
    }
    try {
      const nextRole = database.getRole(user);
      setRole(nextRole);
      const allTickets = database.getTickets();
      setTickets(nextRole === "requester"
        ? allTickets.filter((ticket) => ticket.requesterUid === user.uid)
        : allTickets);
      setAssets(nextRole === "requester" ? [] : database.getAssets());
      setNotifications(database.getNotifications(user));
      setMembers(nextRole === "admin" ? database.getMembers() : []);
      setDataError("");
    } catch (error) {
      setDataError(`The Excel workbook could not be read: ${friendlyError(error)}`);
    }
  }, [user, database, dataVersion]);

  useEffect(() => {
    if ((page === "Members" && !isAdmin) || (["Assets", "Reports"].includes(page) && !isAgent)) {
      setPage("Dashboard");
    }
  }, [page, role]);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(null), 4800);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  useEffect(() => {
    if ((page === "Members" && !isAdmin) || (["Assets", "Reports"].includes(page) && !isAgent)) {
      setPage("Dashboard");
    }
  }, [page, role]);

  const incidents = useMemo(() => tickets.filter((ticket) => ticket.kind === "Incident"), [tickets]);
  const requests = useMemo(() => tickets.filter((ticket) => ticket.kind === "Service request"), [tickets]);
  const openCount = tickets.filter((ticket) => ticket.status === "New").length;
  const openIncidentCount = incidents.filter((ticket) => ticket.status === "New").length;
  const inProgressCount = tickets.filter((ticket) => ticket.status === "In Progress").length;
  const resolvedTodayCount = tickets.filter((ticket) => {
    if (ticket.status !== "Resolved" || !ticket.updatedAt) return false;
    return ticket.updatedAt.toDateString() === new Date().toDateString();
  }).length;
  const unreadCount = notifications.filter((item) => !item.read).length;

  async function handleMicrosoftLogin() {
    if (loginInProgress.current) return;
    loginInProgress.current = true;
    setBusy(true);
    setAuthError("");
    try {
      await signInWithMicrosoft();
    } catch (error) {
      setAuthError(friendlyError(error));
      loginInProgress.current = false;
      setBusy(false);
    } finally {
      if (loginInProgress.current) {
        window.setTimeout(() => {
          loginInProgress.current = false;
          setBusy(false);
        }, 1000);
      }
    }
  }

  async function handleSignOut() {
    try {
      await signOutMicrosoft();
      setUser(null);
      setDatabase(null);
      autoConnectAttempted.current = false;
      setWorkbookDirty(false);
      setPage("Dashboard");
      setProfileOpen(false);
    } catch (error) {
      setToast({ kind: "error", message: `Could not sign out: ${friendlyError(error)}` });
    }
  }

  async function commitDatabaseChanges(source: ExcelDatabase): Promise<void> {
    setWorkbookDirty(true);
    setDataVersion((version) => version + 1);
    if (source.isRemote) {
      await source.save();
      setWorkbookDirty(false);
    }
  }

  async function connectSharePointWorkbook(incidentsUrl: string, usersUrl: string) {
    if (!incidentsUrl.trim() || !usersUrl.trim()) {
      setDataError("Enter the SharePoint URLs for both the incidents and user details workbooks.");
      return;
    }
    setWorkbookBusy(true);
    setDataError("");
    try {
      const workbook = await ExcelDatabase.fromSharePoint(
        incidentsUrl.trim(),
        usersUrl.trim(),
        getMicrosoftAccessTokenForWorkbook,
      );
      if (user) await workbook.ensureUser(user);
      setDatabase(workbook);
      setWorkbookDirty(workbook.schemaChanged);
      setDataVersion((version) => version + 1);
      if (workbook.schemaChanged) {
        try {
          await workbook.save();
          setWorkbookDirty(false);
        } catch (error) {
          setDataError(`Connected to the SharePoint workbooks, but initial setup could not be saved: ${friendlyError(error)} Close both workbooks in Excel for the web, then select Save workbooks to retry.`);
          return;
        }
      }
      setDataError("");
      setToast({ kind: "success", message: "Connected to the incidents and user details workbooks." });
    } catch (error) {
      setDataError(`SharePoint workbooks could not be opened: ${friendlyError(error)}`);
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function connectLocalWorkbooks(incidentsFile: File, usersFile: File) {
    setWorkbookBusy(true);
    setDataError("");
    try {
      const workbook = await ExcelDatabase.fromFiles(incidentsFile, usersFile);
      if (user) await workbook.ensureUser(user);
      setDatabase(workbook);
      setWorkbookDirty(workbook.schemaChanged);
      setDataVersion((version) => version + 1);
      if (workbook.schemaChanged) setToast({ kind: "info", message: "The workbook tables were added to the in-memory copies. Save to download both updated workbooks." });
      else setToast({ kind: "success", message: "Opened the incidents and user details workbooks." });
    } catch (error) {
      setDataError(friendlyError(error));
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function createNewWorkbook() {
    setWorkbookBusy(true);
    setDataError("");
    try {
      const workbook = await ExcelDatabase.create();
      if (user) {
        await workbook.ensureUser(user);
      }
      setDatabase(workbook);
      setWorkbookDirty(true);
      setDataVersion((version) => version + 1);
      await workbook.download();
      setToast({ kind: "success", message: "The incidents and user details workbook pair was created and downloaded." });
    } catch (error) {
      setDataError(`Excel database could not be created: ${friendlyError(error)}`);
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function saveWorkbook() {
    if (!database) return;
    setWorkbookBusy(true);
    try {
      await database.save();
      setWorkbookDirty(false);
      setDataError("");
      const message = database.storageMode === "SharePoint"
        ? "Changes saved to the connected SharePoint workbook(s)."
        : database.storageMode === "Mixed"
          ? "The SharePoint workbook was saved and the local workbook was downloaded."
          : "Updated Excel workbooks downloaded. Replace both old shared copies with these files.";
      setToast({ kind: "success", message });
    } catch (error) {
      setToast({ kind: "error", message: `Workbook save failed: ${friendlyError(error)}` });
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function reloadWorkbook() {
    if (!database || !user) return;
    if (workbookDirty && !window.confirm("Discard unsaved changes and reload both workbooks?")) return;
    setWorkbookBusy(true);
    try {
      const workbook = await database.reload();
      setDatabase(workbook);
      setWorkbookDirty(workbook.schemaChanged);
      setDataVersion((version) => version + 1);
      setToast({ kind: "success", message: `Reloaded ${workbook.sourceName}.` });
    } catch (error) {
      setToast({ kind: "error", message: `Workbook reload failed: ${friendlyError(error)}` });
    } finally {
      setWorkbookBusy(false);
    }
  }

  function disconnectWorkbook() {
    if (workbookDirty && !window.confirm("There are unsaved changes. Disconnect without saving them?")) return;
    setDatabase(null);
    setWorkbookDirty(false);
    setTickets([]);
    setAssets([]);
    setMembers([]);
    setNotifications([]);
    setRole("requester");
  }

  async function createTicket(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!database || !user) return;
    setBusy(true);
    const ticketId = crypto.randomUUID();
    const ticketNo = `TS-${new Date().getFullYear()}-${ticketId.slice(0, 6).toUpperCase()}`;
    const assigneeName = isAgent ? ticketForm.assigneeName.trim() : "";
    const assigneeEmail = isAgent ? ticketForm.assigneeEmail.trim() : "";
    const newTicket = {
      ticketNo,
      subject: ticketForm.subject.trim(),
      description: ticketForm.description.trim(),
      kind: ticketForm.kind,
      category: ticketForm.category,
      status: "New" as TicketStatus,
      priority: ticketForm.priority,
      requesterUid: user.uid,
      requesterName: user.displayName || firstName(user),
      requesterEmail: user.email || "",
      assigneeName,
      assigneeEmail,
      id: ticketId,
      createdAt: new Date(),
      updatedAt: new Date(),
    };
    try {
      await database.addTicket(newTicket, user);
      await commitDatabaseChanges(database);
      setModal(null);
      setTicketForm({
        kind: "Incident",
        subject: "",
        category: "Hardware",
        priority: "Medium",
        description: "",
        assigneeName: "",
        assigneeEmail: "",
      });
      const emailProblem = await notify(
        ticketId,
        ticketForm.kind === "Incident" ? "Incident logged" : "Service request received",
        isAgent && assigneeName
          ? `${ticketNo}: ${ticketForm.subject.trim()} has been created and assigned to ${assigneeName}.`
          : `${ticketNo}: ${ticketForm.subject.trim()} has been received by the IT service desk.`,
        isAgent && assigneeEmail ? assigneeEmail : user.email || "",
      );
      setToast(emailProblem
        ? { kind: "info", message: `${ticketNo} was created. Notification issue: ${emailProblem}` }
        : { kind: "success", message: `${ticketNo} created and Outlook notification sent.` });
    } catch (error) {
      setToast({ kind: "error", message: `Ticket could not be created: ${friendlyError(error)}` });
    } finally {
      setBusy(false);
    }
  }

  async function notify(
    ticketId: string,
    title: string,
    message: string,
    email: string,
  ): Promise<string> {
    if (!database || !user) return "Excel database is not connected.";
    const notification: AppNotification = {
      id: crypto.randomUUID(),
      recipientUid: email,
      title,
      message,
      ticketId,
      createdAt: new Date(),
      read: false,
      emailStatus: "pending",
      emailError: "",
    };
    let notificationWriteError = "";
    let notificationAdded = false;
    try {
      await database.addNotification(notification);
      notificationAdded = true;
      await commitDatabaseChanges(database);
    } catch (error) {
      notificationWriteError = `In-app notification could not be saved to the workbook: ${friendlyError(error)}`;
    }

    let emailError = "";
    try {
      if (!email) throw new Error("The recipient does not have an email address.");
      await sendMicrosoftMail(email, title, message);
    } catch (error) {
      emailError = friendlyError(error);
    }

    if (notificationAdded) {
      try {
        await database.updateNotification(notification, {
          emailStatus: emailError ? "failed" : "sent",
          emailError,
        });
        await commitDatabaseChanges(database);
        notificationWriteError = "";
      } catch (error) {
        const statusProblem = `Notification status could not be saved: ${friendlyError(error)}`;
        return emailError
          ? `${emailError}; ${statusProblem}`
          : `Outlook sent the email, but ${statusProblem}`;
      }
    }

    if (emailError && notificationWriteError) return `${emailError}; ${notificationWriteError}`;
    if (emailError) return emailError;
    if (notificationWriteError) return `Outlook sent the email, but ${notificationWriteError}`;
    return "";
  }

  async function updateTicket(ticket: Ticket, changes: Partial<Pick<Ticket, "status" | "priority" | "assigneeName" | "assigneeEmail">>) {
    if (!database || !user) return;
    const isRequesterClosure = !isAgent
      && ticket.requesterUid === user.uid
      && ticket.status === "Resolved"
      && changes.status === "Closed"
      && Object.keys(changes).length === 1;
    if (!isAgent && !isRequesterClosure) return;
    setBusy(true);
    const requesterNeedsNotice = isAgent && Boolean(
      (changes.status && changes.status !== ticket.status)
      || (changes.assigneeName && changes.assigneeName !== ticket.assigneeName)
      || (changes.assigneeEmail && changes.assigneeEmail !== ticket.assigneeEmail),
    ) && Boolean(ticket.requesterEmail);
    try {
      await database.updateTicket(ticket, changes, user);
      await commitDatabaseChanges(database);
      setSelectedTicket(null);
      if (requesterNeedsNotice && ticket.requesterEmail) {
        const statusChanged = changes.status !== undefined && changes.status !== ticket.status;
        const subject = statusChanged ? `${ticket.ticketNo} status updated` : `${ticket.ticketNo} assignment updated`;
        const message = statusChanged
          ? `Your ticket "${ticket.subject}" is now ${changes.status}.`
          : `Your ticket "${ticket.subject}" is assigned to ${changes.assigneeName || ticket.assigneeName}.`;
        const emailProblem = await notify(ticket.id, subject, message, ticket.requesterEmail);
        setToast(emailProblem
          ? { kind: "info", message: `Ticket updated. Notification issue: ${emailProblem}` }
          : { kind: "success", message: "Ticket updated and requester notified." });
      } else {
        setToast({ kind: "success", message: "Ticket updated." });
      }
    } catch (error) {
      setToast({ kind: "error", message: `Ticket update failed: ${friendlyError(error)}` });
    } finally {
      setBusy(false);
    }
  }

  async function createAsset(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!database || !isAgent || !user) return;
    setBusy(true);
    try {
      await database.addAsset({
        id: crypto.randomUUID(),
        ...assetForm,
        status: "In use",
        updatedAt: new Date(),
      }, user);
      await commitDatabaseChanges(database);
      setModal(null);
      setAssetForm({ name: "", assetTag: "", category: "Laptop", assignedTo: "", location: "" });
      setToast({ kind: "success", message: "Asset added to the inventory." });
    } catch (error) {
      setToast({ kind: "error", message: `Asset could not be added: ${friendlyError(error)}` });
    } finally {
      setBusy(false);
    }
  }

  async function saveMember(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!database || !isAdmin || !user) return;
    const email = memberForm.email.trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      setToast({ kind: "error", message: "Enter a valid email address." });
      return;
    }
    setBusy(true);
    try {
      await database.addMember(email, memberForm.name.trim() || email.split("@")[0], user);
      await commitDatabaseChanges(database);
      setMemberForm({ email: "", name: "" });
      setToast({ kind: "success", message: "Workspace membership saved." });
    } catch (error) {
      setToast({ kind: "error", message: `Membership could not be saved: ${friendlyError(error)}` });
    } finally {
      setBusy(false);
    }
  }

  async function changeAssetStatus(asset: Asset, status: Asset["status"]) {
    if (!database || !user || !isAgent) return;
    try {
      await database.updateAsset(asset, status, user);
      await commitDatabaseChanges(database);
      setToast({ kind: "success", message: `${asset.assetTag} marked ${status.toLowerCase()}.` });
    } catch (error) {
      setToast({ kind: "error", message: `Asset could not be updated: ${friendlyError(error)}` });
    }
  }

  async function markNotificationRead(notification: AppNotification) {
    if (!database || !notification.read) return;
    try {
      await database.updateNotification(notification, { read: true });
      await commitDatabaseChanges(database);
    } catch (error) {
      setToast({ kind: "error", message: `Notification could not be updated: ${friendlyError(error)}` });
    }
  }

  async function testOutlook() {
    if (!user?.email) return;
    setBusy(true);
    try {
      await sendMicrosoftMail(
        user.email,
        "Onedesk — test notification",
        "Outlook notifications are connected to your Onedesk account.",
      );
      setToast({ kind: "success", message: "Test email sent to your Outlook inbox." });
    } catch (error) {
      setToast({ kind: "error", message: `Test email failed: ${friendlyError(error)}` });
    } finally {
      setBusy(false);
    }
  }

  function exportCsv() {
    const rows = [
      ["Ticket", "Title", "Type", "Status", "Priority", "Category", "Requester", "Assignee", "Created"],
      ...tickets.map((ticket) => [
        ticket.ticketNo,
        ticket.subject,
        ticket.kind,
        ticket.status,
        ticket.priority,
        ticket.category,
        ticket.requesterName,
        ticket.assigneeName,
        ticket.createdAt?.toISOString() || "",
      ]),
    ];
    const safeCell = (value: string) => {
      const formulaSafeValue = /^[\t\r\n ]*[=+\-@]/.test(value) ? `'${value}` : value;
      return `"${formulaSafeValue.replaceAll('"', '""')}"`;
    };
    const csv = rows.map((row) => row.map(safeCell).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `onedesk-report-${new Date().toISOString().slice(0, 10)}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  const filteredTickets = useMemo(() => {
    const source = page === "Requests" ? requests : incidents;
    return source.filter((ticket) => {
      const matchesStatus = ticketKindFilter === "All" || ticket.status === ticketKindFilter;
      const term = search.trim().toLowerCase();
      const matchesSearch = !term || [
        ticket.ticketNo,
        ticket.subject,
        ticket.requesterName,
        ticket.assigneeName,
        ticket.category,
      ].some((value) => value.toLowerCase().includes(term));
      return matchesStatus && matchesSearch;
    });
  }, [incidents, page, requests, search, ticketKindFilter]);

  const filteredArticles = articles.filter((article) =>
    `${article.title} ${article.category}`.toLowerCase().includes(articleSearch.toLowerCase()),
  );

  if (authLoading) return <div className="loading-screen"><div className="loading-mark">O</div><span>Opening your workspace...</span></div>;
  if (!microsoftConfigured) {
    return <ConfigurationScreen />;
  }
  if (!user) {
    return <LoginScreen
      mode={authMode}
      busy={busy}
      error={authError}
      onLogin={() => void handleMicrosoftLogin()}
      onModeChange={setAuthMode}
    />;
  }
  if (!database) {
    return <WorkbookConnectionScreen
      user={user}
      error={dataError}
      busy={workbookBusy}
      onSharePoint={(incidentsUrl, usersUrl) => void connectSharePointWorkbook(incidentsUrl, usersUrl)}
      onLocalFiles={(incidentsFile, usersFile) => void connectLocalWorkbooks(incidentsFile, usersFile)}
      onCreate={() => void createNewWorkbook()}
      onSignOut={() => void handleSignOut()}
    />;
  }

  const visibleNavigation = navigation.filter((item) =>
    (isAgent || !["Assets", "Reports"].includes(item.label))
    && (isAdmin || item.label !== "Members"),
  );
  const navGroups = Array.from(new Set(visibleNavigation.map((item) => item.group)));

  return (
    <div className="app-shell">
      <aside className={`sidebar ${mobileNavOpen ? "mobile-open" : ""}`}>
        <div className="brand">
          <div className="brand-mark"><span>O</span><i /></div>
          <div className="brand-name">Onedesk<span>IT SERVICE DESK</span></div>
          <button className="icon-button sidebar-close" aria-label="Close menu" onClick={() => setMobileNavOpen(false)}><X size={18} /></button>
        </div>
        <button className="workspace-switcher" onClick={() => setPage("Dashboard")}>
          <span className="workspace-avatar">TT</span>
          <span className="workspace-label"><strong>Topin Technologies</strong><small>IT workspace</small></span>
          <ChevronDown size={15} />
        </button>
        <nav className="side-navigation" aria-label="Main navigation">
          {navGroups.map((group) => (
            <div className="nav-group" key={group}>
              <p>{group}</p>
              {visibleNavigation.filter((item) => item.group === group).map(({ label, icon: Icon }) => (
                <button
                  key={label}
                  className={`nav-item ${page === label ? "active" : ""}`}
                  onClick={() => { setPage(label); setMobileNavOpen(false); setSearch(""); }}
                >
                  <Icon size={17} strokeWidth={1.8} />
                  <span>{label}</span>
                  {label === "Incidents" && openIncidentCount > 0 && <span className="nav-count">{openIncidentCount}</span>}
                </button>
              ))}
            </div>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="help-card">
            <div className="help-icon"><CircleHelp size={16} /></div>
            <strong>Need a hand?</strong>
            <span>Our IT team is here to help.</span>
            <button onClick={() => { setTicketForm((form) => ({ ...form, kind: "Service request" })); setModal("ticket"); }}>Get support <ArrowUpRight size={13} /></button>
          </div>
          <div className="sidebar-foot">
            <span className="online-dot" /> All systems operational
          </div>
          <div className="sidebar-copyright">Secure IT. Better work. <span>v1.0</span></div>
        </div>
      </aside>

      {mobileNavOpen && <button className="nav-scrim" aria-label="Close navigation" onClick={() => setMobileNavOpen(false)} />}
      <main className="main-shell">
        <header className="topbar">
          <div className="topbar-left">
            <button className="icon-button mobile-menu" aria-label="Open menu" onClick={() => setMobileNavOpen(true)}><Menu size={19} /></button>
            <div className="breadcrumb"><span>Workspace</span><ChevronRight size={14} /><strong>{page}</strong></div>
          </div>
          <div className="topbar-right">
            <span className={`workbook-indicator ${workbookDirty ? "unsaved" : ""}`} title={database.sourceName}>
              <FileText size={13} />{workbookDirty ? "Unsaved" : "Excel connected"}
            </span>
            <button className="icon-button workbook-action" aria-label="Reload workbook" title="Reload workbook" disabled={workbookBusy} onClick={() => void reloadWorkbook()}><Activity size={16} /></button>
            <button className="button secondary save-workbook-button" disabled={!workbookDirty || workbookBusy} onClick={() => void saveWorkbook()}><Download size={14} />{workbookBusy ? "Saving..." : "Save workbooks"}</button>
            <button className="icon-button workbook-action" aria-label="Change workbook" title="Change workbook" onClick={disconnectWorkbook}><HardDrive size={16} /></button>
            <div className="global-search">
              <Search size={15} />
              <input aria-label="Search workspace" placeholder="Search anything..." value={search} onChange={(event) => { setSearch(event.target.value); if (page !== "Incidents" && page !== "Requests") setPage("Incidents"); }} />
              <kbd><Command size={10} /> K</kbd>
            </div>
            <div className="topbar-divider" />
            <div className="popover-wrap">
              <button className={`icon-button notification-trigger ${notificationOpen ? "pressed" : ""}`} aria-label="Notifications" onClick={() => { setNotificationOpen(!notificationOpen); setProfileOpen(false); }}>
                <Bell size={18} />{unreadCount > 0 && <i />}
              </button>
              {notificationOpen && <div className="popover notification-popover">
                <div className="popover-head"><div><strong>Notifications</strong><span>{unreadCount ? `${unreadCount} unread` : "You're all caught up"}</span></div><button className="text-button" onClick={async () => { for (const item of notifications.filter((entry) => !entry.read)) await markNotificationRead(item); }}>Mark all read</button></div>
                <div className="notification-list">
                  {notifications.length === 0 && <div className="empty-small"><Bell size={19} /><span>No notifications yet</span></div>}
                  {notifications.map((item) => <button key={item.id} className={`notification-item ${item.read ? "" : "unread"}`} onClick={() => { void markNotificationRead(item); setNotificationOpen(false); const ticket = tickets.find((entry) => entry.id === item.ticketId); if (ticket) setSelectedTicket(ticket); }}>
                    <span className="notification-icon"><TicketCheck size={15} /></span>
                    <span><strong>{item.title}</strong><small>{item.message}</small><em>{timeAgo(item.createdAt)}</em></span>
                  </button>)}
                </div>
              </div>}
            </div>
            <div className="popover-wrap">
              <button className="profile-trigger" onClick={() => { setProfileOpen(!profileOpen); setNotificationOpen(false); }}>
                <span className="avatar">{initials(user.displayName || user.email || "IT")}</span>
                <span className="profile-text"><strong>{user.displayName || firstName(user)}</strong><small>{role}</small></span>
                <ChevronDown size={14} />
              </button>
              {profileOpen && <div className="popover profile-popover">
                <div className="profile-popover-user"><span className="avatar">{initials(user.displayName || user.email || "IT")}</span><span><strong>{user.displayName || firstName(user)}</strong><small>{user.email}</small></span></div>
                <button onClick={() => { setPage("Settings"); setProfileOpen(false); }}><Settings size={15} /> Workspace settings</button>
                <button onClick={() => void handleSignOut()}><LogOut size={15} /> Sign out</button>
              </div>}
            </div>
          </div>
        </header>

        <div className="page-content">
          {dataError && <div className="inline-alert error-alert"><CircleHelp size={16} /><span>{dataError}</span><button onClick={() => setDataError("")} aria-label="Dismiss error"><X size={14} /></button></div>}
          {page === "Dashboard" && <Dashboard
            user={user}
            tickets={tickets}
            incidents={incidents}
            assets={assets}
            isAgent={isAgent}
            openCount={openCount}
            inProgressCount={inProgressCount}
            resolvedTodayCount={resolvedTodayCount}
            closureCounts={database.getTicketClosureCounts(tickets)}
            onNewTicket={(kind = "Incident") => { setTicketForm((form) => ({ ...form, kind })); setModal("ticket"); }}
            onSelectTicket={setSelectedTicket}
            onNavigate={setPage}
          />}
          {(page === "Incidents" || page === "Requests") && <TicketList
            page={page}
            tickets={filteredTickets}
            total={page === "Requests" ? requests.length : incidents.length}
            search={search}
            statusFilter={ticketKindFilter}
            onSearch={setSearch}
            onStatusFilter={setTicketKindFilter}
            onSelect={setSelectedTicket}
            onCreate={() => { setTicketForm((form) => ({ ...form, kind: page === "Requests" ? "Service request" : "Incident" })); setModal("ticket"); }}
          />}
          {page === "Assets" && isAgent && <AssetList assets={assets} onCreate={() => setModal("asset")} onStatus={changeAssetStatus} />}
          {page === "Knowledge" && <KnowledgePage search={articleSearch} onSearch={setArticleSearch} articles={filteredArticles} onSelectArticle={setSelectedArticle} onCreateTicket={() => { setTicketForm((form) => ({ ...form, kind: "Service request" })); setModal("ticket"); }} />}
          {page === "Reports" && isAgent && <ReportsPage tickets={tickets} incidents={incidents} requests={requests} assets={assets} onExport={exportCsv} />}
          {page === "Members" && isAdmin && <MembersPage members={members} form={memberForm} busy={busy} currentEmail={user.email} onFormChange={setMemberForm} onSave={saveMember} />}
          {page === "Settings" && <SettingsPage user={user} role={role} database={database} onChangeWorkbook={disconnectWorkbook} busy={busy} emailConfigured={emailNotificationsConfigured} onTestEmail={() => void testOutlook()} />}
        </div>
        <footer className="main-footer"><span>© 2026 Onedesk · Topin Technologies</span><span><i /> Secure connection</span><span>Built for better work <Sparkles size={12} /></span></footer>
      </main>

      {modal === "ticket" && <Modal title={ticketForm.kind === "Incident" ? "Report an incident" : "New service request"} subtitle="Tell us what you need and we'll take it from here." onClose={() => setModal(null)}>
        <form className="form-grid" onSubmit={(event) => void createTicket(event)}>
          <label className="field full-field">Request type
            <select value={ticketForm.kind} onChange={(event) => setTicketForm({ ...ticketForm, kind: event.target.value as TicketKind })}><option>Incident</option><option>Service request</option></select>
          </label>
          <label className="field full-field">Subject <span className="required">*</span>
            <input required maxLength={120} value={ticketForm.subject} placeholder="A short summary of the issue" onChange={(event) => setTicketForm({ ...ticketForm, subject: event.target.value })} />
          </label>
          <label className="field">Category
            <select value={ticketForm.category} onChange={(event) => setTicketForm({ ...ticketForm, category: event.target.value })}>{["Hardware", "Software", "Network", "Access", "Email", "Other"].map((item) => <option key={item}>{item}</option>)}</select>
          </label>
          <label className="field">Priority
            <select value={ticketForm.priority} onChange={(event) => setTicketForm({ ...ticketForm, priority: event.target.value as TicketPriority })}>{["Low", "Medium", "High", "Urgent"].map((item) => <option key={item}>{item}</option>)}</select>
          </label>
          <label className="field full-field">Description <span className="required">*</span>
            <textarea required rows={4} maxLength={4000} value={ticketForm.description} placeholder="Add helpful details, error messages, or steps to reproduce..." onChange={(event) => setTicketForm({ ...ticketForm, description: event.target.value })} />
          </label>
          {isAgent && <>
            <label className="field">Assign to
              <input value={ticketForm.assigneeName} placeholder="Technician name" onChange={(event) => setTicketForm({ ...ticketForm, assigneeName: event.target.value })} />
            </label>
            <label className="field">Assignee email
              <input type="email" value={ticketForm.assigneeEmail} placeholder="technician@example.com" onChange={(event) => setTicketForm({ ...ticketForm, assigneeEmail: event.target.value })} />
            </label>
          </>}
          <div className="form-note full-field"><Mail size={14} /> {isAgent && ticketForm.assigneeEmail ? "Outlook will notify the assigned technician." : "Outlook will send a confirmation to your account; ticket updates will notify you."}</div>
          <div className="modal-actions full-field"><button type="button" className="button secondary" onClick={() => setModal(null)}>Cancel</button><button className="button primary" disabled={busy}><Send size={15} />{busy ? "Submitting..." : "Submit request"}</button></div>
        </form>
      </Modal>}

      {modal === "asset" && isAgent && <Modal title="Add an asset" subtitle="Register a device or resource in your inventory." onClose={() => setModal(null)}>
        <form className="form-grid" onSubmit={(event) => void createAsset(event)}>
          <label className="field full-field">Asset name <span className="required">*</span><input required value={assetForm.name} placeholder="e.g. ThinkPad X1 Carbon" onChange={(event) => setAssetForm({ ...assetForm, name: event.target.value })} /></label>
          <label className="field">Asset tag <span className="required">*</span><input required value={assetForm.assetTag} placeholder="TS-LT-0042" onChange={(event) => setAssetForm({ ...assetForm, assetTag: event.target.value })} /></label>
          <label className="field">Type<select value={assetForm.category} onChange={(event) => setAssetForm({ ...assetForm, category: event.target.value })}>{["Laptop", "Desktop", "Monitor", "Mobile", "Network", "Other"].map((item) => <option key={item}>{item}</option>)}</select></label>
          <label className="field">Assigned to<input value={assetForm.assignedTo} placeholder="Employee name" onChange={(event) => setAssetForm({ ...assetForm, assignedTo: event.target.value })} /></label>
          <label className="field">Location<input value={assetForm.location} placeholder="Office or site" onChange={(event) => setAssetForm({ ...assetForm, location: event.target.value })} /></label>
          <div className="modal-actions full-field"><button type="button" className="button secondary" onClick={() => setModal(null)}>Cancel</button><button className="button primary" disabled={busy}><Plus size={15} />{busy ? "Saving..." : "Add asset"}</button></div>
        </form>
      </Modal>}

      {selectedArticle && <Modal title={selectedArticle.title} subtitle={`${selectedArticle.category} · ${selectedArticle.time}`} onClose={() => setSelectedArticle(null)}>
        <div className="article-modal-body"><p>{selectedArticle.summary}</p><ol>{selectedArticle.steps.map((step) => <li key={step}>{step}</li>)}</ol><div className="form-note"><Headset size={14} /> Need a hand? Submit a service request and the IT team will help.</div><div className="modal-actions"><button className="button secondary" onClick={() => setSelectedArticle(null)}>Close guide</button></div></div>
      </Modal>}

      {selectedTicket && <TicketDetails ticket={selectedTicket} canManage={isAgent} canCloseByUser={role === "requester" && selectedTicket.requesterUid === user.uid} busy={busy} onClose={() => setSelectedTicket(null)} onUpdate={(changes) => void updateTicket(selectedTicket, changes)} />}
      {toast && <div className={`toast ${toast.kind}`} role="status"><span className="toast-mark">{toast.kind === "success" ? <Check size={15} /> : toast.kind === "error" ? <X size={15} /> : <Bell size={15} />}</span>{toast.message}<button onClick={() => setToast(null)} aria-label="Dismiss"><X size={14} /></button></div>}
    </div>
  );
}

function ConfigurationScreen() {
  const missingSettings = [
    !tenantIdConfigured && "VITE_AZURE_TENANT_ID (or VITE_MICROSOFT_TENANT_ID)",
    !clientIdConfigured && "VITE_MICROSOFT_CLIENT_ID",
  ].filter(Boolean);

  return <div className="config-screen">
    <div className="config-card">
      <div className="brand-mark"><span>O</span><i /></div>
      <span className="eyebrow">MICROSOFT SETUP</span>
      <h1>Connect your tenant</h1>
      <p>Microsoft Entra sign-in needs your tenant ID and application client ID.</p>
      <div className="config-steps">
        <div><span>{tenantIdConfigured ? <Check size={13} /> : "1"}</span><div><strong>Topin tenant ID</strong><small>{tenantIdConfigured ? "Configured" : "Add the Topin tenant ID to .env.local."}</small></div></div>
        <div><span>{clientIdConfigured ? <Check size={13} /> : "2"}</span><div><strong>Application client ID</strong><small>{clientIdConfigured ? "Configured" : "Add the app registration client ID to .env.local."}</small></div></div>
      </div>
      <p className="config-foot"><ShieldCheck size={15} />{missingSettings.length ? <>Missing: <code>{missingSettings.join(", ")}</code>. Restart the dev server after updating .env.local.</> : "Configuration is complete. Restart the dev server to enable sign-in."}</p>
    </div>
  </div>;
}

function WorkbookConnectionScreen({ user, error, busy, onSharePoint, onLocalFiles, onCreate, onSignOut }: {
  user: AppUser;
  error: string;
  busy: boolean;
  onSharePoint: (incidentsUrl: string, usersUrl: string) => void;
  onLocalFiles: (incidentsFile: File, usersFile: File) => void;
  onCreate: () => void;
  onSignOut: () => void;
}) {
  const [incidentsFile, setIncidentsFile] = useState<File | null>(null);
  const [usersFile, setUsersFile] = useState<File | null>(null);
  const [incidentsUrl, setIncidentsUrl] = useState(incidentsWorkbookUrl);
  const [usersUrl, setUsersUrl] = useState(usersWorkbookUrl);

  return <div className="config-screen workbook-screen"><section className="config-card workbook-connect-card">
    <div className="workbook-connect-top"><div className="brand-mark"><span>T</span><i /></div><button className="text-button" onClick={onSignOut}><LogOut size={14} /> Sign out</button></div>
    <span className="eyebrow">MICROSOFT CONNECTED</span><h1>Connect your Excel databases</h1><p>Signed in as <strong>{user.email}</strong>. The service desk stores incidents and user details in separate workbooks.</p>
    {error && <div className="inline-alert error-alert"><CircleHelp size={16} /><span>{error}</span></div>}
    <div className="sharepoint-url-fields">
      <label>Incidents workbook SharePoint URL<input type="url" value={incidentsUrl} placeholder="https://.../incidents.xlsx" disabled={busy} onChange={(event) => setIncidentsUrl(event.target.value)} /></label>
      <label>User details workbook SharePoint URL<input type="url" value={usersUrl} placeholder="https://.../users.xlsx" disabled={busy} onChange={(event) => setUsersUrl(event.target.value)} /></label>
    </div>
    <button className="workbook-choice recommended" disabled={busy || !incidentsUrl.trim() || !usersUrl.trim()} onClick={() => onSharePoint(incidentsUrl, usersUrl)}><span className="workbook-choice-icon sharepoint"><BriefcaseBusiness size={19} /></span><span><strong>{busy ? "Connecting to SharePoint..." : "Connect both SharePoint workbooks"}</strong><small>Use Graph to open both files and save changes back to SharePoint.</small></span><ArrowUpRight size={16} /></button>
    <div className="workbook-file-pair">
      <label className="workbook-choice"><span className="workbook-choice-icon"><FileText size={19} /></span><span><strong>{incidentsFile?.name || "Choose incidents workbook"}</strong><small>Incidents, service requests, assets, notifications, audit log, and lookups.</small></span><input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" aria-label="Choose incidents workbook" disabled={busy} onChange={(event) => { setIncidentsFile(event.target.files?.[0] || null); event.currentTarget.value = ""; }} /></label>
      <label className="workbook-choice"><span className="workbook-choice-icon"><UsersRound size={19} /></span><span><strong>{usersFile?.name || "Choose user details workbook"}</strong><small>User identities, contact details, and workspace roles.</small></span><input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" aria-label="Choose user details workbook" disabled={busy} onChange={(event) => { setUsersFile(event.target.files?.[0] || null); event.currentTarget.value = ""; }} /></label>
    </div>
    <button className="button primary workbook-open-button" disabled={busy || !incidentsFile || !usersFile} onClick={() => { if (incidentsFile && usersFile) onLocalFiles(incidentsFile, usersFile); }}><HardDrive size={15} /> Open both selected workbooks</button>
    <p className="workbook-env-hint">Paste each workbook's SharePoint sharing link above. Workbook links are entered after sign-in and are not bundled with the app.</p>
    <button className="workbook-create-link" disabled={busy} onClick={onCreate}><Plus size={14} /> Create and download blank workbook pair</button>
    <div className="workbook-security-note"><ShieldCheck size={15} /><span><strong>Keep both workbooks in restricted SharePoint locations.</strong> Role labels are application data, not a security boundary; Microsoft file permissions determine who can read or change each workbook.</span></div>
  </section></div>;
}

function LoginScreen({ mode, busy, error, onLogin, onModeChange }: {
  mode: "login" | "register";
  busy: boolean;
  error: string;
  onLogin: () => void;
  onModeChange: (mode: "login" | "register") => void;
}) {
  const registering = mode === "register";

  return <div className="login-screen">
    <section className="login-brand">
      <div className="brand"><div className="brand-mark"><span>O</span><i /></div><div className="brand-name">Onedesk<span>IT SERVICE DESK</span></div></div>
      <div className="login-brand-copy"><span className="eyebrow light">{registering ? "WELCOME TO ONEDESK" : "YOUR WORK, IN FLOW"}</span><h1>{registering ? <>Your support<br />starts here.</> : <>Technology that<br />keeps you moving.</>}</h1><p>One place for incidents, service requests, assets, and answers.</p></div>
      <div className="login-brand-bottom"><span><ShieldCheck size={15} /> Topin Technologies</span><span>Powered by Microsoft</span></div>
      <div className="login-art"><div /><div /><div /></div>
    </section>
    <section className="login-panel"><div className="login-panel-inner">
      <span className="eyebrow">{registering ? "NEW TO ONEDESK" : "WELCOME BACK"}</span>
      <h2>{registering ? "Create your workspace profile." : "Your support starts here."}</h2>
      <p>{registering
        ? "Continue with your organization-issued Microsoft account. On your first workbook connection, Onedesk adds your verified profile to the workspace."
        : "Sign in with your company Microsoft account to access the IT workspace."}</p>
      {error && <div className="inline-alert error-alert"><CircleHelp size={16} /><span>{error}</span></div>}
      <button className="microsoft-button" disabled={busy} onClick={onLogin}><MicrosoftMark />{busy ? "Redirecting to Microsoft..." : "Continue with Microsoft"}<ArrowUpRight size={16} /></button>
      {registering && <div className="registration-details">
        <div><Check size={14} /><span><strong>Verified work identity</strong><small>Your name and email come from Microsoft; Onedesk never stores a password.</small></span></div>
        <div><Check size={14} /><span><strong>Requester access by default</strong><small>Only murthy@topin.co.in has administrator access.</small></span></div>
        <div><ShieldCheck size={14} /><span><strong>Workspace access is managed separately</strong><small>You also need permission to open the team's SharePoint workbooks.</small></span></div>
      </div>}
      <div className="login-divider"><span>SECURE COMPANY ACCESS</span></div>
      {!registering && <div className="login-assurance"><div><ShieldCheck size={17} /><span><strong>Protected by your organization</strong><small>Your account is verified by Microsoft Entra ID.</small></span></div><div><Mail size={17} /><span><strong>Outlook notifications</strong><small>Ticket updates arrive right in your inbox.</small></span></div></div>}
      <div className="login-help">{registering
        ? <>Already registered? <button className="auth-mode-link" onClick={() => onModeChange("login")}>Sign in</button></>
        : <>New to Onedesk? <button className="auth-mode-link" onClick={() => onModeChange("register")}>Register here</button><br />Need help signing in? Contact your IT administrator.</>}</div>
    </div><div className="login-footer">© 2026 Onedesk · Topin Technologies <span>·</span> Made for better work</div></section>
  </div>;
}

function MicrosoftMark() {
  return <span className="microsoft-mark" aria-hidden="true"><i /><i /><i /><i /></span>;
}

function Dashboard({ user, tickets, incidents, assets, isAgent, openCount, inProgressCount, resolvedTodayCount, closureCounts, onNewTicket, onSelectTicket, onNavigate }: {
  user: AppUser; tickets: Ticket[]; incidents: Ticket[]; assets: Asset[]; isAgent: boolean; openCount: number; inProgressCount: number; resolvedTodayCount: number; closureCounts: { byAdmin: number; byUser: number };
  onNewTicket: (kind?: TicketKind) => void; onSelectTicket: (ticket: Ticket) => void; onNavigate: (page: Page) => void;
}) {
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const months = Array.from({ length: 6 }, (_, index) => {
    const date = new Date();
    date.setMonth(date.getMonth() - (5 - index));
    return { label: monthNames[date.getMonth()], year: date.getFullYear(), month: date.getMonth() };
  });
  const monthlyCounts = months.map((month) => tickets.filter((ticket) => {
    if (!ticket.createdAt) return false;
    const date = ticket.createdAt;
    return date.getMonth() === month.month && date.getFullYear() === month.year;
  }).length);
  const chartMax = Math.max(...monthlyCounts, 1);
  const latestTickets = [...tickets].slice(0, 5);
  const resolvedCount = tickets.filter((ticket) => ticket.status === "Resolved").length;
  const pendingCount = tickets.filter((ticket) => ticket.status === "Pending").length;
  const maintenanceCount = tickets.filter((ticket) => ticket.status === "Under Maintenance").length;
  const closedCount = tickets.filter((ticket) => ticket.status === "Closed").length;
  const unattributedClosedCount = Math.max(0, closedCount - closureCounts.byAdmin - closureCounts.byUser);
  const statusBreakdown = [
    { label: "Open", value: openCount, color: "#ef8b53" },
    { label: "In progress", value: inProgressCount, color: "#668fcb" },
    { label: "Pending", value: pendingCount, color: "#d0a344" },
    { label: "Under maintenance", value: maintenanceCount, color: "#987ac6" },
    { label: "Resolved", value: resolvedCount, color: "#60a879" },
    { label: "Closed by admin", value: closureCounts.byAdmin, color: "#347c53" },
    { label: "Closed by user", value: closureCounts.byUser, color: "#5eaa9a" },
    { label: "Other closed", value: unattributedClosedCount, color: "#a3aaa5" },
  ];
  const greeting = new Date().getHours() < 12 ? "Good morning" : new Date().getHours() < 18 ? "Good afternoon" : "Good evening";
  const todayLabel = new Date().toLocaleDateString("en-GB", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  }).toUpperCase();
  const first = firstName(user);

  return <div className="dashboard-page">
    <div className="page-heading dashboard-heading">
      <div><div className="eyebrow"><span className="eyebrow-dot" /> {todayLabel}</div><h1>{greeting}, {first}<span className="wave">✳</span></h1><p>Here’s what’s happening across IT today.</p></div>
      <div className="heading-actions">{isAgent && <button className="button secondary" onClick={() => onNavigate("Reports")}><CalendarDays size={15} /> View reports</button>}<button className="button primary" onClick={() => onNewTicket("Incident")}><Plus size={16} /> Create ticket</button></div>
    </div>
    <div className="welcome-banner"><div className="welcome-banner-icon"><Zap size={19} /></div><div><strong>Your IT workspace, all in one place.</strong><span>Track issues, request help, and stay on top of your technology.</span></div><button onClick={() => onNavigate("Knowledge")}>Explore the help center <ArrowUpRight size={14} /></button><div className="banner-orbit one" /><div className="banner-orbit two" /></div>
    <div className="stat-grid ticket-status-grid">
      <StatCard label="Open tickets" value={openCount} caption="Awaiting a first response" icon={LifeBuoy} accent="orange" trend={<><ArrowUpRight size={13} /> Active</>} />
      <StatCard label="Work in progress" value={inProgressCount} caption="Currently being worked" icon={Activity} accent="blue" trend={<><span className="trend-dot" /> In motion</>} />
      <StatCard label="Resolved" value={resolvedCount} caption={`${resolvedTodayCount} resolved today`} icon={CheckCheck} accent="green" trend={<><ArrowDownRight size={13} /> Complete</>} />
      <StatCard label="Closed by admin" value={closureCounts.byAdmin} caption="Closed by service desk" icon={ShieldCheck} accent="green" trend={<><span className="trend-dot" /> Admin</>} />
      <StatCard label="Closed by user" value={closureCounts.byUser} caption="Confirmed by requester" icon={UserRound} accent="purple" trend={<><span className="trend-dot" /> Requester</>} />
      <StatCard label="Under maintenance" value={maintenanceCount} caption="In planned maintenance" icon={HardDrive} accent="blue" trend={<><span className="trend-dot" /> Maintenance</>} />
      <StatCard label="Other states" value={pendingCount + unattributedClosedCount} caption="Pending or legacy closures" icon={Clock3} accent="orange" trend={<><span className="trend-dot" /> Pending</>} />
    </div>
    <div className="dashboard-grid">
      <section className="panel chart-panel">
        <div className="panel-heading"><div><span className="panel-kicker">OVERVIEW</span><h2>{isAgent ? "Ticket activity" : "Your activity"}</h2><p>{isAgent ? "Tickets created over the last 6 months" : "Your tickets created over the last 6 months"}</p></div>{isAgent && <button className="more-button" onClick={() => onNavigate("Reports")} aria-label="Open reports"><MoreHorizontal size={20} /></button>}</div>
        <div className="chart-legend"><span><i /> Incidents & requests</span><span className="chart-total">{tickets.length} total</span></div>
        <div className="bar-chart">
          <div className="chart-y-axis"><span>{chartMax}</span><span>{Math.round(chartMax / 2)}</span><span>0</span></div>
          <div className="chart-plot">
            <div className="chart-grid-lines"><i /><i /><i /></div>
            <div className="bars">{months.map((month, index) => {
              const height = Math.max(4, (monthlyCounts[index] / chartMax) * 100);
              return <div className="bar-column" key={`${month.label}-${month.year}`}><span className="bar-value">{monthlyCounts[index]}</span><div className={`bar ${index === months.length - 1 ? "current" : ""}`} style={{ height: `${height}%` }} /><span className={`bar-label ${index === months.length - 1 ? "current-label" : ""}`}>{month.label}</span></div>;
            })}</div>
          </div>
        </div>
      </section>
      <section className="panel status-chart-panel">
        <div className="panel-heading"><div><span className="panel-kicker">AT A GLANCE</span><h2>Ticket status</h2><p>Current distribution across {tickets.length} tickets</p></div><span className="live-indicator"><i /> LIVE</span></div>
        <div className="status-chart-content">
          <StatusDonut entries={statusBreakdown} total={tickets.length} />
          <div className="status-legend">{statusBreakdown.map((item) => <div key={item.label}><i style={{ backgroundColor: item.color }} /><span>{item.label}</span><strong>{item.value}</strong></div>)}</div>
        </div>
        <button className="panel-link" onClick={() => onNavigate("Incidents")}>View all tickets <ChevronRight size={15} /></button>
      </section>
    </div>
    <section className="panel recent-panel">
      <div className="panel-heading recent-heading"><div><span className="panel-kicker">THE LATEST</span><h2>Recent tickets</h2><p>Your team's newest conversations</p></div><button className="button quiet" onClick={() => onNavigate("Incidents")}>View all tickets <ChevronRight size={14} /></button></div>
      {latestTickets.length ? <div className="table-wrap"><table><thead><tr><th>Ticket</th><th>Subject</th><th>Requester</th><th>Priority</th><th>Status</th><th>Created</th><th /></tr></thead><tbody>{latestTickets.map((ticket) => <tr key={ticket.id} onClick={() => onSelectTicket(ticket)}><td><span className="ticket-number">{ticket.ticketNo}</span></td><td><div className="ticket-subject-cell"><strong>{ticket.subject}</strong><span>{ticket.category}</span></div></td><td><div className="person-cell"><span className="mini-avatar">{initials(ticket.requesterName)}</span>{ticket.requesterName}</div></td><td><PriorityPill priority={ticket.priority} /></td><td><StatusPill status={ticket.status} /></td><td className="date-cell">{dateLabel(ticket.createdAt)}</td><td><ChevronRight size={15} className="row-chevron" /></td></tr>)}</tbody></table></div> : <EmptyState icon={TicketPlus} title="Your queue is clear" text="When a new ticket comes in, it’ll show up right here." action="Create your first ticket" onAction={() => onNewTicket()} />}
      <div className="recent-footer"><span><span className="online-dot" /> {incidents.length} incidents <i /> {tickets.length - incidents.length} requests</span><span><Clock3 size={13} /> Updated just now</span></div>
    </section>
    <div className="dashboard-bottom">
      <button className="quick-card" onClick={() => onNewTicket("Service request")}><span className="quick-icon peach"><TicketPlus size={19} /></span><span><strong>Request a service</strong><small>Get access, equipment, or software</small></span><ArrowUpRight size={16} /></button>
      {isAgent && <button className="quick-card" onClick={() => onNavigate("Assets")}><span className="quick-icon mint"><HardDrive size={19} /></span><span><strong>Browse your assets</strong><small>{assets.length} items in the inventory</small></span><ArrowUpRight size={16} /></button>}
      <button className="quick-card" onClick={() => onNavigate("Knowledge")}><span className="quick-icon lilac"><BookOpen size={19} /></span><span><strong>Find an answer</strong><small>Explore guides from the IT team</small></span><ArrowUpRight size={16} /></button>
    </div>
  </div>;
}

function StatusDonut({ entries, total }: { entries: { label: string; value: number; color: string }[]; total: number }) {
  const radius = 35;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;
  return <div className="status-donut" role="img" aria-label={`Ticket status breakdown: ${entries.map((item) => `${item.label} ${item.value}`).join(", ")}`}>
    <svg viewBox="0 0 100 100" aria-hidden="true">
      <circle className="donut-track" cx="50" cy="50" r={radius} />
      {entries.filter((item) => item.value > 0).map((item) => {
        const length = total ? (item.value / total) * circumference : 0;
        const circle = <circle key={item.label} cx="50" cy="50" r={radius} stroke={item.color} strokeDasharray={`${length} ${circumference - length}`} strokeDashoffset={-offset} />;
        offset += length;
        return circle;
      })}
    </svg>
    <div><strong>{total}</strong><span>tickets</span></div>
  </div>;
}

function StatCard({ label, value, caption, icon: Icon, accent, trend }: { label: string; value: number; caption: string; icon: typeof LayoutDashboard; accent: string; trend: ReactNode }) {
  return <div className="stat-card"><div className="stat-top"><span>{label}</span><span className={`stat-icon ${accent}`}><Icon size={17} /></span></div><strong className="stat-value">{value.toLocaleString()}</strong><div className="stat-bottom"><span>{caption}</span><span className={`stat-trend ${accent}`}>{trend}</span></div></div>;
}

function TicketList({ page, tickets, total, search, statusFilter, onSearch, onStatusFilter, onSelect, onCreate }: {
  page: "Incidents" | "Requests"; tickets: Ticket[]; total: number; search: string; statusFilter: "All" | TicketStatus;
  onSearch: (value: string) => void; onStatusFilter: (value: "All" | TicketStatus) => void; onSelect: (ticket: Ticket) => void; onCreate: () => void;
}) {
  return <div className="list-page">
    <div className="page-heading"><div><div className="eyebrow"><span className="eyebrow-dot" /> SERVICE DESK</div><h1>{page}</h1><p>{page === "Incidents" ? "Track, triage, and resolve reported issues." : "Manage employee requests from intake to fulfilment."}</p></div><button className="button primary" onClick={onCreate}><Plus size={16} /> {page === "Incidents" ? "Report incident" : "New request"}</button></div>
    <div className="list-summary">
      <div><span className="list-summary-icon"><TicketCheck size={17} /></span><span><strong>{total}</strong><small>Total {page.toLowerCase()}</small></span></div>
      <div><span className="list-summary-icon orange"><Clock3 size={17} /></span><span><strong>{tickets.filter((item) => item.status !== "Resolved" && item.status !== "Closed").length}</strong><small>Awaiting resolution</small></span></div>
      <div><span className="list-summary-icon mint"><CheckCheck size={17} /></span><span><strong>{tickets.filter((item) => item.status === "Resolved").length}</strong><small>Resolved</small></span></div>
      <span className="summary-caption"><span className="online-dot" /> Excel workbook</span>
    </div>
    <section className="panel list-panel">
      <div className="list-toolbar">
        <div className="filter-tabs">{(["All", "New", "In Progress", "Pending", "Under Maintenance", "Resolved", "Closed"] as const).map((item) => <button key={item} className={statusFilter === item ? "selected" : ""} onClick={() => onStatusFilter(item)}>{item}{item === "All" && <span>{total}</span>}</button>)}</div>
        <div className="list-toolbar-right"><label className="table-search"><Search size={15} /><input placeholder="Search tickets..." value={search} onChange={(event) => onSearch(event.target.value)} /></label><button className="icon-button filter-button" aria-label="Filter tickets" onClick={() => onStatusFilter(statusFilter === "All" ? "New" : "All")}><Filter size={16} /></button></div>
      </div>
      {tickets.length ? <div className="table-wrap"><table><thead><tr><th>Ticket</th><th>Subject</th><th>Requester</th><th>Assignee</th><th>Priority</th><th>Status</th><th>Created</th><th /></tr></thead><tbody>{tickets.map((ticket) => <tr key={ticket.id} onClick={() => onSelect(ticket)}><td><span className="ticket-number">{ticket.ticketNo}</span></td><td><div className="ticket-subject-cell"><strong>{ticket.subject}</strong><span>{ticket.category} <i /> {ticket.kind}</span></div></td><td><div className="person-cell"><span className="mini-avatar">{initials(ticket.requesterName)}</span>{ticket.requesterName}</div></td><td><div className="person-cell assignee-cell"><span className="mini-avatar assignee">{initials(ticket.assigneeName)}</span>{ticket.assigneeName || "Unassigned"}</div></td><td><PriorityPill priority={ticket.priority} /></td><td><StatusPill status={ticket.status} /></td><td className="date-cell">{dateLabel(ticket.createdAt)}</td><td><ChevronRight size={15} className="row-chevron" /></td></tr>)}</tbody></table></div> : <EmptyState icon={Search} title={search || statusFilter !== "All" ? "No matching tickets" : `No ${page.toLowerCase()} yet`} text={search || statusFilter !== "All" ? "Try another search term or status filter." : "Create a ticket to get your service desk moving."} action={search || statusFilter !== "All" ? undefined : "Create a ticket"} onAction={onCreate} />}
      <div className="list-footer"><span>Showing <strong>{tickets.length}</strong> of <strong>{total}</strong> tickets</span><span>SharePoint changes sync automatically; local files need saving <Activity size={13} /></span></div>
    </section>
  </div>;
}

function AssetList({ assets, onCreate, onStatus }: { assets: Asset[]; onCreate: () => void; onStatus: (asset: Asset, status: Asset["status"]) => void }) {
  const [search, setSearch] = useState("");
  const filtered = assets.filter((asset) => `${asset.name} ${asset.assetTag} ${asset.assignedTo} ${asset.location}`.toLowerCase().includes(search.toLowerCase()));
  return <div className="list-page">
    <div className="page-heading"><div><div className="eyebrow"><span className="eyebrow-dot" /> CONFIGURATION MANAGEMENT</div><h1>Asset inventory</h1><p>A clear view of the devices and resources your team depends on.</p></div><button className="button primary" onClick={onCreate}><Plus size={16} /> Add asset</button></div>
    <div className="asset-stats"><div><span className="asset-stat-icon mint"><Laptop size={18} /></span><span><strong>{assets.length}</strong><small>Total assets</small></span></div><div><span className="asset-stat-icon blue"><UserRound size={18} /></span><span><strong>{assets.filter((item) => item.status === "In use").length}</strong><small>In use</small></span></div><div><span className="asset-stat-icon orange"><BriefcaseBusiness size={18} /></span><span><strong>{assets.filter((item) => item.status === "In repair").length}</strong><small>In repair</small></span></div><div className="inventory-tip"><ShieldCheck size={16} /><span>Keep your inventory up to date to speed up support.</span></div></div>
    <section className="panel list-panel"><div className="list-toolbar assets-toolbar"><div><span className="panel-kicker">YOUR ORGANIZATION</span><h2>All assets <span className="asset-count">{assets.length}</span></h2></div><label className="table-search"><Search size={15} /><input placeholder="Search assets..." value={search} onChange={(event) => setSearch(event.target.value)} /></label></div>
      {filtered.length ? <div className="table-wrap"><table><thead><tr><th>Asset</th><th>Asset tag</th><th>Assigned to</th><th>Location</th><th>Last updated</th><th>Status</th><th /></tr></thead><tbody>{filtered.map((asset) => <tr key={asset.id}><td><div className="asset-name-cell"><span className="asset-device"><Laptop size={17} /></span><span><strong>{asset.name}</strong><small>{asset.category}</small></span></div></td><td><span className="ticket-number">{asset.assetTag}</span></td><td>{asset.assignedTo || <span className="muted">Unassigned</span>}</td><td>{asset.location || <span className="muted">Not set</span>}</td><td className="date-cell">{dateLabel(asset.updatedAt)}</td><td><select className={`asset-status-select ${asset.status.toLowerCase().replaceAll(" ", "-")}`} value={asset.status} aria-label={`Status for ${asset.assetTag}`} onChange={(event) => void onStatus(asset, event.target.value as Asset["status"])}>{["In use", "In repair", "Available", "Retired"].map((status) => <option key={status}>{status}</option>)}</select></td><td><MoreHorizontal size={17} className="row-chevron" /></td></tr>)}</tbody></table></div> : <EmptyState icon={HardDrive} title={search ? "No matching assets" : "Your inventory is waiting"} text={search ? "Try searching by asset name, tag, owner, or location." : "Add your first asset to start managing your inventory."} action={search ? undefined : "Add an asset"} onAction={onCreate} />}
      <div className="list-footer"><span>Showing <strong>{filtered.length}</strong> of <strong>{assets.length}</strong> assets</span><span>SharePoint changes sync automatically; local files need saving <Activity size={13} /></span></div>
    </section>
  </div>;
}

function KnowledgePage({ search, onSearch, articles: visibleArticles, onSelectArticle, onCreateTicket }: { search: string; onSearch: (value: string) => void; articles: typeof articles; onSelectArticle: (article: (typeof articles)[number]) => void; onCreateTicket: () => void }) {
  return <div className="knowledge-page">
    <section className="knowledge-hero"><span className="eyebrow light">THAI SUMMIT HELP CENTER</span><h1>Answers, before you ask.</h1><p>Guides from your IT team to help you get back to what matters.</p><label className="knowledge-search"><Search size={18} /><input placeholder="Search articles, topics, or keywords..." value={search} onChange={(event) => onSearch(event.target.value)} /><kbd>⌘ K</kbd></label><div className="knowledge-topics"><span>Popular:</span>{["Outlook", "VPN", "Laptop setup", "Password"].map((topic) => <button key={topic} onClick={() => onSearch(topic)}>{topic}</button>)}</div></section>
    <div className="knowledge-title"><div><span className="panel-kicker">KNOWLEDGE BASE</span><h2>{search ? "Search results" : "Popular articles"}</h2></div><button className="button quiet" onClick={() => onSearch("")}>View all guides <ChevronRight size={14} /></button></div>
    {visibleArticles.length ? <div className="article-grid">{visibleArticles.map((article, index) => <button className="article-card" key={article.title} onClick={() => onSelectArticle(article)}><span className={`article-icon article-${index % 4}`}><article.icon size={18} /></span><span className="article-category">{article.category}</span><strong>{article.title}</strong><span className="article-meta"><span>{article.time}</span><ArrowUpRight size={14} /></span></button>)}</div> : <EmptyState icon={BookOpen} title="No articles found" text="Try a broader search or browse all of our guides." action="Clear search" onAction={() => onSearch("")} />}
    <div className="knowledge-help"><div className="help-icon"><Headset size={18} /></div><div><strong>Still need a human?</strong><span>Our service desk is happy to help with anything else.</span></div><button className="button primary" onClick={onCreateTicket}>Ask IT for help <ArrowUpRight size={14} /></button></div>
  </div>;
}

function ReportsPage({ tickets, incidents, requests, assets, onExport }: { tickets: Ticket[]; incidents: Ticket[]; requests: Ticket[]; assets: Asset[]; onExport: () => void }) {
  const resolved = tickets.filter((item) => item.status === "Resolved").length;
  const resolutionRate = tickets.length ? Math.round(resolved / tickets.length * 100) : 0;
  const byPriority = (["Urgent", "High", "Medium", "Low"] as TicketPriority[]).map((priority) => ({ priority, count: tickets.filter((item) => item.priority === priority).length }));
  return <div className="reports-page">
    <div className="page-heading"><div><div className="eyebrow"><span className="eyebrow-dot" /> SERVICE ANALYTICS</div><h1>Reports</h1><p>Understand demand, spot patterns, and keep service moving.</p></div><button className="button primary" onClick={onExport}><Download size={15} /> Export ticket data</button></div>
    <div className="reports-summary"><div className="report-summary-card"><span className="report-card-icon orange"><TicketPlus size={17} /></span><span><strong>{tickets.length}</strong><small>Total tickets</small></span><span className="report-card-foot">All time</span></div><div className="report-summary-card"><span className="report-card-icon blue"><LifeBuoy size={17} /></span><span><strong>{incidents.length}</strong><small>Incidents</small></span><span className="report-card-foot">Reported issues</span></div><div className="report-summary-card"><span className="report-card-icon purple"><TicketCheck size={17} /></span><span><strong>{requests.length}</strong><small>Service requests</small></span><span className="report-card-foot">Employee requests</span></div><div className="report-summary-card"><span className="report-card-icon mint"><CheckCheck size={17} /></span><span><strong>{resolutionRate}%</strong><small>Resolution rate</small></span><span className="report-card-foot">Of all tickets</span></div></div>
    <div className="reports-grid"><section className="panel report-panel"><div className="panel-heading"><div><span className="panel-kicker">PRIORITY BREAKDOWN</span><h2>Tickets by priority</h2><p>Where the team should focus first</p></div><span className="report-heading-icon"><Activity size={16} /></span></div><div className="priority-bars">{byPriority.map(({ priority, count }) => <div className="priority-bar-row" key={priority}><span><i className={`priority-dot ${priority.toLowerCase()}`} />{priority}</span><div className="priority-track"><i className={`priority-fill ${priority.toLowerCase()}`} style={{ width: `${tickets.length ? Math.max(count ? 6 : 0, count / tickets.length * 100) : 0}%` }} /></div><strong>{count}</strong></div>)}</div></section>
      <section className="panel report-panel"><div className="panel-heading"><div><span className="panel-kicker">ASSET LIFECYCLE</span><h2>Inventory health</h2><p>Current asset distribution</p></div><span className="report-heading-icon"><HardDrive size={16} /></span></div><div className="inventory-total"><strong>{assets.length}</strong><span>tracked assets</span><div className="inventory-progress"><i style={{ width: `${assets.length ? assets.filter((item) => item.status === "In use").length / assets.length * 100 : 0}%` }} /></div><div className="inventory-legend"><span><i className="pulse-resolved" /> In use <strong>{assets.filter((item) => item.status === "In use").length}</strong></span><span><i className="pulse-open" /> In repair <strong>{assets.filter((item) => item.status === "In repair").length}</strong></span><span><i className="pulse-progress" /> Available <strong>{assets.filter((item) => item.status === "Available").length}</strong></span></div></div></section></div>
    <div className="report-note"><FileText size={16} /><span>Reports reflect the tickets and assets in the connected Excel workbook. Export a CSV for further analysis.</span><button className="text-button" onClick={onExport}>Download CSV <Download size={13} /></button></div>
  </div>;
}

function MembersPage({ members, form, busy, currentEmail, onFormChange, onSave }: {
  members: Member[];
  form: { email: string; name: string };
  busy: boolean;
  currentEmail: string;
  onFormChange: (value: { email: string; name: string }) => void;
  onSave: (event: FormEvent<HTMLFormElement>) => void;
}) {
  return <div className="members-page">
    <div className="page-heading"><div><div className="eyebrow"><span className="eyebrow-dot" /> ACCESS MANAGEMENT</div><h1>Workspace members</h1><p>Manage requester access to the service desk.</p></div><span className="member-count"><UsersRound size={15} /> {members.length} members</span></div>
    <div className="members-grid">
      <section className="panel member-add-panel">
        <div className="panel-heading"><div><span className="panel-kicker">ADD OR UPDATE ACCESS</span><h2>Add a requester</h2><p>Register a Microsoft account for workspace access.</p></div><span className="report-heading-icon"><UserRound size={16} /></span></div>
        <form className="member-form" onSubmit={onSave}>
          <label className="field full-field">Microsoft email<input required type="email" value={form.email} placeholder="employee@example.com" onChange={(event) => onFormChange({ ...form, email: event.target.value })} /></label>
          <label className="field">Display name<input value={form.name} placeholder="Employee name" onChange={(event) => onFormChange({ ...form, name: event.target.value })} /></label>
          <div className="members-note"><ShieldCheck size={14} /><span>Only murthy@topin.co.in has the administrator role. All other accounts are requesters; manage workbook permissions separately in SharePoint.</span></div>
          <button className="button primary" disabled={busy}><Plus size={15} />{busy ? "Saving..." : "Save requester"}</button>
        </form>
      </section>
      <section className="panel member-list-panel">
        <div className="panel-heading"><div><span className="panel-kicker">ONEDESK · TOPIN TECHNOLOGIES</span><h2>People with access</h2><p>All members are requesters except the administrator account.</p></div></div>
        {members.length ? <div className="table-wrap"><table><thead><tr><th>Member</th><th>Role</th></tr></thead><tbody>{members.map((member) => <tr key={member.uid}><td><div className="person-cell"><span className="mini-avatar">{initials(member.name || member.email)}</span><span><strong className="member-name">{member.name || member.email}</strong><small className="member-email">{member.email}</small></span></div></td><td><span className={`role-badge ${member.role}`}>{member.email.toLowerCase() === currentEmail.toLowerCase() ? `${member.role} · you` : member.role}</span></td></tr>)}</tbody></table></div> : <EmptyState icon={UsersRound} title="No managed members yet" text="Members are added when they sign in or when you save a requester." />}
        <div className="member-list-foot"><span><i className="role-dot requester" /> Requester</span><span><i className="role-dot admin" /> Admin · murthy@topin.co.in</span></div>
      </section>
    </div>
  </div>;
}

function SettingsPage({ user, role, database, onChangeWorkbook, busy, emailConfigured, onTestEmail }: { user: AppUser; role: UserRole; database: ExcelDatabase; onChangeWorkbook: () => void; busy: boolean; emailConfigured: boolean; onTestEmail: () => void }) {
  return <div className="settings-page">
    <div className="page-heading"><div><div className="eyebrow"><span className="eyebrow-dot" /> PREFERENCES & INTEGRATIONS</div><h1>Settings</h1><p>Manage how your workspace is connected and how updates reach your team.</p></div></div>
    <div className="settings-grid">    <section className="panel settings-panel"><div className="settings-panel-heading"><span className="settings-icon microsoft-setting"><MicrosoftMark /></span><div><span className="panel-kicker">ACCOUNT</span><h2>Microsoft identity</h2><p>Your Microsoft Entra account and workspace access.</p></div><span className="connected-badge"><i /> Connected</span></div><div className="settings-row"><span className="settings-row-icon"><UserRound size={16} /></span><span><strong>Signed-in account</strong><small>{user.displayName || "Microsoft user"}</small></span><span className="settings-value">{user.email}</span></div><div className="settings-row"><span className="settings-row-icon"><ShieldCheck size={16} /></span><span><strong>Microsoft account</strong><small>Signed in securely through Microsoft</small></span><span className="settings-value">{user.email}</span></div><div className="settings-row"><span className="settings-row-icon"><KeyIcon /></span><span><strong>Workspace role</strong><small>Controls access to shared service-desk data</small></span><span className="settings-value">{role}</span></div></section>
      <section className="panel settings-panel"><div className="settings-panel-heading"><span className="settings-icon outlook-setting"><Mail size={18} /></span><div><span className="panel-kicker">NOTIFICATIONS</span><h2>Outlook email</h2><p>Send ticket updates from the admin mailbox.</p></div><span className={`connected-badge ${emailConfigured ? "" : "needs-setup"}`}><i /> {emailConfigured ? "Configured" : "Needs setup"}</span></div><div className="outlook-info"><div className="outlook-info-icon"><Send size={16} /></div><p>Ticket emails are sent by the secure Azure notification service from murthy@topin.co.in. The mailbox password is not stored in this app.</p></div><button className="button secondary test-email-button" disabled={busy} onClick={onTestEmail}><Send size={14} />{busy ? "Sending..." : "Send me a test email"}</button><div className="settings-permission"><ShieldCheck size={14} /> Microsoft Graph · Azure managed identity</div></section>
      <section className="panel settings-panel workbook-panel"><div className="settings-panel-heading"><span className="settings-icon workbook-setting"><FileText size={18} /></span><div><span className="panel-kicker">DATA STORAGE</span><h2>Excel databases</h2><p>{database.sourceName}</p></div><span className="connected-badge"><i /> {database.storageMode}</span></div><div className="collection-chips"><span>Incidents: Tickets <strong>Table</strong></span><span>Assets <strong>Table</strong></span><span>Notifications <strong>Table</strong></span><span>Audit log <strong>Table</strong></span><span>User details: Members <strong>Table</strong></span></div><button className="button secondary test-email-button" onClick={onChangeWorkbook}><HardDrive size={14} /> Change workbooks</button><div className="settings-permission"><ShieldCheck size={14} /> Keep both shared workbooks in restricted SharePoint locations</div></section>
      <div className="settings-tip"><Sparkles size={16} /><span><strong>Good to know</strong>Your administrator manages Microsoft consent and workspace membership.</span></div>
    </div>
  </div>;
}

function KeyIcon() {
  return <ShieldCheck size={16} />;
}

function TicketDetails({ ticket, canManage, canCloseByUser, busy, onClose, onUpdate }: { ticket: Ticket; canManage: boolean; canCloseByUser: boolean; busy: boolean; onClose: () => void; onUpdate: (changes: Partial<Pick<Ticket, "status" | "priority" | "assigneeName" | "assigneeEmail">>) => void }) {
  const [status, setStatus] = useState<TicketStatus>(ticket.status);
  const [priority, setPriority] = useState<TicketPriority>(ticket.priority);
  const [assigneeName, setAssigneeName] = useState(ticket.assigneeName || "");
  const [assigneeEmail, setAssigneeEmail] = useState(ticket.assigneeEmail || "");
  const changed = status !== ticket.status || priority !== ticket.priority || assigneeName !== ticket.assigneeName || assigneeEmail !== ticket.assigneeEmail;
  return <div className="detail-scrim" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><aside className="ticket-drawer">
    <div className="drawer-header"><span className="eyebrow">TICKET DETAILS</span><button className="icon-button" aria-label="Close ticket details" onClick={onClose}><X size={17} /></button></div>
    <div className="drawer-content"><div className="drawer-ticket-id"><span className="ticket-number">{ticket.ticketNo}</span><span className={`kind-chip ${ticket.kind === "Incident" ? "incident" : "request"}`}>{ticket.kind}</span></div><h2>{ticket.subject}</h2><div className="drawer-pills"><StatusPill status={ticket.status} /><PriorityPill priority={ticket.priority} /></div>
      <div className="detail-section"><span className="panel-kicker">DESCRIPTION</span><p>{ticket.description || "No additional description provided."}</p></div>
      <div className="detail-meta"><div><span>Category</span><strong><Tag size={14} />{ticket.category}</strong></div><div><span>Created</span><strong><CalendarDays size={14} />{dateLabel(ticket.createdAt)}</strong></div></div>
      <div className="detail-person"><span className="panel-kicker">REQUESTER</span><div className="detail-person-row"><span className="avatar">{initials(ticket.requesterName)}</span><span><strong>{ticket.requesterName}</strong><small>{ticket.requesterEmail}</small></span></div></div>
      {canManage
        ? <div className="detail-section edit-section"><span className="panel-kicker">WORKFLOW</span><label className="field">Status<select value={status} onChange={(event) => setStatus(event.target.value as TicketStatus)}>{["New", "In Progress", "Pending", "Under Maintenance", "Resolved", "Closed"].map((item) => <option key={item}>{item}</option>)}</select></label><label className="field">Priority<select value={priority} onChange={(event) => setPriority(event.target.value as TicketPriority)}>{["Low", "Medium", "High", "Urgent"].map((item) => <option key={item}>{item}</option>)}</select></label><label className="field">Assignee name<input value={assigneeName} onChange={(event) => setAssigneeName(event.target.value)} /></label><label className="field">Assignee email<input type="email" value={assigneeEmail} onChange={(event) => setAssigneeEmail(event.target.value)} /></label></div>
        : <div className="detail-section"><span className="panel-kicker">ASSIGNED TO</span><p>{ticket.assigneeName || "The service desk will assign a technician."}</p></div>}
      {canManage && <div className="drawer-email-note"><Mail size={14} /> The requester will be notified by email when you save changes.</div>}
    </div>
    <div className="drawer-footer"><button className="button secondary" onClick={onClose}>Close</button>{canCloseByUser && ticket.status === "Resolved" && <button className="button primary" disabled={busy} onClick={() => onUpdate({ status: "Closed" })}>{busy ? "Closing..." : "Close ticket"}</button>}{canManage && <button className="button primary" disabled={busy || !changed} onClick={() => onUpdate({ status, priority, assigneeName, assigneeEmail })}>{busy ? "Saving..." : "Save changes"}</button>}</div>
  </aside></div>;
}

function Modal({ title, subtitle, onClose, children }: { title: string; subtitle: string; onClose: () => void; children: ReactNode }) {
  return <div className="modal-scrim" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="modal-card"><div className="modal-header"><div><span className="eyebrow">ONEDESK · TOPIN TECHNOLOGIES</span><h2>{title}</h2><p>{subtitle}</p></div><button className="icon-button" aria-label="Close dialog" onClick={onClose}><X size={17} /></button></div>{children}</section></div>;
}

function EmptyState({ icon: Icon, title, text, action, onAction }: { icon: typeof LayoutDashboard; title: string; text: string; action?: string; onAction?: () => void }) {
  return <div className="empty-state"><span className="empty-state-icon"><Icon size={20} /></span><strong>{title}</strong><p>{text}</p>{action && onAction && <button className="button secondary" onClick={onAction}><Plus size={14} />{action}</button>}</div>;
}

export default App;
