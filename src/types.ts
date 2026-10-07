export type TicketStatus = "New" | "In Progress" | "Pending" | "Under Maintenance" | "Resolved" | "Closed";
export type TicketPriority = "Low" | "Medium" | "High" | "Urgent";
export type TicketKind = "Incident" | "Service request";
export type UserRole = "requester" | "agent" | "admin";

export interface AppUser {
  uid: string;
  displayName: string;
  email: string;
}

export interface Member {
  uid: string;
  email: string;
  name: string;
  role: UserRole;
  createdAt: Date | null;
  updatedAt: Date | null;
}

export interface Ticket {
  id: string;
  ticketNo: string;
  subject: string;
  description: string;
  kind: TicketKind;
  category: string;
  status: TicketStatus;
  priority: TicketPriority;
  requesterUid: string;
  requesterName: string;
  requesterEmail: string;
  assigneeName: string;
  assigneeEmail: string;
  createdAt: Date | null;
  updatedAt: Date | null;
}

export interface Asset {
  id: string;
  name: string;
  assetTag: string;
  category: string;
  assignedTo: string;
  location: string;
  status: "In use" | "In repair" | "Available" | "Retired";
  updatedAt: Date | null;
}

export interface AppNotification {
  id: string;
  recipientUid: string;
  title: string;
  message: string;
  ticketId: string;
  createdAt: Date | null;
  read: boolean;
  emailStatus: string;
  emailError: string;
}
