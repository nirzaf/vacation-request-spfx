/**
 * Interface representing a Leave Request item from SharePoint
 */
export interface ILeaveRequest {
  Id: number;
  Title: string;
  Requester: {
    Id: number;
    Title: string;
    EMail: string;
  };
  EmployeeID?: string;
  Department?: string;
  Manager?: {
    Id: number;
    Title: string;
    EMail: string;
  };
  LeaveType: {
    Id: number;
    Title: string;
  };
  StartDate: Date;
  EndDate: Date;
  TotalDays?: number;
  IsPartialDay: boolean;
  PartialDayHours?: number;
  RequestComments?: string;
  ApprovalStatus: ApprovalStatus;
  ApprovalDate?: Date;
  ApprovalComments?: string;
  SubmissionDate: Date;
  LastModified: Date;
  AttachmentURL?: string;
  WorkflowInstanceID?: string;
  NotificationsSent: boolean;
  CalendarEventID?: string;
}

/**
 * Interface for creating a new Leave Request
 */
export interface ILeaveRequestCreate {
  Title?: string; // Will be auto-generated if not provided
  LeaveTypeId: number;
  StartDate: Date;
  EndDate: Date;
  IsPartialDay: boolean;
  PartialDayHours?: number;
  RequestComments?: string;
  AttachmentURL?: string;
}

/**
 * Interface for updating a Leave Request
 */
export interface ILeaveRequestUpdate {
  LeaveTypeId?: number;
  StartDate?: Date;
  EndDate?: Date;
  IsPartialDay?: boolean;
  PartialDayHours?: number;
  RequestComments?: string;
  AttachmentURL?: string;
  ApprovalStatus?: ApprovalStatus;
  ApprovalComments?: string;
  ApprovalDate?: Date;
  CalendarEventID?: string;
}

/**
 * Enum for Leave Request Approval Status
 */
export enum ApprovalStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled'
}

/**
 * Interface for Leave Request form data
 */
export interface ILeaveRequestFormData {
  leaveTypeId: number;
  startDate: Date;
  endDate: Date;
  isPartialDay: boolean;
  partialDayHours?: number;
  comments?: string;
  attachmentUrl?: string;
}

/**
 * Interface for Leave Request validation result
 */
export interface ILeaveRequestValidation {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Interface for Leave Request calendar event
 */
export interface ILeaveRequestCalendarEvent {
  id: string;
  subject: string;
  start: Date;
  end: Date;
  isAllDay: boolean;
  showAs: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere';
  categories: string[];
}

/**
 * Interface for Leave Request statistics
 */
export interface ILeaveRequestStats {
  totalRequests: number;
  pendingRequests: number;
  approvedRequests: number;
  rejectedRequests: number;
  totalDaysRequested: number;
  totalDaysApproved: number;
}

/**
 * Interface for Leave Request filter options
 */
export interface ILeaveRequestFilter {
  employeeId?: number;
  managerId?: number;
  leaveTypeId?: number;
  status?: ApprovalStatus;
  startDateFrom?: Date;
  startDateTo?: Date;
  endDateFrom?: Date;
  endDateTo?: Date;
  department?: string;
}

/**
 * Interface for Leave Request list view options
 */
export interface ILeaveRequestViewOptions {
  filter?: ILeaveRequestFilter;
  orderBy?: string;
  orderDirection?: 'asc' | 'desc';
  top?: number;
  skip?: number;
  select?: string[];
  expand?: string[];
}
