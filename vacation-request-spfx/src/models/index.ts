/**
 * Export all model interfaces and utilities
 */

// Leave Request models
export * from './ILeaveRequest';

// Leave Type models
export * from './ILeaveType';

// Leave Balance models
export * from './ILeaveBalance';

/**
 * Common interfaces used across the application
 */

/**
 * Interface for SharePoint user information
 */
export interface IUser {
  Id: number;
  Title: string;
  EMail: string;
  LoginName?: string;
  Department?: string;
  JobTitle?: string;
}

/**
 * Interface for SharePoint lookup field
 */
export interface ILookupValue {
  Id: number;
  Title: string;
}

/**
 * Interface for API response wrapper
 */
export interface IApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

/**
 * Interface for paginated results
 */
export interface IPaginatedResult<T> {
  items: T[];
  totalCount: number;
  hasMore: boolean;
  nextSkip?: number;
}

/**
 * Interface for sort options
 */
export interface ISortOption {
  field: string;
  direction: 'asc' | 'desc';
}

/**
 * Interface for common list operations
 */
export interface IListOperations<T, TCreate, TUpdate> {
  getAll(): Promise<T[]>;
  getById(id: number): Promise<T | null>;
  create(item: TCreate): Promise<T>;
  update(id: number, item: TUpdate): Promise<T>;
  delete(id: number): Promise<boolean>;
}

/**
 * Common constants
 */
export const LIST_NAMES = {
  LEAVE_REQUESTS: 'LeaveRequests',
  LEAVE_TYPES: 'LeaveTypes',
  LEAVE_BALANCES: 'LeaveBalances'
} as const;

export const FIELD_NAMES = {
  LEAVE_REQUESTS: {
    REQUESTER: 'Requester',
    EMPLOYEE_ID: 'EmployeeID',
    DEPARTMENT: 'Department',
    MANAGER: 'Manager',
    LEAVE_TYPE: 'LeaveType',
    START_DATE: 'StartDate',
    END_DATE: 'EndDate',
    TOTAL_DAYS: 'TotalDays',
    IS_PARTIAL_DAY: 'IsPartialDay',
    PARTIAL_DAY_HOURS: 'PartialDayHours',
    REQUEST_COMMENTS: 'RequestComments',
    APPROVAL_STATUS: 'ApprovalStatus',
    APPROVAL_DATE: 'ApprovalDate',
    APPROVAL_COMMENTS: 'ApprovalComments',
    SUBMISSION_DATE: 'SubmissionDate',
    ATTACHMENT_URL: 'AttachmentURL',
    WORKFLOW_INSTANCE_ID: 'WorkflowInstanceID',
    NOTIFICATIONS_SENT: 'NotificationsSent',
    CALENDAR_EVENT_ID: 'CalendarEventID'
  },
  LEAVE_TYPES: {
    TITLE: 'Title',
    DESCRIPTION: 'Description',
    IS_ACTIVE: 'IsActive',
    REQUIRES_APPROVAL: 'RequiresApproval',
    MAX_DAYS_PER_REQUEST: 'MaxDaysPerRequest',
    REQUIRES_DOCUMENTATION: 'RequiresDocumentation',
    COLOR_CODE: 'ColorCode',
    POLICY_URL: 'PolicyURL'
  },
  LEAVE_BALANCES: {
    EMPLOYEE: 'Employee',
    LEAVE_TYPE: 'LeaveType',
    TOTAL_ALLOWANCE: 'TotalAllowance',
    USED_DAYS: 'UsedDays',
    REMAINING_DAYS: 'RemainingDays',
    CARRY_OVER_DAYS: 'CarryOverDays',
    EFFECTIVE_DATE: 'EffectiveDate',
    EXPIRATION_DATE: 'ExpirationDate'
  }
} as const;

/**
 * Utility functions for common operations
 */
export class CommonUtils {
  /**
   * Format date for SharePoint
   */
  public static formatDateForSharePoint(date: Date): string {
    return date.toISOString();
  }

  /**
   * Parse SharePoint date
   */
  public static parseSharePointDate(dateString: string): Date {
    return new Date(dateString);
  }

  /**
   * Generate title for leave request
   */
  public static generateLeaveRequestTitle(employeeName: string, leaveTypeName: string, startDate: Date): string {
    const formattedDate = startDate.toLocaleDateString();
    return `${employeeName} - ${leaveTypeName} - ${formattedDate}`;
  }

  /**
   * Calculate business days between two dates
   */
  public static calculateBusinessDays(startDate: Date, endDate: Date): number {
    let count = 0;
    const current = new Date(startDate.getTime());

    while (current <= endDate) {
      const dayOfWeek = current.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Not Sunday (0) or Saturday (6)
        count++;
      }
      current.setDate(current.getDate() + 1);
    }

    return count;
  }

  /**
   * Validate email format
   */
  public static isValidEmail(email: string): boolean {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  /**
   * Truncate text to specified length
   */
  public static truncateText(text: string, maxLength: number): string {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * Deep clone object
   */
  public static deepClone<T>(obj: T): T {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * Check if date is weekend
   */
  public static isWeekend(date: Date): boolean {
    const dayOfWeek = date.getDay();
    return dayOfWeek === 0 || dayOfWeek === 6; // Sunday or Saturday
  }

  /**
   * Get next business day
   */
  public static getNextBusinessDay(date: Date): Date {
    const nextDay = new Date(date.getTime());
    nextDay.setDate(nextDay.getDate() + 1);

    while (this.isWeekend(nextDay)) {
      nextDay.setDate(nextDay.getDate() + 1);
    }

    return nextDay;
  }
}
