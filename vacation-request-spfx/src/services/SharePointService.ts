import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  ILeaveRequest,
  ILeaveRequestCreate,
  ILeaveRequestUpdate,
  ILeaveType,
  ILeaveBalance,
  LIST_NAMES,
  FIELD_NAMES,
  ApprovalStatus,
  CommonUtils
} from '../models';

/**
 * Service class for SharePoint operations
 */
export class SharePointService {
  private context: WebPartContext;
  private sp: ReturnType<typeof spfi>;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get all active leave types
   */
  public async getLeaveTypes(): Promise<ILeaveType[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_TYPES)
        .items
        .select('Id', 'Title', 'Description', 'IsActive', 'RequiresApproval',
                'MaxDaysPerRequest', 'RequiresDocumentation', 'ColorCode', 'PolicyURL', 'Created', 'Modified')
        .filter('IsActive eq true')
        .orderBy('Title')();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Description: item.Description,
        IsActive: item.IsActive,
        RequiresApproval: item.RequiresApproval,
        MaxDaysPerRequest: item.MaxDaysPerRequest,
        RequiresDocumentation: item.RequiresDocumentation,
        ColorCode: item.ColorCode,
        PolicyURL: item.PolicyURL,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified)
      }));
    } catch (error) {
      console.error('Error fetching leave types:', error);
      throw new Error('Failed to fetch leave types');
    }
  }

  /**
   * Get leave type by ID
   */
  public async getLeaveTypeById(id: number): Promise<ILeaveType | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_TYPES)
        .items
        .getById(id)
        .select('Id', 'Title', 'Description', 'IsActive', 'RequiresApproval',
                'MaxDaysPerRequest', 'RequiresDocumentation', 'ColorCode', 'PolicyURL', 'Created', 'Modified')();

      return {
        Id: item.Id,
        Title: item.Title,
        Description: item.Description,
        IsActive: item.IsActive,
        RequiresApproval: item.RequiresApproval,
        MaxDaysPerRequest: item.MaxDaysPerRequest,
        RequiresDocumentation: item.RequiresDocumentation,
        ColorCode: item.ColorCode,
        PolicyURL: item.PolicyURL,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified)
      };
    } catch (error) {
      console.error('Error fetching leave type:', error);
      return null;
    }
  }

  /**
   * Create a new leave request
   */
  public async createLeaveRequest(request: ILeaveRequestCreate): Promise<ILeaveRequest> {
    try {
      // Generate title if not provided
      const title = request.Title || CommonUtils.generateLeaveRequestTitle(
        this.context.pageContext.user.displayName,
        'Leave Request',
        request.StartDate
      );

      const itemData = {
        Title: title,
        [FIELD_NAMES.LEAVE_REQUESTS.LEAVE_TYPE + 'Id']: request.LeaveTypeId,
        [FIELD_NAMES.LEAVE_REQUESTS.START_DATE]: CommonUtils.formatDateForSharePoint(request.StartDate),
        [FIELD_NAMES.LEAVE_REQUESTS.END_DATE]: CommonUtils.formatDateForSharePoint(request.EndDate),
        [FIELD_NAMES.LEAVE_REQUESTS.IS_PARTIAL_DAY]: request.IsPartialDay,
        [FIELD_NAMES.LEAVE_REQUESTS.PARTIAL_DAY_HOURS]: request.PartialDayHours,
        [FIELD_NAMES.LEAVE_REQUESTS.REQUEST_COMMENTS]: request.RequestComments,
        [FIELD_NAMES.LEAVE_REQUESTS.ATTACHMENT_URL]: request.AttachmentURL,
        [FIELD_NAMES.LEAVE_REQUESTS.REQUESTER + 'Id']: this.context.pageContext.user.loginName,
        [FIELD_NAMES.LEAVE_REQUESTS.APPROVAL_STATUS]: ApprovalStatus.Pending,
        [FIELD_NAMES.LEAVE_REQUESTS.SUBMISSION_DATE]: CommonUtils.formatDateForSharePoint(new Date()),
        [FIELD_NAMES.LEAVE_REQUESTS.NOTIFICATIONS_SENT]: false
      };

      const result = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_REQUESTS)
        .items
        .add(itemData);

      // Fetch the created item with all fields
      return await this.getLeaveRequestById(result.data.Id);
    } catch (error) {
      console.error('Error creating leave request:', error);
      throw new Error('Failed to create leave request');
    }
  }

  /**
   * Get leave request by ID
   */
  public async getLeaveRequestById(id: number): Promise<ILeaveRequest> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_REQUESTS)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'StartDate', 'EndDate', 'TotalDays', 'IsPartialDay', 'PartialDayHours',
          'RequestComments', 'ApprovalStatus', 'ApprovalDate', 'ApprovalComments',
          'SubmissionDate', 'Modified', 'AttachmentURL', 'NotificationsSent',
          'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'Manager/Id', 'Manager/Title', 'Manager/EMail',
          'LeaveType/Id', 'LeaveType/Title'
        )
        .expand('Requester', 'Manager', 'LeaveType')();

      return {
        Id: item.Id,
        Title: item.Title,
        Requester: {
          Id: item.Requester?.Id || 0,
          Title: item.Requester?.Title || '',
          EMail: item.Requester?.EMail || ''
        },
        Manager: item.Manager ? {
          Id: item.Manager.Id,
          Title: item.Manager.Title,
          EMail: item.Manager.EMail
        } : undefined,
        LeaveType: {
          Id: item.LeaveType?.Id || 0,
          Title: item.LeaveType?.Title || ''
        },
        StartDate: new Date(item.StartDate),
        EndDate: new Date(item.EndDate),
        TotalDays: item.TotalDays,
        IsPartialDay: item.IsPartialDay,
        PartialDayHours: item.PartialDayHours,
        RequestComments: item.RequestComments,
        ApprovalStatus: item.ApprovalStatus as ApprovalStatus,
        ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
        ApprovalComments: item.ApprovalComments,
        SubmissionDate: new Date(item.SubmissionDate),
        LastModified: new Date(item.Modified),
        AttachmentURL: item.AttachmentURL,
        NotificationsSent: item.NotificationsSent
      };
    } catch (error) {
      console.error('Error fetching leave request:', error);
      throw new Error('Failed to fetch leave request');
    }
  }

  /**
   * Get current user's leave requests
   */
  public async getCurrentUserLeaveRequests(): Promise<ILeaveRequest[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_REQUESTS)
        .items
        .select(
          'Id', 'Title', 'StartDate', 'EndDate', 'TotalDays', 'IsPartialDay', 'PartialDayHours',
          'RequestComments', 'ApprovalStatus', 'ApprovalDate', 'ApprovalComments',
          'SubmissionDate', 'Modified', 'AttachmentURL', 'NotificationsSent',
          'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'Manager/Id', 'Manager/Title', 'Manager/EMail',
          'LeaveType/Id', 'LeaveType/Title'
        )
        .expand('Requester', 'Manager', 'LeaveType')
        .filter(`Requester/Id eq ${this.context.pageContext.user.loginName}`)
        .orderBy('SubmissionDate', false)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Requester: {
          Id: item.Requester?.Id || 0,
          Title: item.Requester?.Title || '',
          EMail: item.Requester?.EMail || ''
        },
        Manager: item.Manager ? {
          Id: item.Manager.Id,
          Title: item.Manager.Title,
          EMail: item.Manager.EMail
        } : undefined,
        LeaveType: {
          Id: item.LeaveType?.Id || 0,
          Title: item.LeaveType?.Title || ''
        },
        StartDate: new Date(item.StartDate),
        EndDate: new Date(item.EndDate),
        TotalDays: item.TotalDays,
        IsPartialDay: item.IsPartialDay,
        PartialDayHours: item.PartialDayHours,
        RequestComments: item.RequestComments,
        ApprovalStatus: item.ApprovalStatus as ApprovalStatus,
        ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
        ApprovalComments: item.ApprovalComments,
        SubmissionDate: new Date(item.SubmissionDate),
        LastModified: new Date(item.Modified),
        AttachmentURL: item.AttachmentURL,
        NotificationsSent: item.NotificationsSent
      }));
    } catch (error) {
      console.error('Error fetching user leave requests:', error);
      throw new Error('Failed to fetch leave requests');
    }
  }

  /**
   * Get current user's leave balances
   */
  public async getCurrentUserLeaveBalances(): Promise<ILeaveBalance[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_BALANCES)
        .items
        .select(
          'Id', 'TotalAllowance', 'UsedDays', 'RemainingDays', 'CarryOverDays',
          'EffectiveDate', 'ExpirationDate',
          'Employee/Id', 'Employee/Title', 'Employee/EMail',
          'LeaveType/Id', 'LeaveType/Title'
        )
        .expand('Employee', 'LeaveType')
        .filter(`Employee/Id eq ${this.context.pageContext.user.loginName}`)();

      return items.map((item: any) => ({
        Id: item.Id,
        Employee: {
          Id: item.Employee?.Id || 0,
          Title: item.Employee?.Title || '',
          EMail: item.Employee?.EMail || ''
        },
        LeaveType: {
          Id: item.LeaveType?.Id || 0,
          Title: item.LeaveType?.Title || ''
        },
        TotalAllowance: item.TotalAllowance,
        UsedDays: item.UsedDays,
        RemainingDays: item.RemainingDays,
        CarryOverDays: item.CarryOverDays,
        EffectiveDate: new Date(item.EffectiveDate),
        ExpirationDate: new Date(item.ExpirationDate)
      }));
    } catch (error) {
      console.error('Error fetching leave balances:', error);
      throw new Error('Failed to fetch leave balances');
    }
  }

  /**
   * Validate leave request against business rules
   */
  public async validateLeaveRequest(request: ILeaveRequestCreate): Promise<{ isValid: boolean; errors: string[] }> {
    const errors: string[] = [];

    try {
      // Get leave type details
      const leaveType = await this.getLeaveTypeById(request.LeaveTypeId);
      if (!leaveType) {
        errors.push('Invalid leave type selected');
        return { isValid: false, errors };
      }

      // Validate dates
      if (request.EndDate < request.StartDate) {
        errors.push('End date must be after start date');
      }

      // Validate partial day hours
      if (request.IsPartialDay) {
        if (!request.PartialDayHours || request.PartialDayHours <= 0 || request.PartialDayHours > 8) {
          errors.push('Partial day hours must be between 0.5 and 8 hours');
        }
      }

      // Validate max days per request
      if (leaveType.MaxDaysPerRequest) {
        const totalDays = CommonUtils.calculateBusinessDays(request.StartDate, request.EndDate);
        if (totalDays > leaveType.MaxDaysPerRequest) {
          errors.push(`Maximum ${leaveType.MaxDaysPerRequest} days allowed for ${leaveType.Title}`);
        }
      }

      // Validate documentation requirement
      if (leaveType.RequiresDocumentation && !request.AttachmentURL) {
        errors.push(`Documentation is required for ${leaveType.Title}`);
      }

      // Validate past dates (except for emergency leave)
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      if (request.StartDate < today && leaveType.Title !== 'Emergency Leave') {
        errors.push('Cannot request leave for past dates');
      }

      return { isValid: errors.length === 0, errors };
    } catch (error) {
      console.error('Error validating leave request:', error);
      return { isValid: false, errors: ['Validation failed'] };
    }
  }

  /**
   * Update leave request
   */
  public async updateLeaveRequest(id: number, updates: ILeaveRequestUpdate): Promise<ILeaveRequest> {
    try {
      const updateData: any = {};

      if (updates.LeaveTypeId !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.LEAVE_TYPE + 'Id'] = updates.LeaveTypeId;
      }
      if (updates.StartDate !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.START_DATE] = CommonUtils.formatDateForSharePoint(updates.StartDate);
      }
      if (updates.EndDate !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.END_DATE] = CommonUtils.formatDateForSharePoint(updates.EndDate);
      }
      if (updates.IsPartialDay !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.IS_PARTIAL_DAY] = updates.IsPartialDay;
      }
      if (updates.PartialDayHours !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.PARTIAL_DAY_HOURS] = updates.PartialDayHours;
      }
      if (updates.RequestComments !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.REQUEST_COMMENTS] = updates.RequestComments;
      }
      if (updates.AttachmentURL !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.ATTACHMENT_URL] = updates.AttachmentURL;
      }
      if (updates.ApprovalStatus !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.APPROVAL_STATUS] = updates.ApprovalStatus;
      }
      if (updates.ApprovalComments !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.APPROVAL_COMMENTS] = updates.ApprovalComments;
      }
      if (updates.ApprovalDate !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.APPROVAL_DATE] = CommonUtils.formatDateForSharePoint(updates.ApprovalDate);
      }
      if (updates.CalendarEventID !== undefined) {
        updateData[FIELD_NAMES.LEAVE_REQUESTS.CALENDAR_EVENT_ID] = updates.CalendarEventID;
      }

      await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_REQUESTS)
        .items
        .getById(id)
        .update(updateData);

      // Return updated item
      return await this.getLeaveRequestById(id);
    } catch (error) {
      console.error('Error updating leave request:', error);
      throw new Error('Failed to update leave request');
    }
  }

  /**
   * Get all leave requests (for calendar and admin views)
   */
  public async getAllLeaveRequests(): Promise<ILeaveRequest[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.LEAVE_REQUESTS)
        .items
        .select(
          'Id', 'Title', 'StartDate', 'EndDate', 'TotalDays', 'IsPartialDay', 'PartialDayHours',
          'RequestComments', 'ApprovalStatus', 'ApprovalDate', 'ApprovalComments',
          'SubmissionDate', 'Modified', 'AttachmentURL', 'NotificationsSent', 'Department',
          'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'Manager/Id', 'Manager/Title', 'Manager/EMail',
          'LeaveType/Id', 'LeaveType/Title'
        )
        .expand('Requester', 'Manager', 'LeaveType')
        .orderBy('StartDate', false)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Requester: {
          Id: item.Requester?.Id || 0,
          Title: item.Requester?.Title || '',
          EMail: item.Requester?.EMail || ''
        },
        Department: item.Department,
        Manager: item.Manager ? {
          Id: item.Manager.Id,
          Title: item.Manager.Title,
          EMail: item.Manager.EMail
        } : undefined,
        LeaveType: {
          Id: item.LeaveType?.Id || 0,
          Title: item.LeaveType?.Title || ''
        },
        StartDate: new Date(item.StartDate),
        EndDate: new Date(item.EndDate),
        TotalDays: item.TotalDays,
        IsPartialDay: item.IsPartialDay,
        PartialDayHours: item.PartialDayHours,
        RequestComments: item.RequestComments,
        ApprovalStatus: item.ApprovalStatus as ApprovalStatus,
        ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
        ApprovalComments: item.ApprovalComments,
        SubmissionDate: new Date(item.SubmissionDate),
        LastModified: new Date(item.Modified),
        AttachmentURL: item.AttachmentURL,
        NotificationsSent: item.NotificationsSent
      }));
    } catch (error) {
      console.error('Error fetching all leave requests:', error);
      throw new Error('Failed to fetch leave requests');
    }
  }
}
