import { 
  ILeaveRequestCreate, 
  ILeaveType, 
  ILeaveBalance, 
  CommonUtils 
} from '../models';

/**
 * Interface for validation result
 */
export interface IValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Interface for conflict detection result
 */
export interface IConflictResult {
  hasConflicts: boolean;
  conflicts: IConflictDetail[];
}

/**
 * Interface for conflict detail
 */
export interface IConflictDetail {
  type: 'team-member' | 'blackout-date' | 'holiday' | 'overlap';
  message: string;
  severity: 'error' | 'warning';
  details?: any;
}

/**
 * Service class for validation and business rules
 */
export class ValidationService {
  
  /**
   * Validate leave request against business rules
   */
  public static validateLeaveRequest(
    request: ILeaveRequestCreate,
    leaveType: ILeaveType,
    leaveBalance?: ILeaveBalance,
    existingRequests?: any[]
  ): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Basic field validation
    this.validateBasicFields(request, errors);

    // Date validation
    this.validateDates(request, errors, warnings);

    // Leave type specific validation
    this.validateLeaveType(request, leaveType, errors, warnings);

    // Balance validation
    if (leaveBalance) {
      this.validateBalance(request, leaveBalance, errors, warnings);
    }

    // Overlap validation
    if (existingRequests) {
      this.validateOverlaps(request, existingRequests, errors, warnings);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate basic required fields
   */
  private static validateBasicFields(request: ILeaveRequestCreate, errors: string[]): void {
    if (!request.LeaveTypeId) {
      errors.push('Leave type is required');
    }

    if (!request.StartDate) {
      errors.push('Start date is required');
    }

    if (!request.EndDate) {
      errors.push('End date is required');
    }

    if (request.IsPartialDay && (!request.PartialDayHours || request.PartialDayHours <= 0)) {
      errors.push('Partial day hours must be specified for partial day requests');
    }

    if (request.PartialDayHours && (request.PartialDayHours < 0.5 || request.PartialDayHours > 8)) {
      errors.push('Partial day hours must be between 0.5 and 8 hours');
    }
  }

  /**
   * Validate dates
   */
  private static validateDates(
    request: ILeaveRequestCreate, 
    errors: string[], 
    warnings: string[]
  ): void {
    if (!request.StartDate || !request.EndDate) {
      return; // Already handled in basic validation
    }

    // End date must be after or equal to start date
    if (request.EndDate < request.StartDate) {
      errors.push('End date must be after or equal to start date');
      return;
    }

    // Check for past dates (with some exceptions)
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    if (request.StartDate < today) {
      errors.push('Cannot request leave for past dates');
    }

    // Warn about short notice (less than 2 business days)
    const businessDaysNotice = CommonUtils.calculateBusinessDays(today, request.StartDate);
    if (businessDaysNotice < 2 && businessDaysNotice >= 0) {
      warnings.push('Short notice: Consider providing more advance notice for leave requests');
    }

    // Warn about weekend dates
    if (CommonUtils.isWeekend(request.StartDate) || CommonUtils.isWeekend(request.EndDate)) {
      warnings.push('Leave request includes weekend dates');
    }
  }

  /**
   * Validate against leave type rules
   */
  private static validateLeaveType(
    request: ILeaveRequestCreate,
    leaveType: ILeaveType,
    errors: string[],
    warnings: string[]
  ): void {
    // Check if leave type is active
    if (!leaveType.IsActive) {
      errors.push(`Leave type "${leaveType.Title}" is not currently available`);
      return;
    }

    // Check maximum days per request
    if (leaveType.MaxDaysPerRequest) {
      const requestedDays = CommonUtils.calculateBusinessDays(request.StartDate, request.EndDate);
      if (requestedDays > leaveType.MaxDaysPerRequest) {
        errors.push(
          `Maximum ${leaveType.MaxDaysPerRequest} days allowed per request for ${leaveType.Title}. ` +
          `You requested ${requestedDays} days.`
        );
      }
    }

    // Check documentation requirement
    if (leaveType.RequiresDocumentation && !request.AttachmentURL) {
      errors.push(`Documentation is required for ${leaveType.Title} requests`);
    }

    // Partial day validation for specific leave types
    if (request.IsPartialDay && leaveType.Title.toLowerCase().includes('maternity')) {
      warnings.push('Partial day requests are unusual for maternity/paternity leave');
    }
  }

  /**
   * Validate against leave balance
   */
  private static validateBalance(
    request: ILeaveRequestCreate,
    balance: ILeaveBalance,
    errors: string[],
    warnings: string[]
  ): void {
    const requestedDays = request.IsPartialDay && request.PartialDayHours ? 
      request.PartialDayHours / 8 : 
      CommonUtils.calculateBusinessDays(request.StartDate, request.EndDate);

    // Check if sufficient balance
    if (requestedDays > balance.RemainingDays) {
      errors.push(
        `Insufficient leave balance. Requested: ${requestedDays} days, ` +
        `Available: ${balance.RemainingDays} days`
      );
    }

    // Warn if using significant portion of balance
    const usagePercentage = (requestedDays / balance.TotalAllowance) * 100;
    if (usagePercentage > 50) {
      warnings.push(
        `This request will use ${usagePercentage.toFixed(1)}% of your annual ${balance.LeaveType.Title} allowance`
      );
    }

    // Check expiration
    const today = new Date();
    const daysUntilExpiry = Math.ceil(
      (balance.ExpirationDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24)
    );
    
    if (daysUntilExpiry <= 30 && daysUntilExpiry > 0) {
      warnings.push(
        `Your ${balance.LeaveType.Title} balance expires in ${daysUntilExpiry} days. ` +
        `Consider using remaining days before expiration.`
      );
    }
  }

  /**
   * Validate for overlapping requests
   */
  private static validateOverlaps(
    request: ILeaveRequestCreate,
    existingRequests: any[],
    errors: string[],
    warnings: string[]
  ): void {
    const overlappingRequests = existingRequests.filter(existing => {
      // Skip cancelled or rejected requests
      if (existing.ApprovalStatus === 'Cancelled' || existing.ApprovalStatus === 'Rejected') {
        return false;
      }

      // Check for date overlap
      return (
        (request.StartDate <= existing.EndDate && request.EndDate >= existing.StartDate)
      );
    });

    if (overlappingRequests.length > 0) {
      errors.push(
        `You have overlapping leave requests. Please check your existing requests and modify dates if needed.`
      );
    }
  }

  /**
   * Detect conflicts with team members and blackout dates
   */
  public static async detectConflicts(
    request: ILeaveRequestCreate,
    teamMembers: any[],
    blackoutDates: Date[],
    holidays: Date[]
  ): Promise<IConflictResult> {
    const conflicts: IConflictDetail[] = [];

    // Check blackout dates
    this.checkBlackoutDates(request, blackoutDates, conflicts);

    // Check holidays
    this.checkHolidays(request, holidays, conflicts);

    // Check team member conflicts
    this.checkTeamConflicts(request, teamMembers, conflicts);

    return {
      hasConflicts: conflicts.some(c => c.severity === 'error'),
      conflicts
    };
  }

  /**
   * Check against blackout dates
   */
  private static checkBlackoutDates(
    request: ILeaveRequestCreate,
    blackoutDates: Date[],
    conflicts: IConflictDetail[]
  ): void {
    const requestDates = this.getDateRange(request.StartDate, request.EndDate);
    
    const conflictingDates = requestDates.filter(date => 
      blackoutDates.some(blackout => 
        date.toDateString() === blackout.toDateString()
      )
    );

    if (conflictingDates.length > 0) {
      conflicts.push({
        type: 'blackout-date',
        severity: 'error',
        message: `Your request conflicts with company blackout dates: ${
          conflictingDates.map(d => d.toLocaleDateString()).join(', ')
        }`,
        details: { conflictingDates }
      });
    }
  }

  /**
   * Check against holidays
   */
  private static checkHolidays(
    request: ILeaveRequestCreate,
    holidays: Date[],
    conflicts: IConflictDetail[]
  ): void {
    const requestDates = this.getDateRange(request.StartDate, request.EndDate);
    
    const holidayConflicts = requestDates.filter(date => 
      holidays.some(holiday => 
        date.toDateString() === holiday.toDateString()
      )
    );

    if (holidayConflicts.length > 0) {
      conflicts.push({
        type: 'holiday',
        severity: 'warning',
        message: `Your request includes company holidays: ${
          holidayConflicts.map(d => d.toLocaleDateString()).join(', ')
        }. These days may not count against your leave balance.`,
        details: { holidayConflicts }
      });
    }
  }

  /**
   * Check team member conflicts
   */
  private static checkTeamConflicts(
    request: ILeaveRequestCreate,
    teamMembers: any[],
    conflicts: IConflictDetail[]
  ): void {
    const conflictingMembers = teamMembers.filter(member => {
      return member.leaveRequests?.some((leave: any) => 
        leave.ApprovalStatus === 'Approved' &&
        request.StartDate <= leave.EndDate &&
        request.EndDate >= leave.StartDate
      );
    });

    if (conflictingMembers.length > 0) {
      const memberNames = conflictingMembers.map(m => m.displayName).join(', ');
      conflicts.push({
        type: 'team-member',
        severity: 'warning',
        message: `Team members with overlapping leave: ${memberNames}. ` +
                `Please coordinate with your team to ensure adequate coverage.`,
        details: { conflictingMembers }
      });
    }
  }

  /**
   * Get array of dates between start and end date
   */
  private static getDateRange(startDate: Date, endDate: Date): Date[] {
    const dates: Date[] = [];
    const current = new Date(startDate.getTime());
    
    while (current <= endDate) {
      dates.push(new Date(current.getTime()));
      current.setDate(current.getDate() + 1);
    }
    
    return dates;
  }

  /**
   * Validate bulk operations
   */
  public static validateBulkOperation(
    requests: ILeaveRequestCreate[],
    maxBulkSize: number = 10
  ): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (requests.length === 0) {
      errors.push('No requests provided for bulk operation');
    }

    if (requests.length > maxBulkSize) {
      errors.push(`Bulk operation limited to ${maxBulkSize} requests. Provided: ${requests.length}`);
    }

    // Check for duplicate requests
    const duplicates = this.findDuplicateRequests(requests);
    if (duplicates.length > 0) {
      warnings.push(`Found ${duplicates.length} potential duplicate requests`);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Find duplicate requests in bulk operation
   */
  private static findDuplicateRequests(requests: ILeaveRequestCreate[]): number[] {
    const duplicateIndices: number[] = [];
    
    for (let i = 0; i < requests.length; i++) {
      for (let j = i + 1; j < requests.length; j++) {
        if (
          requests[i].LeaveTypeId === requests[j].LeaveTypeId &&
          requests[i].StartDate.getTime() === requests[j].StartDate.getTime() &&
          requests[i].EndDate.getTime() === requests[j].EndDate.getTime()
        ) {
          duplicateIndices.push(j);
        }
      }
    }
    
    return duplicateIndices;
  }
}
