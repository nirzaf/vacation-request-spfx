/**
 * Interface representing a Leave Type item from SharePoint
 */
export interface ILeaveType {
  Id: number;
  Title: string;
  Description?: string;
  IsActive: boolean;
  RequiresApproval: boolean;
  MaxDaysPerRequest?: number;
  RequiresDocumentation: boolean;
  ColorCode?: string;
  PolicyURL?: string;
  Created: Date;
  Modified: Date;
}

/**
 * Interface for creating a new Leave Type
 */
export interface ILeaveTypeCreate {
  Title: string;
  Description?: string;
  IsActive: boolean;
  RequiresApproval: boolean;
  MaxDaysPerRequest?: number;
  RequiresDocumentation: boolean;
  ColorCode?: string;
  PolicyURL?: string;
}

/**
 * Interface for updating a Leave Type
 */
export interface ILeaveTypeUpdate {
  Title?: string;
  Description?: string;
  IsActive?: boolean;
  RequiresApproval?: boolean;
  MaxDaysPerRequest?: number;
  RequiresDocumentation?: boolean;
  ColorCode?: string;
  PolicyURL?: string;
}

/**
 * Interface for Leave Type dropdown options
 */
export interface ILeaveTypeOption {
  key: number;
  text: string;
  data?: ILeaveType;
}

/**
 * Interface for Leave Type validation rules
 */
export interface ILeaveTypeValidationRules {
  maxDaysPerRequest?: number;
  requiresApproval: boolean;
  requiresDocumentation: boolean;
  isActive: boolean;
}

/**
 * Interface for Leave Type statistics
 */
export interface ILeaveTypeStats {
  leaveTypeId: number;
  leaveTypeName: string;
  totalRequests: number;
  totalDaysRequested: number;
  averageDaysPerRequest: number;
  mostRequestedMonth: string;
}

/**
 * Default Leave Types for initial setup
 */
export const DEFAULT_LEAVE_TYPES: ILeaveTypeCreate[] = [
  {
    Title: 'Annual Leave',
    Description: 'Standard annual vacation leave',
    IsActive: true,
    RequiresApproval: true,
    MaxDaysPerRequest: 30,
    RequiresDocumentation: false,
    ColorCode: '#4CAF50'
  },
  {
    Title: 'Sick Leave',
    Description: 'Medical leave for illness',
    IsActive: true,
    RequiresApproval: false,
    MaxDaysPerRequest: 5,
    RequiresDocumentation: true,
    ColorCode: '#FF9800'
  },
  {
    Title: 'Personal Leave',
    Description: 'Personal time off',
    IsActive: true,
    RequiresApproval: true,
    MaxDaysPerRequest: 10,
    RequiresDocumentation: false,
    ColorCode: '#2196F3'
  },
  {
    Title: 'Maternity/Paternity Leave',
    Description: 'Family leave for new parents',
    IsActive: true,
    RequiresApproval: true,
    MaxDaysPerRequest: 90,
    RequiresDocumentation: true,
    ColorCode: '#E91E63'
  },
  {
    Title: 'Emergency Leave',
    Description: 'Urgent personal matters',
    IsActive: true,
    RequiresApproval: false,
    MaxDaysPerRequest: 3,
    RequiresDocumentation: false,
    ColorCode: '#F44336'
  }
];

/**
 * Utility functions for Leave Types
 */
export class LeaveTypeUtils {
  /**
   * Get active leave types only
   */
  public static getActiveLeaveTypes(leaveTypes: ILeaveType[]): ILeaveType[] {
    return leaveTypes.filter(lt => lt.IsActive);
  }

  /**
   * Convert leave types to dropdown options
   */
  public static toDropdownOptions(leaveTypes: ILeaveType[]): ILeaveTypeOption[] {
    return leaveTypes
      .filter(lt => lt.IsActive)
      .map(lt => ({
        key: lt.Id,
        text: lt.Title,
        data: lt
      }));
  }

  /**
   * Get leave type by ID
   */
  public static getLeaveTypeById(leaveTypes: ILeaveType[], id: number): ILeaveType | undefined {
    return leaveTypes.filter((lt: ILeaveType) => lt.Id === id)[0];
  }

  /**
   * Validate color code format
   */
  public static isValidColorCode(colorCode: string): boolean {
    const colorRegex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
    return colorRegex.test(colorCode);
  }

  /**
   * Get default color if none provided
   */
  public static getDefaultColor(title: string): string {
    const colors = ['#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0', '#607D8B'];
    const index = title.length % colors.length;
    return colors[index];
  }
}
