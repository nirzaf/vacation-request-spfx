/**
 * Interface representing a Leave Balance item from SharePoint
 */
export interface ILeaveBalance {
  Id: number;
  Employee: {
    Id: number;
    Title: string;
    EMail: string;
  };
  LeaveType: {
    Id: number;
    Title: string;
  };
  TotalAllowance: number;
  UsedDays: number;
  RemainingDays: number;
  CarryOverDays: number;
  EffectiveDate: Date;
  ExpirationDate: Date;
}

/**
 * Interface for creating a new Leave Balance
 */
export interface ILeaveBalanceCreate {
  EmployeeId: number;
  LeaveTypeId: number;
  TotalAllowance: number;
  CarryOverDays?: number;
  EffectiveDate: Date;
  ExpirationDate: Date;
}

/**
 * Interface for updating a Leave Balance
 */
export interface ILeaveBalanceUpdate {
  TotalAllowance?: number;
  UsedDays?: number;
  CarryOverDays?: number;
  EffectiveDate?: Date;
  ExpirationDate?: Date;
}

/**
 * Interface for Leave Balance summary
 */
export interface ILeaveBalanceSummary {
  employeeId: number;
  employeeName: string;
  balances: ILeaveBalanceDetail[];
  totalAllowance: number;
  totalUsed: number;
  totalRemaining: number;
}

/**
 * Interface for detailed Leave Balance information
 */
export interface ILeaveBalanceDetail {
  leaveTypeId: number;
  leaveTypeName: string;
  leaveTypeColor?: string;
  totalAllowance: number;
  usedDays: number;
  remainingDays: number;
  carryOverDays: number;
  effectiveDate: Date;
  expirationDate: Date;
  isExpiringSoon: boolean;
  daysUntilExpiry: number;
}

/**
 * Interface for Leave Balance calculation
 */
export interface ILeaveBalanceCalculation {
  leaveBalanceId: number;
  previousUsedDays: number;
  newUsedDays: number;
  previousRemainingDays: number;
  newRemainingDays: number;
  calculationDate: Date;
}

/**
 * Interface for Leave Balance filter options
 */
export interface ILeaveBalanceFilter {
  employeeId?: number;
  leaveTypeId?: number;
  isExpiringSoon?: boolean;
  hasRemainingDays?: boolean;
  effectiveDateFrom?: Date;
  effectiveDateTo?: Date;
  expirationDateFrom?: Date;
  expirationDateTo?: Date;
}

/**
 * Interface for Leave Balance analytics
 */
export interface ILeaveBalanceAnalytics {
  totalEmployees: number;
  totalLeaveTypes: number;
  averageAllowancePerEmployee: number;
  averageUsagePercentage: number;
  balancesExpiringSoon: number;
  topLeaveTypesByUsage: ILeaveTypeUsage[];
}

/**
 * Interface for Leave Type usage statistics
 */
export interface ILeaveTypeUsage {
  leaveTypeId: number;
  leaveTypeName: string;
  totalAllowance: number;
  totalUsed: number;
  usagePercentage: number;
  employeeCount: number;
}

/**
 * Utility functions for Leave Balances
 */
export class LeaveBalanceUtils {
  /**
   * Calculate remaining days
   */
  public static calculateRemainingDays(totalAllowance: number, usedDays: number, carryOverDays: number = 0): number {
    return Math.max(0, totalAllowance + carryOverDays - usedDays);
  }

  /**
   * Check if balance is expiring soon (within 30 days)
   */
  public static isExpiringSoon(expirationDate: Date, daysThreshold: number = 30): boolean {
    const today = new Date();
    const diffTime = expirationDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays <= daysThreshold && diffDays >= 0;
  }

  /**
   * Calculate days until expiry
   */
  public static getDaysUntilExpiry(expirationDate: Date): number {
    const today = new Date();
    const diffTime = expirationDate.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  }

  /**
   * Calculate usage percentage
   */
  public static calculateUsagePercentage(usedDays: number, totalAllowance: number, carryOverDays: number = 0): number {
    const totalAvailable = totalAllowance + carryOverDays;
    return totalAvailable > 0 ? Math.round((usedDays / totalAvailable) * 100) : 0;
  }

  /**
   * Group balances by employee
   */
  public static groupByEmployee(balances: ILeaveBalance[]): Map<number, ILeaveBalance[]> {
    const grouped = new Map<number, ILeaveBalance[]>();
    
    balances.forEach(balance => {
      const employeeId = balance.Employee.Id;
      if (!grouped.has(employeeId)) {
        grouped.set(employeeId, []);
      }
      grouped.get(employeeId)!.push(balance);
    });

    return grouped;
  }

  /**
   * Create Leave Balance summary for an employee
   */
  public static createEmployeeSummary(employeeBalances: ILeaveBalance[]): ILeaveBalanceSummary | null {
    if (employeeBalances.length === 0) return null;

    const employee = employeeBalances[0].Employee;
    const balances: ILeaveBalanceDetail[] = employeeBalances.map(balance => ({
      leaveTypeId: balance.LeaveType.Id,
      leaveTypeName: balance.LeaveType.Title,
      totalAllowance: balance.TotalAllowance,
      usedDays: balance.UsedDays,
      remainingDays: balance.RemainingDays,
      carryOverDays: balance.CarryOverDays,
      effectiveDate: balance.EffectiveDate,
      expirationDate: balance.ExpirationDate,
      isExpiringSoon: this.isExpiringSoon(balance.ExpirationDate),
      daysUntilExpiry: this.getDaysUntilExpiry(balance.ExpirationDate)
    }));

    return {
      employeeId: employee.Id,
      employeeName: employee.Title,
      balances,
      totalAllowance: balances.reduce((sum, b) => sum + b.totalAllowance, 0),
      totalUsed: balances.reduce((sum, b) => sum + b.usedDays, 0),
      totalRemaining: balances.reduce((sum, b) => sum + b.remainingDays, 0)
    };
  }

  /**
   * Validate leave balance data
   */
  public static validateBalance(balance: ILeaveBalanceCreate | ILeaveBalanceUpdate): string[] {
    const errors: string[] = [];

    if ('TotalAllowance' in balance && balance.TotalAllowance !== undefined) {
      if (balance.TotalAllowance < 0) {
        errors.push('Total allowance cannot be negative');
      }
      if (balance.TotalAllowance > 365) {
        errors.push('Total allowance cannot exceed 365 days');
      }
    }

    if ('CarryOverDays' in balance && balance.CarryOverDays !== undefined) {
      if (balance.CarryOverDays < 0) {
        errors.push('Carry over days cannot be negative');
      }
      if (balance.CarryOverDays > 30) {
        errors.push('Carry over days cannot exceed 30 days');
      }
    }

    if ('EffectiveDate' in balance && 'ExpirationDate' in balance && 
        balance.EffectiveDate && balance.ExpirationDate) {
      if (balance.ExpirationDate <= balance.EffectiveDate) {
        errors.push('Expiration date must be after effective date');
      }
    }

    return errors;
  }
}
