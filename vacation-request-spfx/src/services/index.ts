/**
 * Export all services
 */

export * from './SharePointService';
export * from './GraphService';
export * from './NotificationService';
export * from './ValidationService';

/**
 * Service factory for creating service instances
 */
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { GraphService } from './GraphService';
import { NotificationService } from './NotificationService';

export class ServiceFactory {
  /**
   * Create SharePoint service instance
   */
  public static createSharePointService(context: WebPartContext): SharePointService {
    return new SharePointService(context);
  }

  /**
   * Create Graph service instance
   */
  public static createGraphService(context: WebPartContext): GraphService {
    return new GraphService(context);
  }

  /**
   * Create Notification service instance
   */
  public static createNotificationService(context: WebPartContext): NotificationService {
    return new NotificationService(context);
  }
}

/**
 * Service manager for coordinating multiple services
 */
export class ServiceManager {
  private sharePointService: SharePointService;
  private graphService: GraphService;
  private notificationService: NotificationService;

  constructor(context: WebPartContext) {
    this.sharePointService = new SharePointService(context);
    this.graphService = new GraphService(context);
    this.notificationService = new NotificationService(context);
  }

  /**
   * Get SharePoint service
   */
  public getSharePointService(): SharePointService {
    return this.sharePointService;
  }

  /**
   * Get Graph service
   */
  public getGraphService(): GraphService {
    return this.graphService;
  }

  /**
   * Get Notification service
   */
  public getNotificationService(): NotificationService {
    return this.notificationService;
  }

  /**
   * Submit leave request with full workflow
   */
  public async submitLeaveRequestWithWorkflow(
    request: any,
    sendNotifications: boolean = true
  ): Promise<any> {
    try {
      // 1. Create the leave request
      const createdRequest = await this.sharePointService.createLeaveRequest(request);

      // 2. Get leave type details
      const leaveType = await this.sharePointService.getLeaveTypeById(request.LeaveTypeId);
      
      if (!leaveType) {
        throw new Error('Leave type not found');
      }

      // 3. Send notifications if enabled
      if (sendNotifications) {
        await this.notificationService.sendSubmissionNotification(createdRequest, leaveType);
      }

      // 4. Create calendar event if auto-approved
      if (!leaveType.RequiresApproval) {
        try {
          const eventId = await this.graphService.createLeaveCalendarEvent(
            leaveType.Title,
            createdRequest.StartDate,
            createdRequest.EndDate,
            createdRequest.IsPartialDay,
            createdRequest.PartialDayHours,
            createdRequest.RequestComments
          );

          // Update request with calendar event ID
          await this.sharePointService.updateLeaveRequest(createdRequest.Id, {
            CalendarEventID: eventId,
            ApprovalStatus: 'Approved' as any,
            ApprovalDate: new Date()
          });
        } catch (calendarError) {
          console.warn('Could not create calendar event:', calendarError);
        }
      }

      return createdRequest;
    } catch (error) {
      console.error('Error in submitLeaveRequestWithWorkflow:', error);
      throw error;
    }
  }

  /**
   * Approve leave request with full workflow
   */
  public async approveLeaveRequestWithWorkflow(
    requestId: number,
    approverComments?: string,
    sendNotifications: boolean = true
  ): Promise<void> {
    try {
      // 1. Get the leave request
      const leaveRequest = await this.sharePointService.getLeaveRequestById(requestId);
      
      // 2. Get leave type
      const leaveType = await this.sharePointService.getLeaveTypeById(leaveRequest.LeaveType.Id);
      
      if (!leaveType) {
        throw new Error('Leave type not found');
      }

      // 3. Update request status
      await this.sharePointService.updateLeaveRequest(requestId, {
        ApprovalStatus: 'Approved' as any,
        ApprovalDate: new Date(),
        ApprovalComments: approverComments
      });

      // 4. Create calendar event
      try {
        const eventId = await this.graphService.createLeaveCalendarEvent(
          leaveType.Title,
          leaveRequest.StartDate,
          leaveRequest.EndDate,
          leaveRequest.IsPartialDay,
          leaveRequest.PartialDayHours,
          leaveRequest.RequestComments
        );

        // Update request with calendar event ID
        await this.sharePointService.updateLeaveRequest(requestId, {
          CalendarEventID: eventId
        });
      } catch (calendarError) {
        console.warn('Could not create calendar event:', calendarError);
      }

      // 5. Send notifications
      if (sendNotifications) {
        const updatedRequest = await this.sharePointService.getLeaveRequestById(requestId);
        await this.notificationService.sendApprovalNotification(
          updatedRequest,
          leaveType,
          true,
          approverComments
        );
      }
    } catch (error) {
      console.error('Error in approveLeaveRequestWithWorkflow:', error);
      throw error;
    }
  }

  /**
   * Reject leave request with full workflow
   */
  public async rejectLeaveRequestWithWorkflow(
    requestId: number,
    approverComments?: string,
    sendNotifications: boolean = true
  ): Promise<void> {
    try {
      // 1. Get the leave request
      const leaveRequest = await this.sharePointService.getLeaveRequestById(requestId);
      
      // 2. Get leave type
      const leaveType = await this.sharePointService.getLeaveTypeById(leaveRequest.LeaveType.Id);
      
      if (!leaveType) {
        throw new Error('Leave type not found');
      }

      // 3. Update request status
      await this.sharePointService.updateLeaveRequest(requestId, {
        ApprovalStatus: 'Rejected' as any,
        ApprovalDate: new Date(),
        ApprovalComments: approverComments
      });

      // 4. Send notifications
      if (sendNotifications) {
        const updatedRequest = await this.sharePointService.getLeaveRequestById(requestId);
        await this.notificationService.sendApprovalNotification(
          updatedRequest,
          leaveType,
          false,
          approverComments
        );
      }
    } catch (error) {
      console.error('Error in rejectLeaveRequestWithWorkflow:', error);
      throw error;
    }
  }

  /**
   * Cancel leave request with cleanup
   */
  public async cancelLeaveRequestWithCleanup(
    requestId: number,
    sendNotifications: boolean = true
  ): Promise<void> {
    try {
      // 1. Get the leave request
      const leaveRequest = await this.sharePointService.getLeaveRequestById(requestId);
      
      // 2. Delete calendar event if exists
      if (leaveRequest.CalendarEventID) {
        try {
          await this.graphService.deleteCalendarEvent(leaveRequest.CalendarEventID);
        } catch (calendarError) {
          console.warn('Could not delete calendar event:', calendarError);
        }
      }

      // 3. Update request status
      await this.sharePointService.updateLeaveRequest(requestId, {
        ApprovalStatus: 'Cancelled' as any,
        CalendarEventID: undefined
      });

      // 4. Send notifications if needed
      if (sendNotifications) {
        // Implementation for cancellation notifications
        console.log('Leave request cancelled:', requestId);
      }
    } catch (error) {
      console.error('Error in cancelLeaveRequestWithCleanup:', error);
      throw error;
    }
  }

  /**
   * Get comprehensive user dashboard data
   */
  public async getUserDashboardData(userId?: string): Promise<any> {
    try {
      const [
        leaveRequests,
        leaveBalances,
        leaveTypes,
        userProfile
      ] = await Promise.all([
        this.sharePointService.getCurrentUserLeaveRequests(),
        this.sharePointService.getCurrentUserLeaveBalances(),
        this.sharePointService.getLeaveTypes(),
        this.graphService.getCurrentUserProfile()
      ]);

      return {
        leaveRequests,
        leaveBalances,
        leaveTypes,
        userProfile,
        summary: {
          totalRequests: leaveRequests.length,
          pendingRequests: leaveRequests.filter(r => r.ApprovalStatus === 'Pending').length,
          approvedRequests: leaveRequests.filter(r => r.ApprovalStatus === 'Approved').length,
          totalBalance: leaveBalances.reduce((sum, b) => sum + b.RemainingDays, 0)
        }
      };
    } catch (error) {
      console.error('Error getting user dashboard data:', error);
      throw error;
    }
  }
}
