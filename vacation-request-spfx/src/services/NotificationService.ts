import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphService } from './GraphService';
import { ILeaveRequest, ILeaveType } from '../models';

/**
 * Interface for notification template data
 */
export interface INotificationTemplate {
  subject: string;
  body: string;
  isHtml: boolean;
}

/**
 * Service class for handling notifications
 */
export class NotificationService {
  private context: WebPartContext;
  private graphService: GraphService;

  constructor(context: WebPartContext) {
    this.context = context;
    this.graphService = new GraphService(context);
  }

  /**
   * Send leave request submission notification
   */
  public async sendSubmissionNotification(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType
  ): Promise<void> {
    try {
      const template = this.getSubmissionTemplate(leaveRequest, leaveType);
      const recipients = [leaveRequest.Requester.EMail];

      // Add manager email if available
      if (leaveRequest.Manager?.EMail) {
        recipients.push(leaveRequest.Manager.EMail);
      }

      await this.graphService.sendNotificationEmail(
        recipients,
        template.subject,
        template.body,
        template.isHtml
      );
    } catch (error) {
      console.error('Error sending submission notification:', error);
      throw new Error('Failed to send submission notification');
    }
  }

  /**
   * Send approval notification
   */
  public async sendApprovalNotification(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType,
    isApproved: boolean,
    approverComments?: string
  ): Promise<void> {
    try {
      const template = this.getApprovalTemplate(
        leaveRequest,
        leaveType,
        isApproved,
        approverComments
      );
      
      const recipients = [leaveRequest.Requester.EMail];

      await this.graphService.sendNotificationEmail(
        recipients,
        template.subject,
        template.body,
        template.isHtml
      );
    } catch (error) {
      console.error('Error sending approval notification:', error);
      throw new Error('Failed to send approval notification');
    }
  }

  /**
   * Send reminder notification for pending approvals
   */
  public async sendPendingApprovalReminder(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType
  ): Promise<void> {
    try {
      if (!leaveRequest.Manager?.EMail) {
        console.warn('No manager email available for reminder');
        return;
      }

      const template = this.getPendingApprovalReminderTemplate(leaveRequest, leaveType);
      
      await this.graphService.sendNotificationEmail(
        [leaveRequest.Manager.EMail],
        template.subject,
        template.body,
        template.isHtml
      );
    } catch (error) {
      console.error('Error sending pending approval reminder:', error);
      throw new Error('Failed to send pending approval reminder');
    }
  }

  /**
   * Send leave balance warning notification
   */
  public async sendBalanceWarningNotification(
    userEmail: string,
    leaveTypeName: string,
    remainingDays: number,
    expirationDate: Date
  ): Promise<void> {
    try {
      const template = this.getBalanceWarningTemplate(
        leaveTypeName,
        remainingDays,
        expirationDate
      );
      
      await this.graphService.sendNotificationEmail(
        [userEmail],
        template.subject,
        template.body,
        template.isHtml
      );
    } catch (error) {
      console.error('Error sending balance warning notification:', error);
      throw new Error('Failed to send balance warning notification');
    }
  }

  /**
   * Get submission notification template
   */
  private getSubmissionTemplate(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType
  ): INotificationTemplate {
    const startDate = leaveRequest.StartDate.toLocaleDateString();
    const endDate = leaveRequest.EndDate.toLocaleDateString();
    const totalDays = leaveRequest.TotalDays || 1;

    const subject = `Leave Request Submitted - ${leaveType.Title}`;
    
    const body = `
Dear ${leaveRequest.Requester.Title},

Your leave request has been successfully submitted and is now pending approval.

Request Details:
- Leave Type: ${leaveType.Title}
- Start Date: ${startDate}
- End Date: ${endDate}
- Total Days: ${totalDays}
${leaveRequest.IsPartialDay ? `- Partial Day Hours: ${leaveRequest.PartialDayHours}` : ''}
${leaveRequest.RequestComments ? `- Comments: ${leaveRequest.RequestComments}` : ''}

${leaveType.RequiresApproval ? 
  `Your request will be reviewed by your manager and you will be notified of the decision.` :
  `This leave type does not require approval and has been automatically approved.`
}

You can track the status of your request in the Leave Request system.

Best regards,
HR Team
    `.trim();

    return {
      subject,
      body,
      isHtml: false
    };
  }

  /**
   * Get approval notification template
   */
  private getApprovalTemplate(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType,
    isApproved: boolean,
    approverComments?: string
  ): INotificationTemplate {
    const startDate = leaveRequest.StartDate.toLocaleDateString();
    const endDate = leaveRequest.EndDate.toLocaleDateString();
    const status = isApproved ? 'APPROVED' : 'REJECTED';
    const statusText = isApproved ? 'approved' : 'rejected';

    const subject = `Leave Request ${status} - ${leaveType.Title}`;
    
    const body = `
Dear ${leaveRequest.Requester.Title},

Your leave request has been ${statusText}.

Request Details:
- Leave Type: ${leaveType.Title}
- Start Date: ${startDate}
- End Date: ${endDate}
- Total Days: ${leaveRequest.TotalDays || 1}
- Status: ${status}
${approverComments ? `- Manager Comments: ${approverComments}` : ''}

${isApproved ? 
  `Your leave has been approved and a calendar event has been created. Please ensure proper handover of your responsibilities before your leave begins.` :
  `If you have questions about this decision, please contact your manager or HR.`
}

Best regards,
HR Team
    `.trim();

    return {
      subject,
      body,
      isHtml: false
    };
  }

  /**
   * Get pending approval reminder template
   */
  private getPendingApprovalReminderTemplate(
    leaveRequest: ILeaveRequest,
    leaveType: ILeaveType
  ): INotificationTemplate {
    const startDate = leaveRequest.StartDate.toLocaleDateString();
    const endDate = leaveRequest.EndDate.toLocaleDateString();
    const submissionDate = leaveRequest.SubmissionDate.toLocaleDateString();

    const subject = `Reminder: Pending Leave Request Approval - ${leaveRequest.Requester.Title}`;
    
    const body = `
Dear ${leaveRequest.Manager?.Title || 'Manager'},

This is a reminder that you have a pending leave request that requires your approval.

Request Details:
- Employee: ${leaveRequest.Requester.Title}
- Leave Type: ${leaveType.Title}
- Start Date: ${startDate}
- End Date: ${endDate}
- Total Days: ${leaveRequest.TotalDays || 1}
- Submitted: ${submissionDate}
${leaveRequest.RequestComments ? `- Comments: ${leaveRequest.RequestComments}` : ''}

Please review and approve or reject this request at your earliest convenience.

You can access the Leave Request system to take action on this request.

Best regards,
HR Team
    `.trim();

    return {
      subject,
      body,
      isHtml: false
    };
  }

  /**
   * Get balance warning notification template
   */
  private getBalanceWarningTemplate(
    leaveTypeName: string,
    remainingDays: number,
    expirationDate: Date
  ): INotificationTemplate {
    const expirationDateStr = expirationDate.toLocaleDateString();

    const subject = `Leave Balance Expiration Warning - ${leaveTypeName}`;
    
    const body = `
Dear Team Member,

This is a reminder that you have unused leave balance that will expire soon.

Leave Balance Details:
- Leave Type: ${leaveTypeName}
- Remaining Days: ${remainingDays}
- Expiration Date: ${expirationDateStr}

Please consider using your remaining leave days before they expire. You can submit a leave request through the Leave Request system.

If you have any questions about your leave balance or policies, please contact HR.

Best regards,
HR Team
    `.trim();

    return {
      subject,
      body,
      isHtml: false
    };
  }

  /**
   * Create notification for Teams (if available)
   */
  public async sendTeamsNotification(
    message: string,
    title?: string
  ): Promise<void> {
    try {
      // Check if running in Teams context
      if (this.context.sdks.microsoftTeams) {
        // This would require additional Teams SDK integration
        console.log('Teams notification:', { title, message });
        // Implementation would depend on Teams notification capabilities
      }
    } catch (error) {
      console.error('Error sending Teams notification:', error);
    }
  }

  /**
   * Batch send notifications for multiple requests
   */
  public async sendBatchNotifications(
    notifications: Array<{
      type: 'submission' | 'approval' | 'reminder' | 'balance-warning';
      data: any;
    }>
  ): Promise<void> {
    const promises = notifications.map(async (notification) => {
      try {
        switch (notification.type) {
          case 'submission':
            await this.sendSubmissionNotification(
              notification.data.leaveRequest,
              notification.data.leaveType
            );
            break;
          case 'approval':
            await this.sendApprovalNotification(
              notification.data.leaveRequest,
              notification.data.leaveType,
              notification.data.isApproved,
              notification.data.approverComments
            );
            break;
          case 'reminder':
            await this.sendPendingApprovalReminder(
              notification.data.leaveRequest,
              notification.data.leaveType
            );
            break;
          case 'balance-warning':
            await this.sendBalanceWarningNotification(
              notification.data.userEmail,
              notification.data.leaveTypeName,
              notification.data.remainingDays,
              notification.data.expirationDate
            );
            break;
        }
      } catch (error) {
        console.error(`Error sending ${notification.type} notification:`, error);
      }
    });

    // Wait for all promises to complete (some may fail)
    for (const promise of promises) {
      try {
        await promise;
      } catch (error) {
        // Individual errors are already logged in the promise handlers
      }
    }
  }
}
