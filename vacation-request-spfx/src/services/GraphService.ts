import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Interface for calendar event
 */
export interface ICalendarEvent {
  id?: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  isAllDay: boolean;
  showAs: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere';
  categories: string[];
  body?: {
    contentType: 'text' | 'html';
    content: string;
  };
}

/**
 * Interface for user profile information
 */
export interface IUserProfile {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  manager?: {
    id: string;
    displayName: string;
    mail: string;
  };
}

/**
 * Service class for Microsoft Graph operations
 */
export class GraphService {
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | undefined;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Initialize Graph client
   */
  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  /**
   * Get current user profile with manager information
   */
  public async getCurrentUserProfile(): Promise<IUserProfile> {
    try {
      const client = await this.getGraphClient();
      
      // Get user profile
      const userResponse = await client
        .api('/me')
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
        .get();

      // Get manager information
      let manager;
      try {
        const managerResponse = await client
          .api('/me/manager')
          .select('id,displayName,mail')
          .get();
        manager = {
          id: managerResponse.id,
          displayName: managerResponse.displayName,
          mail: managerResponse.mail
        };
      } catch (error) {
        console.warn('Manager information not available:', error);
        manager = undefined;
      }

      return {
        id: userResponse.id,
        displayName: userResponse.displayName,
        mail: userResponse.mail,
        userPrincipalName: userResponse.userPrincipalName,
        jobTitle: userResponse.jobTitle,
        department: userResponse.department,
        manager
      };
    } catch (error) {
      console.error('Error fetching user profile:', error);
      throw new Error('Failed to fetch user profile');
    }
  }

  /**
   * Create calendar event for approved leave
   */
  public async createCalendarEvent(event: ICalendarEvent): Promise<string> {
    try {
      const client = await this.getGraphClient();
      
      const response = await client
        .api('/me/events')
        .post(event);

      return response.id;
    } catch (error) {
      console.error('Error creating calendar event:', error);
      throw new Error('Failed to create calendar event');
    }
  }

  /**
   * Update existing calendar event
   */
  public async updateCalendarEvent(eventId: string, event: Partial<ICalendarEvent>): Promise<void> {
    try {
      const client = await this.getGraphClient();
      
      await client
        .api(`/me/events/${eventId}`)
        .patch(event);
    } catch (error) {
      console.error('Error updating calendar event:', error);
      throw new Error('Failed to update calendar event');
    }
  }

  /**
   * Delete calendar event
   */
  public async deleteCalendarEvent(eventId: string): Promise<void> {
    try {
      const client = await this.getGraphClient();
      
      await client
        .api(`/me/events/${eventId}`)
        .delete();
    } catch (error) {
      console.error('Error deleting calendar event:', error);
      throw new Error('Failed to delete calendar event');
    }
  }

  /**
   * Get team calendar events for conflict detection
   */
  public async getTeamCalendarEvents(startDate: Date, endDate: Date, userIds: string[]): Promise<ICalendarEvent[]> {
    try {
      const client = await this.getGraphClient();
      const events: ICalendarEvent[] = [];

      // Get calendar view for each team member
      for (const userId of userIds) {
        try {
          const response = await client
            .api(`/users/${userId}/calendarView`)
            .query({
              startDateTime: startDate.toISOString(),
              endDateTime: endDate.toISOString()
            })
            .select('id,subject,start,end,isAllDay,showAs,categories')
            .get();

          events.push(...response.value);
        } catch (error) {
          console.warn(`Could not fetch calendar for user ${userId}:`, error);
        }
      }

      return events;
    } catch (error) {
      console.error('Error fetching team calendar events:', error);
      throw new Error('Failed to fetch team calendar events');
    }
  }

  /**
   * Send notification email
   */
  public async sendNotificationEmail(
    to: string[], 
    subject: string, 
    body: string, 
    isHtml: boolean = false
  ): Promise<void> {
    try {
      const client = await this.getGraphClient();
      
      const message = {
        subject,
        body: {
          contentType: isHtml ? 'html' : 'text',
          content: body
        },
        toRecipients: to.map(email => ({
          emailAddress: {
            address: email
          }
        }))
      };

      await client
        .api('/me/sendMail')
        .post({
          message,
          saveToSentItems: true
        });
    } catch (error) {
      console.error('Error sending notification email:', error);
      throw new Error('Failed to send notification email');
    }
  }

  /**
   * Get user's direct reports (for managers)
   */
  public async getDirectReports(): Promise<IUserProfile[]> {
    try {
      const client = await this.getGraphClient();
      
      const response = await client
        .api('/me/directReports')
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
        .get();

      return response.value.map((user: any) => ({
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        userPrincipalName: user.userPrincipalName,
        jobTitle: user.jobTitle,
        department: user.department
      }));
    } catch (error) {
      console.error('Error fetching direct reports:', error);
      throw new Error('Failed to fetch direct reports');
    }
  }

  /**
   * Check if user has manager permissions
   */
  public async hasManagerPermissions(): Promise<boolean> {
    try {
      const directReports = await this.getDirectReports();
      return directReports.length > 0;
    } catch (error) {
      console.warn('Could not check manager permissions:', error);
      return false;
    }
  }

  /**
   * Create leave request calendar event
   */
  public async createLeaveCalendarEvent(
    leaveTypeName: string,
    startDate: Date,
    endDate: Date,
    isPartialDay: boolean,
    partialDayHours?: number,
    comments?: string
  ): Promise<string> {
    const subject = `${leaveTypeName} - Out of Office`;
    const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    
    let body = `Leave Type: ${leaveTypeName}`;
    if (isPartialDay && partialDayHours) {
      body += `\nPartial Day: ${partialDayHours} hours`;
    }
    if (comments) {
      body += `\nComments: ${comments}`;
    }

    const event: ICalendarEvent = {
      subject,
      start: {
        dateTime: startDate.toISOString(),
        timeZone
      },
      end: {
        dateTime: endDate.toISOString(),
        timeZone
      },
      isAllDay: !isPartialDay,
      showAs: 'oof',
      categories: ['Leave Request'],
      body: {
        contentType: 'text',
        content: body
      }
    };

    return await this.createCalendarEvent(event);
  }
}
