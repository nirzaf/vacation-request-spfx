import * as React from 'react';
import styles from './TeamCalendar.module.scss';
import type { ITeamCalendarProps } from './ITeamCalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import listPlugin from '@fullcalendar/list';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  DefaultButton,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Panel,
  PanelType,
  TextField,
  Toggle
} from '@fluentui/react';
import { SharePointService, GraphService } from '../../../services';
import {
  ILeaveRequest,
  ILeaveType,
  LeaveTypeUtils,
  CommonUtils
} from '../../../models';

interface ICalendarEvent {
  id: string;
  title: string;
  start: string;
  end: string;
  allDay: boolean;
  backgroundColor: string;
  borderColor: string;
  textColor: string;
  extendedProps: {
    leaveRequest: ILeaveRequest;
    leaveType: ILeaveType;
  };
}

interface ITeamCalendarState {
  events: ICalendarEvent[];
  leaveTypes: ILeaveType[];
  isLoading: boolean;
  error: string;
  currentView: string;
  selectedEvent: ICalendarEvent | undefined;
  isPanelOpen: boolean;
  filters: {
    leaveTypeId: number | undefined;
    department: string;
    showOnlyApproved: boolean;
  };
}

export default class TeamCalendar extends React.Component<ITeamCalendarProps, ITeamCalendarState> {
  private sharePointService: SharePointService;
  private graphService: GraphService;
  private calendarRef = React.createRef<FullCalendar>();

  constructor(props: ITeamCalendarProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);
    this.graphService = new GraphService(props.context);

    this.state = {
      events: [],
      leaveTypes: [],
      isLoading: true,
      error: '',
      currentView: 'dayGridMonth',
      selectedEvent: undefined,
      isPanelOpen: false,
      filters: {
        leaveTypeId: undefined,
        department: '',
        showOnlyApproved: true
      }
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      this.setState({ isLoading: true, error: '' });

      const [leaveTypes, leaveRequests] = await Promise.all([
        this.sharePointService.getLeaveTypes(),
        this.sharePointService.getAllLeaveRequests()
      ]);

      const events = this.convertToCalendarEvents(leaveRequests, leaveTypes);

      this.setState({
        leaveTypes,
        events,
        isLoading: false
      });
    } catch (error) {
      console.error('Error loading calendar data:', error);
      this.setState({
        error: 'Failed to load calendar data. Please refresh the page.',
        isLoading: false
      });
    }
  }

  private convertToCalendarEvents(
    leaveRequests: ILeaveRequest[],
    leaveTypes: ILeaveType[]
  ): ICalendarEvent[] {
    return leaveRequests
      .filter(request => this.shouldShowRequest(request))
      .map(request => {
        const leaveType = leaveTypes.filter((lt: ILeaveType) => lt.Id === request.LeaveType.Id)[0];
        const color = leaveType?.ColorCode || '#0078d4';

        return {
          id: request.Id.toString(),
          title: `${request.Requester.Title} - ${request.LeaveType.Title}`,
          start: request.StartDate.toISOString().split('T')[0],
          end: this.getEndDateForCalendar(request.EndDate),
          allDay: !request.IsPartialDay,
          backgroundColor: color,
          borderColor: color,
          textColor: this.getTextColor(color),
          extendedProps: {
            leaveRequest: request,
            leaveType: leaveType || { Id: 0, Title: 'Unknown', IsActive: true, RequiresApproval: true, RequiresDocumentation: false, Created: new Date(), Modified: new Date() }
          }
        };
      });
  }

  private shouldShowRequest(request: ILeaveRequest): boolean {
    const { filters } = this.state;

    // Filter by approval status
    if (filters.showOnlyApproved && request.ApprovalStatus !== 'Approved') {
      return false;
    }

    // Filter by leave type
    if (filters.leaveTypeId && request.LeaveType.Id !== filters.leaveTypeId) {
      return false;
    }

    // Filter by department
    if (filters.department && request.Department !== filters.department) {
      return false;
    }

    return true;
  }

  private getEndDateForCalendar(endDate: Date): string {
    // FullCalendar expects end date to be the day after for all-day events
    const nextDay = new Date(endDate.getTime());
    nextDay.setDate(nextDay.getDate() + 1);
    return nextDay.toISOString().split('T')[0];
  }

  private getTextColor(backgroundColor: string): string {
    // Simple contrast calculation
    const hex = backgroundColor.replace('#', '');
    const r = parseInt(hex.substr(0, 2), 16);
    const g = parseInt(hex.substr(2, 2), 16);
    const b = parseInt(hex.substr(4, 2), 16);
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    return brightness > 128 ? '#000000' : '#ffffff';
  }

  private onEventClick = (info: any): void => {
    const event = info.event;
    const calendarEvent: ICalendarEvent = {
      id: event.id,
      title: event.title,
      start: event.startStr,
      end: event.endStr,
      allDay: event.allDay,
      backgroundColor: event.backgroundColor,
      borderColor: event.borderColor,
      textColor: event.textColor,
      extendedProps: event.extendedProps
    };

    this.setState({
      selectedEvent: calendarEvent,
      isPanelOpen: true
    });
  };

  private onViewChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option && this.calendarRef.current) {
      const calendarApi = this.calendarRef.current.getApi();
      calendarApi.changeView(option.key as string);
      this.setState({ currentView: option.key as string });
    }
  };

  private onFilterChange = async (filterType: string, value: any): Promise<void> => {
    const newFilters = { ...this.state.filters, [filterType]: value };
    this.setState({ filters: newFilters });

    // Reload events with new filters
    await this.loadData();
  };

  private onClosePanel = (): void => {
    this.setState({
      isPanelOpen: false,
      selectedEvent: undefined
    });
  };

  private onRefresh = async (): Promise<void> => {
    await this.loadData();
  };

  private onExportCalendar = (): void => {
    // Implementation for exporting calendar data
    const { events } = this.state;
    const csvContent = this.convertEventsToCSV(events);
    this.downloadCSV(csvContent, 'team-calendar.csv');
  };

  private convertEventsToCSV(events: ICalendarEvent[]): string {
    const headers = ['Employee', 'Leave Type', 'Start Date', 'End Date', 'Status', 'Comments'];
    const rows = events.map(event => [
      event.extendedProps.leaveRequest.Requester.Title,
      event.extendedProps.leaveRequest.LeaveType.Title,
      event.extendedProps.leaveRequest.StartDate.toLocaleDateString(),
      event.extendedProps.leaveRequest.EndDate.toLocaleDateString(),
      event.extendedProps.leaveRequest.ApprovalStatus,
      event.extendedProps.leaveRequest.RequestComments || ''
    ]);

    return [headers, ...rows].map(row => row.join(',')).join('\n');
  }

  private downloadCSV(content: string, filename: string): void {
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
