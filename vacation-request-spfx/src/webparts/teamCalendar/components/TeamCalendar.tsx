import * as React from 'react';
import styles from './TeamCalendar.module.scss';
import type { ITeamCalendarProps } from './ITeamCalendarProps';
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
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Panel,
  PanelType,
  TextField,
  Toggle
} from '@fluentui/react';
import { SharePointService } from '../../../services';
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
  private calendarRef = React.createRef<FullCalendar>();

  constructor(props: ITeamCalendarProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);

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

  public render(): React.ReactElement<ITeamCalendarProps> {
    const { hasTeamsContext } = this.props;
    const {
      events,
      leaveTypes,
      isLoading,
      error,
      currentView,
      selectedEvent,
      isPanelOpen,
      filters
    } = this.state;

    if (isLoading) {
      return (
        <section className={`${styles.teamCalendar} ${hasTeamsContext ? styles.teams : ''}`}>
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Loading team calendar..." />
          </Stack>
        </section>
      );
    }

    const viewOptions: IDropdownOption[] = [
      { key: 'dayGridMonth', text: 'Month View' },
      { key: 'timeGridWeek', text: 'Week View' },
      { key: 'timeGridDay', text: 'Day View' },
      { key: 'listWeek', text: 'List View' }
    ];

    const leaveTypeOptions: IDropdownOption[] = [
      { key: 'all', text: 'All Leave Types' },
      ...LeaveTypeUtils.toDropdownOptions(leaveTypes)
    ];

    return (
      <section className={`${styles.teamCalendar} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="xxLarge" as="h1">Team Leave Calendar</Text>
            <Text variant="medium">View team leave requests and plan coverage</Text>
          </Stack.Item>

          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}

          {/* Toolbar */}
          <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
            <Dropdown
              label="View"
              options={viewOptions}
              selectedKey={currentView}
              onChange={this.onViewChange}
              styles={{ root: { minWidth: 120 } }}
            />

            <Dropdown
              label="Leave Type"
              options={leaveTypeOptions}
              selectedKey={filters.leaveTypeId || 'all'}
              onChange={(e, option) => this.onFilterChange('leaveTypeId', option?.key === 'all' ? undefined : option?.key)}
              styles={{ root: { minWidth: 150 } }}
            />

            <TextField
              label="Department"
              value={filters.department}
              onChange={(e, value) => this.onFilterChange('department', value || '')}
              styles={{ root: { minWidth: 120 } }}
            />

            <Toggle
              label="Show only approved"
              checked={filters.showOnlyApproved}
              onChange={(e, checked) => this.onFilterChange('showOnlyApproved', !!checked)}
            />

            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <DefaultButton
                text="Refresh"
                iconProps={{ iconName: 'Refresh' }}
                onClick={this.onRefresh}
              />
              <DefaultButton
                text="Export"
                iconProps={{ iconName: 'Download' }}
                onClick={this.onExportCalendar}
              />
            </Stack>
          </Stack>

          {/* Calendar */}
          <Stack.Item className={styles.calendarContainer}>
            <FullCalendar
              ref={this.calendarRef}
              plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin, listPlugin]}
              initialView={currentView}
              headerToolbar={{
                left: 'prev,next today',
                center: 'title',
                right: ''
              }}
              events={events}
              eventClick={this.onEventClick}
              height="auto"
              eventDisplay="block"
              dayMaxEvents={3}
              moreLinkClick="popover"
              eventTimeFormat={{
                hour: 'numeric',
                minute: '2-digit',
                meridiem: 'short'
              }}
              slotLabelFormat={{
                hour: 'numeric',
                minute: '2-digit',
                meridiem: 'short'
              }}
            />
          </Stack.Item>
        </Stack>

        {/* Event Details Panel */}
        <Panel
          isOpen={isPanelOpen}
          onDismiss={this.onClosePanel}
          type={PanelType.medium}
          headerText="Leave Request Details"
          closeButtonAriaLabel="Close"
        >
          {selectedEvent && (
            <Stack tokens={{ childrenGap: 15 }}>
              <Stack.Item>
                <Text variant="large" as="h3">
                  {selectedEvent.extendedProps.leaveRequest.Requester.Title}
                </Text>
                <Text variant="medium">
                  {selectedEvent.extendedProps.leaveRequest.LeaveType.Title}
                </Text>
              </Stack.Item>

              <Stack tokens={{ childrenGap: 10 }}>
                <Stack horizontal tokens={{ childrenGap: 20 }}>
                  <Stack.Item>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      Start Date:
                    </Text>
                    <Text variant="small">
                      {selectedEvent.extendedProps.leaveRequest.StartDate.toLocaleDateString()}
                    </Text>
                  </Stack.Item>
                  <Stack.Item>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      End Date:
                    </Text>
                    <Text variant="small">
                      {selectedEvent.extendedProps.leaveRequest.EndDate.toLocaleDateString()}
                    </Text>
                  </Stack.Item>
                </Stack>

                <Stack.Item>
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    Status:
                  </Text>
                  <Text variant="small">
                    {selectedEvent.extendedProps.leaveRequest.ApprovalStatus}
                  </Text>
                </Stack.Item>

                <Stack.Item>
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    Total Days:
                  </Text>
                  <Text variant="small">
                    {selectedEvent.extendedProps.leaveRequest.TotalDays ||
                     CommonUtils.calculateBusinessDays(
                       selectedEvent.extendedProps.leaveRequest.StartDate,
                       selectedEvent.extendedProps.leaveRequest.EndDate
                     )}
                  </Text>
                </Stack.Item>

                {selectedEvent.extendedProps.leaveRequest.IsPartialDay && (
                  <Stack.Item>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      Partial Day Hours:
                    </Text>
                    <Text variant="small">
                      {selectedEvent.extendedProps.leaveRequest.PartialDayHours}
                    </Text>
                  </Stack.Item>
                )}

                {selectedEvent.extendedProps.leaveRequest.RequestComments && (
                  <Stack.Item>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      Comments:
                    </Text>
                    <Text variant="small">
                      {selectedEvent.extendedProps.leaveRequest.RequestComments}
                    </Text>
                  </Stack.Item>
                )}

                <Stack.Item>
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    Submitted:
                  </Text>
                  <Text variant="small">
                    {selectedEvent.extendedProps.leaveRequest.SubmissionDate.toLocaleDateString()}
                  </Text>
                </Stack.Item>

                {selectedEvent.extendedProps.leaveRequest.Department && (
                  <Stack.Item>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      Department:
                    </Text>
                    <Text variant="small">
                      {selectedEvent.extendedProps.leaveRequest.Department}
                    </Text>
                  </Stack.Item>
                )}
              </Stack>
            </Stack>
          )}
        </Panel>
      </section>
    );
  }
}
