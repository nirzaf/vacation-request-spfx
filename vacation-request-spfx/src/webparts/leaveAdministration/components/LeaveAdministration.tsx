import * as React from 'react';
import styles from './LeaveAdministration.module.scss';
import type { ILeaveAdministrationProps } from './ILeaveAdministrationProps';
import {
  Stack,
  Text,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Selection,
  SelectionMode,
  CommandBar,
  ICommandBarItemProps,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Panel,
  PanelType,
  DefaultButton,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Pivot,
  PivotItem,
  SearchBox,
  Label
} from '@fluentui/react';
import { ServiceManager } from '../../../services';
import {
  ILeaveRequest,
  ILeaveType,
  ApprovalStatus,
  LeaveTypeUtils,
  CommonUtils
} from '../../../models';

interface ILeaveAdministrationState {
  allLeaveRequests: ILeaveRequest[];
  filteredRequests: ILeaveRequest[];
  leaveTypes: ILeaveType[];
  isLoading: boolean;
  error: string;
  selectedRequests: ILeaveRequest[];
  selectedRequest: ILeaveRequest | undefined;
  isPanelOpen: boolean;
  isBulkApprovalDialogOpen: boolean;
  isAnalyticsLoading: boolean;
  selectedPivotKey: string;
  searchText: string;
  filters: {
    status: string;
    leaveType: string;
    department: string;
    dateFrom: Date | undefined;
    dateTo: Date | undefined;
  };
  analytics: {
    totalRequests: number;
    pendingRequests: number;
    approvedRequests: number;
    rejectedRequests: number;
    averageProcessingDays: number;
    topLeaveTypes: Array<{ name: string; count: number }>;
    departmentStats: Array<{ department: string; requests: number }>;
  };
}

export default class LeaveAdministration extends React.Component<ILeaveAdministrationProps, ILeaveAdministrationState> {
  private serviceManager: ServiceManager;
  private selection: Selection;

  constructor(props: ILeaveAdministrationProps) {
    super(props);

    this.serviceManager = new ServiceManager(props.context);
    this.selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectedRequests: this.selection.getSelection() as ILeaveRequest[]
        });
      }
    });

    this.state = {
      allLeaveRequests: [],
      filteredRequests: [],
      leaveTypes: [],
      isLoading: true,
      error: '',
      selectedRequests: [],
      selectedRequest: undefined,
      isPanelOpen: false,
      isBulkApprovalDialogOpen: false,
      isAnalyticsLoading: false,
      selectedPivotKey: 'pending',
      searchText: '',
      filters: {
        status: 'Pending',
        leaveType: '',
        department: '',
        dateFrom: undefined,
        dateTo: undefined
      },
      analytics: {
        totalRequests: 0,
        pendingRequests: 0,
        approvedRequests: 0,
        rejectedRequests: 0,
        averageProcessingDays: 0,
        topLeaveTypes: [],
        departmentStats: []
      }
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      this.setState({ isLoading: true, error: '' });

      const [leaveTypes, allRequests] = await Promise.all([
        this.serviceManager.getSharePointService().getLeaveTypes(),
        this.serviceManager.getSharePointService().getAllLeaveRequests()
      ]);

      this.setState({
        leaveTypes,
        allLeaveRequests: allRequests,
        isLoading: false
      });

      this.applyFilters();
      await this.calculateAnalytics();
    } catch (error) {
      console.error('Error loading administration data:', error);
      this.setState({
        error: 'Failed to load administration data. Please refresh the page.',
        isLoading: false
      });
    }
  }

  private applyFilters(): void {
    const { allLeaveRequests, filters, searchText } = this.state;

    let filtered = [...allLeaveRequests];

    // Apply status filter
    if (filters.status && filters.status !== 'All') {
      filtered = filtered.filter(req => req.ApprovalStatus === filters.status);
    }

    // Apply leave type filter
    if (filters.leaveType) {
      filtered = filtered.filter(req => req.LeaveType.Id.toString() === filters.leaveType);
    }

    // Apply department filter
    if (filters.department) {
      filtered = filtered.filter(req => req.Department === filters.department);
    }

    // Apply date range filter
    if (filters.dateFrom) {
      filtered = filtered.filter(req => req.StartDate >= filters.dateFrom!);
    }
    if (filters.dateTo) {
      filtered = filtered.filter(req => req.EndDate <= filters.dateTo!);
    }

    // Apply search filter
    if (searchText) {
      const searchLower = searchText.toLowerCase();
      filtered = filtered.filter(req =>
        req.Requester.Title.toLowerCase().includes(searchLower) ||
        req.LeaveType.Title.toLowerCase().includes(searchLower) ||
        (req.Department && req.Department.toLowerCase().includes(searchLower)) ||
        (req.RequestComments && req.RequestComments.toLowerCase().includes(searchLower))
      );
    }

    this.setState({ filteredRequests: filtered });
  }

  private async calculateAnalytics(): Promise<void> {
    try {
      this.setState({ isAnalyticsLoading: true });

      const { allLeaveRequests, leaveTypes } = this.state;

      const totalRequests = allLeaveRequests.length;
      const pendingRequests = allLeaveRequests.filter(r => r.ApprovalStatus === 'Pending').length;
      const approvedRequests = allLeaveRequests.filter(r => r.ApprovalStatus === 'Approved').length;
      const rejectedRequests = allLeaveRequests.filter(r => r.ApprovalStatus === 'Rejected').length;

      // Calculate average processing days
      const processedRequests = allLeaveRequests.filter(r => r.ApprovalDate);
      const averageProcessingDays = processedRequests.length > 0
        ? processedRequests.reduce((sum, req) => {
            const processingDays = Math.ceil(
              (req.ApprovalDate!.getTime() - req.SubmissionDate.getTime()) / (1000 * 60 * 60 * 24)
            );
            return sum + processingDays;
          }, 0) / processedRequests.length
        : 0;

      // Calculate top leave types
      const leaveTypeCounts = new Map<number, number>();
      allLeaveRequests.forEach(req => {
        const count = leaveTypeCounts.get(req.LeaveType.Id) || 0;
        leaveTypeCounts.set(req.LeaveType.Id, count + 1);
      });

      const topLeaveTypes = Array.from(leaveTypeCounts.entries())
        .map(([typeId, count]) => ({
          name: leaveTypes.filter(lt => lt.Id === typeId)[0]?.Title || 'Unknown',
          count
        }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

      // Calculate department stats
      const departmentCounts = new Map<string, number>();
      allLeaveRequests.forEach(req => {
        if (req.Department) {
          const count = departmentCounts.get(req.Department) || 0;
          departmentCounts.set(req.Department, count + 1);
        }
      });

      const departmentStats = Array.from(departmentCounts.entries())
        .map(([department, requests]) => ({ department, requests }))
        .sort((a, b) => b.requests - a.requests);

      this.setState({
        analytics: {
          totalRequests,
          pendingRequests,
          approvedRequests,
          rejectedRequests,
          averageProcessingDays: Math.round(averageProcessingDays * 10) / 10,
          topLeaveTypes,
          departmentStats
        },
        isAnalyticsLoading: false
      });
    } catch (error) {
      console.error('Error calculating analytics:', error);
      this.setState({ isAnalyticsLoading: false });
    }
  }

  private getColumns(): IColumn[] {
    return [
      {
        key: 'requester',
        name: 'Employee',
        fieldName: 'requester',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <Stack>
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              {item.Requester.Title}
            </Text>
            {item.Department && (
              <Text variant="small" styles={{ root: { color: '#666' } }}>
                {item.Department}
              </Text>
            )}
          </Stack>
        )
      },
      {
        key: 'leaveType',
        name: 'Leave Type',
        fieldName: 'leaveType',
        minWidth: 120,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <span style={{
            color: this.getLeaveTypeColor(item.LeaveType.Id),
            fontWeight: 600
          }}>
            {item.LeaveType.Title}
          </span>
        )
      },
      {
        key: 'dates',
        name: 'Dates',
        fieldName: 'dates',
        minWidth: 180,
        maxWidth: 220,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <Stack>
            <Text variant="small">
              {item.StartDate.toLocaleDateString()} - {item.EndDate.toLocaleDateString()}
            </Text>
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              {item.TotalDays || CommonUtils.calculateBusinessDays(item.StartDate, item.EndDate)} days
              {item.IsPartialDay && ` (${item.PartialDayHours}h)`}
            </Text>
          </Stack>
        )
      },
      {
        key: 'status',
        name: 'Status',
        fieldName: 'status',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <span className={`${styles.statusBadge} ${this.getStatusClass(item.ApprovalStatus)}`}>
            {item.ApprovalStatus}
          </span>
        )
      },
      {
        key: 'submitted',
        name: 'Submitted',
        fieldName: 'submitted',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <Text variant="small">
            {item.SubmissionDate.toLocaleDateString()}
          </Text>
        )
      },
      {
        key: 'actions',
        name: 'Actions',
        fieldName: 'actions',
        minWidth: 120,
        maxWidth: 150,
        isResizable: false,
        onRender: (item: ILeaveRequest) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="View"
              iconProps={{ iconName: 'View' }}
              onClick={() => this.onViewRequest(item)}
              styles={{ root: { minWidth: 'auto' } }}
            />
            {item.ApprovalStatus === ApprovalStatus.Pending && (
              <>
                <PrimaryButton
                  text="Approve"
                  iconProps={{ iconName: 'CheckMark' }}
                  onClick={() => this.onApproveRequest(item)}
                  styles={{ root: { minWidth: 'auto' } }}
                />
                <DefaultButton
                  text="Reject"
                  iconProps={{ iconName: 'Cancel' }}
                  onClick={() => this.onRejectRequest(item)}
                  styles={{ root: { minWidth: 'auto' } }}
                />
              </>
            )}
          </Stack>
        )
      }
    ];
  }

  private getLeaveTypeColor(leaveTypeId: number): string {
    const leaveType = this.state.leaveTypes.filter((lt: ILeaveType) => lt.Id === leaveTypeId)[0];
    return leaveType?.ColorCode || '#0078d4';
  }

  private getStatusClass(status: string): string {
    switch (status.toLowerCase()) {
      case 'pending': return styles.pending;
      case 'approved': return styles.approved;
      case 'rejected': return styles.rejected;
      case 'cancelled': return styles.cancelled;
      default: return '';
    }
  }

  private onViewRequest = (request: ILeaveRequest): void => {
    this.setState({
      selectedRequest: request,
      isPanelOpen: true
    });
  };

  private onApproveRequest = async (request: ILeaveRequest): Promise<void> => {
    try {
      await this.serviceManager.approveLeaveRequestWithWorkflow(request.Id);
      await this.loadData();
    } catch (error) {
      console.error('Error approving request:', error);
      this.setState({ error: 'Failed to approve request. Please try again.' });
    }
  };

  private onRejectRequest = async (request: ILeaveRequest): Promise<void> => {
    try {
      await this.serviceManager.rejectLeaveRequestWithWorkflow(request.Id, 'Request rejected by administrator');
      await this.loadData();
    } catch (error) {
      console.error('Error rejecting request:', error);
      this.setState({ error: 'Failed to reject request. Please try again.' });
    }
  };

  private onBulkApprove = (): void => {
    if (this.state.selectedRequests.length > 0) {
      this.setState({ isBulkApprovalDialogOpen: true });
    }
  };

  private onConfirmBulkApprove = async (): Promise<void> => {
    try {
      const { selectedRequests } = this.state;

      for (const request of selectedRequests) {
        if (request.ApprovalStatus === ApprovalStatus.Pending) {
          await this.serviceManager.approveLeaveRequestWithWorkflow(request.Id);
        }
      }

      await this.loadData();
      this.setState({ isBulkApprovalDialogOpen: false });
      this.selection.setAllSelected(false);
    } catch (error) {
      console.error('Error in bulk approval:', error);
      this.setState({ error: 'Failed to process bulk approval. Please try again.' });
    }
  };

  private onClosePanel = (): void => {
    this.setState({
      isPanelOpen: false,
      selectedRequest: undefined
    });
  };

  private onCloseBulkDialog = (): void => {
    this.setState({ isBulkApprovalDialogOpen: false });
  };

  private onSearchChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    this.setState({ searchText: newValue || '' }, () => {
      this.applyFilters();
    });
  };

  private onFilterChange = (filterType: string, value: any): void => {
    this.setState({
      filters: { ...this.state.filters, [filterType]: value }
    }, () => {
      this.applyFilters();
    });
  };

  private onPivotChange = (item?: PivotItem): void => {
    if (item) {
      const pivotKey = item.props.itemKey || 'pending';
      this.setState({
        selectedPivotKey: pivotKey,
        filters: { ...this.state.filters, status: pivotKey === 'all' ? '' : this.capitalizeFirst(pivotKey) }
      }, () => {
        this.applyFilters();
      });
    }
  };

  private capitalizeFirst(str: string): string {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  private onRefresh = async (): Promise<void> => {
    await this.loadData();
  };

  private getCommandBarItems(): ICommandBarItemProps[] {
    const { selectedRequests } = this.state;
    const pendingSelected = selectedRequests.filter(r => r.ApprovalStatus === ApprovalStatus.Pending);

    return [
      {
        key: 'bulkApprove',
        text: `Approve Selected (${pendingSelected.length})`,
        iconProps: { iconName: 'CheckMark' },
        disabled: pendingSelected.length === 0,
        onClick: this.onBulkApprove
      },
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: this.onRefresh
      },
      {
        key: 'export',
        text: 'Export',
        iconProps: { iconName: 'Download' },
        onClick: this.onExport
      }
    ];
  }

  private onExport = (): void => {
    const { filteredRequests } = this.state;
    const csvContent = this.convertRequestsToCSV(filteredRequests);
    this.downloadCSV(csvContent, 'leave-administration-report.csv');
  };

  private convertRequestsToCSV(requests: ILeaveRequest[]): string {
    const headers = ['Employee', 'Department', 'Leave Type', 'Start Date', 'End Date', 'Total Days', 'Status', 'Submitted', 'Approved', 'Comments'];
    const rows = requests.map(request => [
      request.Requester.Title,
      request.Department || '',
      request.LeaveType.Title,
      request.StartDate.toLocaleDateString(),
      request.EndDate.toLocaleDateString(),
      (request.TotalDays || CommonUtils.calculateBusinessDays(request.StartDate, request.EndDate)).toString(),
      request.ApprovalStatus,
      request.SubmissionDate.toLocaleDateString(),
      request.ApprovalDate ? request.ApprovalDate.toLocaleDateString() : '',
      request.RequestComments || ''
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

  public render(): React.ReactElement<ILeaveAdministrationProps> {
    const { hasTeamsContext, userDisplayName } = this.props;
    const {
      filteredRequests,
      leaveTypes,
      isLoading,
      error,
      selectedRequest,
      isPanelOpen,
      isBulkApprovalDialogOpen,
      selectedRequests,
      selectedPivotKey,
      searchText,
      filters,
      analytics,
      isAnalyticsLoading
    } = this.state;

    if (isLoading) {
      return (
        <section className={`${styles.leaveAdministration} ${hasTeamsContext ? styles.teams : ''}`}>
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Loading administration dashboard..." />
          </Stack>
        </section>
      );
    }

    const leaveTypeOptions: IDropdownOption[] = [
      { key: '', text: 'All Leave Types' },
      ...LeaveTypeUtils.toDropdownOptions(leaveTypes)
    ];

    const statusOptions: IDropdownOption[] = [
      { key: '', text: 'All Statuses' },
      { key: 'Pending', text: 'Pending' },
      { key: 'Approved', text: 'Approved' },
      { key: 'Rejected', text: 'Rejected' },
      { key: 'Cancelled', text: 'Cancelled' }
    ];

    return (
      <section className={`${styles.leaveAdministration} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="xxLarge" as="h1">Leave Administration</Text>
            <Text variant="medium">Welcome, {userDisplayName}! Manage team leave requests and view analytics.</Text>
          </Stack.Item>

          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}

          <Pivot selectedKey={selectedPivotKey} onLinkClick={this.onPivotChange}>
            <PivotItem headerText="Pending Requests" itemKey="pending">
              <Stack tokens={{ childrenGap: 16 }}>
                {/* Filters Section */}
                <div className={styles.filtersSection}>
                  <Stack horizontal wrap tokens={{ childrenGap: 15 }} verticalAlign="end">
                    <SearchBox
                      placeholder="Search by employee, leave type, or department"
                      value={searchText}
                      onChange={this.onSearchChange}
                      styles={{ root: { minWidth: 250 } }}
                    />

                    <Dropdown
                      label="Leave Type"
                      options={leaveTypeOptions}
                      selectedKey={filters.leaveType}
                      onChange={(e, option) => this.onFilterChange('leaveType', option?.key || '')}
                      styles={{ root: { minWidth: 150 } }}
                    />

                    <TextField
                      label="Department"
                      value={filters.department}
                      onChange={(e, value) => this.onFilterChange('department', value || '')}
                      styles={{ root: { minWidth: 120 } }}
                    />

                    <DatePicker
                      label="From Date"
                      value={filters.dateFrom}
                      onSelectDate={(date) => this.onFilterChange('dateFrom', date)}
                      styles={{ root: { minWidth: 120 } }}
                    />

                    <DatePicker
                      label="To Date"
                      value={filters.dateTo}
                      onSelectDate={(date) => this.onFilterChange('dateTo', date)}
                      styles={{ root: { minWidth: 120 } }}
                    />
                  </Stack>
                </div>

                <CommandBar items={this.getCommandBarItems()} />

                <DetailsList
                  items={filteredRequests}
                  columns={this.getColumns()}
                  layoutMode={DetailsListLayoutMode.justified}
                  selection={this.selection}
                  selectionMode={SelectionMode.multiple}
                  isHeaderVisible={true}
                  className={styles.requestsList}
                />

                {filteredRequests.length === 0 && (
                  <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                    <Text variant="large">No requests found</Text>
                    <Text variant="medium">No leave requests match your current filters.</Text>
                  </Stack>
                )}
              </Stack>
            </PivotItem>

            <PivotItem headerText="All Requests" itemKey="all">
              <Stack tokens={{ childrenGap: 16 }}>
                {/* Same filters and list as pending, but with all statuses */}
                <div className={styles.filtersSection}>
                  <Stack horizontal wrap tokens={{ childrenGap: 15 }} verticalAlign="end">
                    <SearchBox
                      placeholder="Search by employee, leave type, or department"
                      value={searchText}
                      onChange={this.onSearchChange}
                      styles={{ root: { minWidth: 250 } }}
                    />

                    <Dropdown
                      label="Status"
                      options={statusOptions}
                      selectedKey={filters.status}
                      onChange={(e, option) => this.onFilterChange('status', option?.key || '')}
                      styles={{ root: { minWidth: 120 } }}
                    />

                    <Dropdown
                      label="Leave Type"
                      options={leaveTypeOptions}
                      selectedKey={filters.leaveType}
                      onChange={(e, option) => this.onFilterChange('leaveType', option?.key || '')}
                      styles={{ root: { minWidth: 150 } }}
                    />

                    <TextField
                      label="Department"
                      value={filters.department}
                      onChange={(e, value) => this.onFilterChange('department', value || '')}
                      styles={{ root: { minWidth: 120 } }}
                    />
                  </Stack>
                </div>

                <CommandBar items={this.getCommandBarItems()} />

                <DetailsList
                  items={filteredRequests}
                  columns={this.getColumns()}
                  layoutMode={DetailsListLayoutMode.justified}
                  selection={this.selection}
                  selectionMode={SelectionMode.multiple}
                  isHeaderVisible={true}
                  className={styles.requestsList}
                />
              </Stack>
            </PivotItem>

            <PivotItem headerText="Analytics" itemKey="analytics">
              {isAnalyticsLoading ? (
                <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                  <Spinner size={SpinnerSize.medium} label="Loading analytics..." />
                </Stack>
              ) : (
                <Stack tokens={{ childrenGap: 24 }}>
                  {/* Summary Cards */}
                  <div className={styles.summaryCards}>
                    <div className={styles.summaryCard}>
                      <div className={styles.summaryNumber}>{analytics.totalRequests}</div>
                      <div className={styles.summaryLabel}>Total Requests</div>
                    </div>
                    <div className={styles.summaryCard}>
                      <div className={styles.summaryNumber}>{analytics.pendingRequests}</div>
                      <div className={styles.summaryLabel}>Pending</div>
                    </div>
                    <div className={styles.summaryCard}>
                      <div className={styles.summaryNumber}>{analytics.approvedRequests}</div>
                      <div className={styles.summaryLabel}>Approved</div>
                    </div>
                    <div className={styles.summaryCard}>
                      <div className={styles.summaryNumber}>{analytics.rejectedRequests}</div>
                      <div className={styles.summaryLabel}>Rejected</div>
                    </div>
                    <div className={styles.summaryCard}>
                      <div className={styles.summaryNumber}>{analytics.averageProcessingDays}</div>
                      <div className={styles.summaryLabel}>Avg. Processing Days</div>
                    </div>
                  </div>

                  <Stack horizontal tokens={{ childrenGap: 32 }}>
                    {/* Top Leave Types */}
                    <Stack.Item grow>
                      <Text variant="xLarge" as="h3">Top Leave Types</Text>
                      <Stack tokens={{ childrenGap: 8 }}>
                        {analytics.topLeaveTypes.map((item, index) => (
                          <Stack key={index} horizontal horizontalAlign="space-between">
                            <Text variant="medium">{item.name}</Text>
                            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                              {item.count}
                            </Text>
                          </Stack>
                        ))}
                      </Stack>
                    </Stack.Item>

                    {/* Department Stats */}
                    <Stack.Item grow>
                      <Text variant="xLarge" as="h3">Department Statistics</Text>
                      <Stack tokens={{ childrenGap: 8 }}>
                        {analytics.departmentStats.slice(0, 10).map((item, index) => (
                          <Stack key={index} horizontal horizontalAlign="space-between">
                            <Text variant="medium">{item.department}</Text>
                            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                              {item.requests}
                            </Text>
                          </Stack>
                        ))}
                      </Stack>
                    </Stack.Item>
                  </Stack>
                </Stack>
              )}
            </PivotItem>
          </Pivot>
        </Stack>

        {/* Request Details Panel */}
        <Panel
          isOpen={isPanelOpen}
          onDismiss={this.onClosePanel}
          type={PanelType.medium}
          headerText="Leave Request Details"
          closeButtonAriaLabel="Close"
        >
          {selectedRequest && (
            <Stack tokens={{ childrenGap: 15 }}>
              <Stack.Item>
                <Text variant="large" as="h3">
                  {selectedRequest.Requester.Title}
                </Text>
                <Text variant="medium">
                  {selectedRequest.LeaveType.Title}
                </Text>
              </Stack.Item>

              <Stack tokens={{ childrenGap: 10 }}>
                <Stack horizontal tokens={{ childrenGap: 20 }}>
                  <Stack.Item>
                    <Label>Start Date:</Label>
                    <Text>{selectedRequest.StartDate.toLocaleDateString()}</Text>
                  </Stack.Item>
                  <Stack.Item>
                    <Label>End Date:</Label>
                    <Text>{selectedRequest.EndDate.toLocaleDateString()}</Text>
                  </Stack.Item>
                </Stack>

                <Stack.Item>
                  <Label>Status:</Label>
                  <span className={`${styles.statusBadge} ${this.getStatusClass(selectedRequest.ApprovalStatus)}`}>
                    {selectedRequest.ApprovalStatus}
                  </span>
                </Stack.Item>

                <Stack.Item>
                  <Label>Total Days:</Label>
                  <Text>
                    {selectedRequest.TotalDays ||
                     CommonUtils.calculateBusinessDays(selectedRequest.StartDate, selectedRequest.EndDate)}
                  </Text>
                </Stack.Item>

                {selectedRequest.IsPartialDay && (
                  <Stack.Item>
                    <Label>Partial Day Hours:</Label>
                    <Text>{selectedRequest.PartialDayHours}</Text>
                  </Stack.Item>
                )}

                {selectedRequest.RequestComments && (
                  <Stack.Item>
                    <Label>Comments:</Label>
                    <Text>{selectedRequest.RequestComments}</Text>
                  </Stack.Item>
                )}

                <Stack.Item>
                  <Label>Submitted:</Label>
                  <Text>{selectedRequest.SubmissionDate.toLocaleDateString()}</Text>
                </Stack.Item>

                {selectedRequest.ApprovalDate && (
                  <Stack.Item>
                    <Label>Approval Date:</Label>
                    <Text>{selectedRequest.ApprovalDate.toLocaleDateString()}</Text>
                  </Stack.Item>
                )}

                {selectedRequest.ApprovalComments && (
                  <Stack.Item>
                    <Label>Approval Comments:</Label>
                    <Text>{selectedRequest.ApprovalComments}</Text>
                  </Stack.Item>
                )}

                {selectedRequest.Department && (
                  <Stack.Item>
                    <Label>Department:</Label>
                    <Text>{selectedRequest.Department}</Text>
                  </Stack.Item>
                )}

                {selectedRequest.AttachmentURL && (
                  <Stack.Item>
                    <Label>Attachment:</Label>
                    <a href={selectedRequest.AttachmentURL} target="_blank" rel="noopener noreferrer">
                      View Document
                    </a>
                  </Stack.Item>
                )}
              </Stack>

              {selectedRequest.ApprovalStatus === ApprovalStatus.Pending && (
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                  <PrimaryButton
                    text="Approve"
                    iconProps={{ iconName: 'CheckMark' }}
                    onClick={() => this.onApproveRequest(selectedRequest)}
                  />
                  <DefaultButton
                    text="Reject"
                    iconProps={{ iconName: 'Cancel' }}
                    onClick={() => this.onRejectRequest(selectedRequest)}
                  />
                </Stack>
              )}
            </Stack>
          )}
        </Panel>

        {/* Bulk Approval Dialog */}
        <Dialog
          hidden={!isBulkApprovalDialogOpen}
          onDismiss={this.onCloseBulkDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Bulk Approval',
            subText: `Are you sure you want to approve ${selectedRequests.filter(r => r.ApprovalStatus === ApprovalStatus.Pending).length} pending leave requests?`
          }}
          modalProps={{ isBlocking: true }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.onConfirmBulkApprove} text="Yes, Approve All" />
            <DefaultButton onClick={this.onCloseBulkDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
