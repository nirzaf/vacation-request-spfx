import * as React from 'react';
import styles from './LeaveHistory.module.scss';
import type { ILeaveHistoryProps } from './ILeaveHistoryProps';
import {
  Stack,
  Text,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
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
  Checkbox,
  Label,
  Pivot,
  PivotItem,

} from '@fluentui/react';
import { ServiceManager } from '../../../services';
import {
  ILeaveRequest,
  ILeaveBalance,
  ILeaveType,
  ApprovalStatus,
  LeaveTypeUtils,
  CommonUtils
} from '../../../models';

interface ILeaveHistoryState {
  leaveRequests: ILeaveRequest[];
  leaveBalances: ILeaveBalance[];
  leaveTypes: ILeaveType[];
  isLoading: boolean;
  error: string;
  selectedRequest: ILeaveRequest | undefined;
  isPanelOpen: boolean;
  isEditDialogOpen: boolean;
  isCancelDialogOpen: boolean;
  editFormData: any;
  selectedPivotKey: string;
}

export default class LeaveHistory extends React.Component<ILeaveHistoryProps, ILeaveHistoryState> {
  private serviceManager: ServiceManager;

  constructor(props: ILeaveHistoryProps) {
    super(props);

    this.serviceManager = new ServiceManager(props.context);

    this.state = {
      leaveRequests: [],
      leaveBalances: [],
      leaveTypes: [],
      isLoading: true,
      error: '',
      selectedRequest: undefined,
      isPanelOpen: false,
      isEditDialogOpen: false,
      isCancelDialogOpen: false,
      editFormData: {},
      selectedPivotKey: 'requests'
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      this.setState({ isLoading: true, error: '' });

      const dashboardData = await this.serviceManager.getUserDashboardData();

      this.setState({
        leaveRequests: dashboardData.leaveRequests,
        leaveBalances: dashboardData.leaveBalances,
        leaveTypes: dashboardData.leaveTypes,
        isLoading: false
      });
    } catch (error) {
      console.error('Error loading leave history data:', error);
      this.setState({
        error: 'Failed to load leave history. Please refresh the page.',
        isLoading: false
      });
    }
  }

  private getColumns(): IColumn[] {
    return [
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
        key: 'startDate',
        name: 'Start Date',
        fieldName: 'startDate',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ILeaveRequest) => item.StartDate.toLocaleDateString()
      },
      {
        key: 'endDate',
        name: 'End Date',
        fieldName: 'endDate',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ILeaveRequest) => item.EndDate.toLocaleDateString()
      },
      {
        key: 'totalDays',
        name: 'Days',
        fieldName: 'totalDays',
        minWidth: 60,
        maxWidth: 80,
        isResizable: true,
        onRender: (item: ILeaveRequest) => (
          <span>
            {item.TotalDays || CommonUtils.calculateBusinessDays(item.StartDate, item.EndDate)}
            {item.IsPartialDay && ` (${item.PartialDayHours}h)`}
          </span>
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
        key: 'submissionDate',
        name: 'Submitted',
        fieldName: 'submissionDate',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ILeaveRequest) => item.SubmissionDate.toLocaleDateString()
      },
      {
        key: 'actions',
        name: 'Actions',
        fieldName: 'actions',
        minWidth: 100,
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
                <DefaultButton
                  text="Edit"
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => this.onEditRequest(item)}
                  styles={{ root: { minWidth: 'auto' } }}
                />
                <DefaultButton
                  text="Cancel"
                  iconProps={{ iconName: 'Cancel' }}
                  onClick={() => this.onCancelRequest(item)}
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

  private onEditRequest = (request: ILeaveRequest): void => {
    this.setState({
      selectedRequest: request,
      editFormData: {
        leaveTypeId: request.LeaveType.Id,
        startDate: request.StartDate,
        endDate: request.EndDate,
        isPartialDay: request.IsPartialDay,
        partialDayHours: request.PartialDayHours,
        comments: request.RequestComments || '',
        attachmentUrl: request.AttachmentURL || ''
      },
      isEditDialogOpen: true
    });
  };

  private onCancelRequest = (request: ILeaveRequest): void => {
    this.setState({
      selectedRequest: request,
      isCancelDialogOpen: true
    });
  };

  private onClosePanel = (): void => {
    this.setState({
      isPanelOpen: false,
      selectedRequest: undefined
    });
  };

  private onCloseEditDialog = (): void => {
    this.setState({
      isEditDialogOpen: false,
      selectedRequest: undefined,
      editFormData: {}
    });
  };

  private onCloseCancelDialog = (): void => {
    this.setState({
      isCancelDialogOpen: false,
      selectedRequest: undefined
    });
  };

  private onSaveEdit = async (): Promise<void> => {
    if (!this.state.selectedRequest) return;

    try {
      const { editFormData, selectedRequest } = this.state;

      await this.serviceManager.getSharePointService().updateLeaveRequest(selectedRequest.Id, {
        LeaveTypeId: editFormData.leaveTypeId,
        StartDate: editFormData.startDate,
        EndDate: editFormData.endDate,
        IsPartialDay: editFormData.isPartialDay,
        PartialDayHours: editFormData.partialDayHours,
        RequestComments: editFormData.comments,
        AttachmentURL: editFormData.attachmentUrl
      });

      await this.loadData();
      this.onCloseEditDialog();
    } catch (error) {
      console.error('Error updating leave request:', error);
      this.setState({ error: 'Failed to update leave request. Please try again.' });
    }
  };

  private onConfirmCancel = async (): Promise<void> => {
    if (!this.state.selectedRequest) return;

    try {
      await this.serviceManager.cancelLeaveRequestWithCleanup(this.state.selectedRequest.Id);
      await this.loadData();
      this.onCloseCancelDialog();
    } catch (error) {
      console.error('Error cancelling leave request:', error);
      this.setState({ error: 'Failed to cancel leave request. Please try again.' });
    }
  };

  private onRefresh = async (): Promise<void> => {
    await this.loadData();
  };

  private onPivotChange = (item?: PivotItem): void => {
    if (item) {
      this.setState({ selectedPivotKey: item.props.itemKey || 'requests' });
    }
  };

  private getCommandBarItems(): ICommandBarItemProps[] {
    return [
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
    const { leaveRequests } = this.state;
    const csvContent = this.convertRequestsToCSV(leaveRequests);
    this.downloadCSV(csvContent, 'my-leave-requests.csv');
  };

  private convertRequestsToCSV(requests: ILeaveRequest[]): string {
    const headers = ['Leave Type', 'Start Date', 'End Date', 'Total Days', 'Status', 'Submitted', 'Comments'];
    const rows = requests.map(request => [
      request.LeaveType.Title,
      request.StartDate.toLocaleDateString(),
      request.EndDate.toLocaleDateString(),
      (request.TotalDays || CommonUtils.calculateBusinessDays(request.StartDate, request.EndDate)).toString(),
      request.ApprovalStatus,
      request.SubmissionDate.toLocaleDateString(),
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

  private renderBalanceCards(): React.ReactElement {
    const { leaveBalances } = this.state;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Text variant="xLarge" as="h2">Leave Balances</Text>
        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          {leaveBalances.map(balance => (
            <div key={balance.Id} className={styles.balanceCard}>
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                  {balance.LeaveType.Title}
                </Text>
                <Stack horizontal horizontalAlign="space-between">
                  <Text variant="medium">Total Allowance:</Text>
                  <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {balance.TotalAllowance} days
                  </Text>
                </Stack>
                <Stack horizontal horizontalAlign="space-between">
                  <Text variant="medium">Used:</Text>
                  <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {balance.UsedDays} days
                  </Text>
                </Stack>
                <Stack horizontal horizontalAlign="space-between">
                  <Text variant="medium">Remaining:</Text>
                  <Text variant="medium" styles={{
                    root: {
                      fontWeight: 600,
                      color: balance.RemainingDays < 5 ? '#d13438' : '#107c10'
                    }
                  }}>
                    {balance.RemainingDays} days
                  </Text>
                </Stack>
                {balance.CarryOverDays > 0 && (
                  <Stack horizontal horizontalAlign="space-between">
                    <Text variant="small">Carry Over:</Text>
                    <Text variant="small">{balance.CarryOverDays} days</Text>
                  </Stack>
                )}
                <Stack horizontal horizontalAlign="space-between">
                  <Text variant="small">Expires:</Text>
                  <Text variant="small">{balance.ExpirationDate.toLocaleDateString()}</Text>
                </Stack>
              </Stack>
            </div>
          ))}
        </Stack>
      </Stack>
    );
  }

  public render(): React.ReactElement<ILeaveHistoryProps> {
    const { hasTeamsContext, userDisplayName } = this.props;
    const {
      leaveRequests,
      leaveTypes,
      isLoading,
      error,
      selectedRequest,
      isPanelOpen,
      isEditDialogOpen,
      isCancelDialogOpen,
      editFormData,
      selectedPivotKey
    } = this.state;

    if (isLoading) {
      return (
        <section className={`${styles.leaveHistory} ${hasTeamsContext ? styles.teams : ''}`}>
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Loading your leave history..." />
          </Stack>
        </section>
      );
    }

    const leaveTypeOptions: IDropdownOption[] = LeaveTypeUtils.toDropdownOptions(leaveTypes);

    return (
      <section className={`${styles.leaveHistory} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="xxLarge" as="h1">My Leave History</Text>
            <Text variant="medium">Welcome back, {userDisplayName}! Track and manage your leave requests.</Text>
          </Stack.Item>

          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}

          <Pivot selectedKey={selectedPivotKey} onLinkClick={this.onPivotChange}>
            <PivotItem headerText="Leave Requests" itemKey="requests">
              <Stack tokens={{ childrenGap: 16 }}>
                <CommandBar items={this.getCommandBarItems()} />

                <DetailsList
                  items={leaveRequests}
                  columns={this.getColumns()}
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  isHeaderVisible={true}
                  className={styles.requestsList}
                />

                {leaveRequests.length === 0 && (
                  <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                    <Text variant="large">No leave requests found</Text>
                    <Text variant="medium">You haven't submitted any leave requests yet.</Text>
                  </Stack>
                )}
              </Stack>
            </PivotItem>

            <PivotItem headerText="Leave Balances" itemKey="balances">
              {this.renderBalanceCards()}
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

                {selectedRequest.AttachmentURL && (
                  <Stack.Item>
                    <Label>Attachment:</Label>
                    <a href={selectedRequest.AttachmentURL} target="_blank" rel="noopener noreferrer">
                      View Document
                    </a>
                  </Stack.Item>
                )}
              </Stack>
            </Stack>
          )}
        </Panel>

        {/* Edit Request Dialog */}
        <Dialog
          hidden={!isEditDialogOpen}
          onDismiss={this.onCloseEditDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Edit Leave Request',
            subText: 'Modify your pending leave request details'
          }}
          modalProps={{ isBlocking: true }}
        >
          <Stack tokens={{ childrenGap: 15 }}>
            <Dropdown
              label="Leave Type"
              options={leaveTypeOptions}
              selectedKey={editFormData.leaveTypeId}
              onChange={(e, option) => this.setState({
                editFormData: { ...editFormData, leaveTypeId: option?.key }
              })}
            />

            <Stack horizontal tokens={{ childrenGap: 15 }}>
              <DatePicker
                label="Start Date"
                value={editFormData.startDate}
                onSelectDate={(date) => this.setState({
                  editFormData: { ...editFormData, startDate: date }
                })}
              />
              <DatePicker
                label="End Date"
                value={editFormData.endDate}
                onSelectDate={(date) => this.setState({
                  editFormData: { ...editFormData, endDate: date }
                })}
              />
            </Stack>

            <Checkbox
              label="Partial Day Request"
              checked={editFormData.isPartialDay}
              onChange={(e, checked) => this.setState({
                editFormData: { ...editFormData, isPartialDay: !!checked }
              })}
            />

            {editFormData.isPartialDay && (
              <TextField
                label="Hours"
                type="number"
                value={editFormData.partialDayHours?.toString() || ''}
                onChange={(e, value) => this.setState({
                  editFormData: { ...editFormData, partialDayHours: parseFloat(value || '0') }
                })}
              />
            )}

            <TextField
              label="Comments"
              multiline
              rows={3}
              value={editFormData.comments}
              onChange={(e, value) => this.setState({
                editFormData: { ...editFormData, comments: value || '' }
              })}
            />

            <TextField
              label="Attachment URL"
              value={editFormData.attachmentUrl}
              onChange={(e, value) => this.setState({
                editFormData: { ...editFormData, attachmentUrl: value || '' }
              })}
            />
          </Stack>

          <DialogFooter>
            <PrimaryButton onClick={this.onSaveEdit} text="Save Changes" />
            <DefaultButton onClick={this.onCloseEditDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>

        {/* Cancel Request Dialog */}
        <Dialog
          hidden={!isCancelDialogOpen}
          onDismiss={this.onCloseCancelDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Cancel Leave Request',
            subText: 'Are you sure you want to cancel this leave request? This action cannot be undone.'
          }}
          modalProps={{ isBlocking: true }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.onConfirmCancel} text="Yes, Cancel Request" />
            <DefaultButton onClick={this.onCloseCancelDialog} text="No, Keep Request" />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
