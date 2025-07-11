import * as React from 'react';
import styles from './LeaveRequestForm.module.scss';
import type { ILeaveRequestFormProps } from './ILeaveRequestFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  PrimaryButton,
  DefaultButton,
  TextField,
  DatePicker,
  Dropdown,
  IDropdownOption,
  Checkbox,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  Label
} from '@fluentui/react';
import { SharePointService } from '../../../services/SharePointService';
import {
  ILeaveType,
  ILeaveRequestCreate,
  ILeaveRequestFormData,
  LeaveTypeUtils,
  CommonUtils
} from '../../../models';

interface ILeaveRequestFormState {
  leaveTypes: ILeaveType[];
  formData: ILeaveRequestFormData;
  isLoading: boolean;
  isSubmitting: boolean;
  errors: string[];
  successMessage: string;
  selectedLeaveType: ILeaveType | null;
}

export default class LeaveRequestForm extends React.Component<ILeaveRequestFormProps, ILeaveRequestFormState> {
  private sharePointService: SharePointService;

  constructor(props: ILeaveRequestFormProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);

    this.state = {
      leaveTypes: [],
      formData: {
        leaveTypeId: 0,
        startDate: new Date(),
        endDate: new Date(),
        isPartialDay: false,
        partialDayHours: undefined,
        comments: '',
        attachmentUrl: ''
      },
      isLoading: true,
      isSubmitting: false,
      errors: [],
      successMessage: '',
      selectedLeaveType: null
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadLeaveTypes();
  }

  private async loadLeaveTypes(): Promise<void> {
    try {
      this.setState({ isLoading: true, errors: [] });
      const leaveTypes = await this.sharePointService.getLeaveTypes();
      this.setState({ leaveTypes, isLoading: false });
    } catch (error) {
      console.error('Error loading leave types:', error);
      this.setState({
        errors: ['Failed to load leave types. Please refresh the page.'],
        isLoading: false
      });
    }
  }

  private onLeaveTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      const selectedLeaveType = this.state.leaveTypes.filter((lt: ILeaveType) => lt.Id === option.key)[0] || null;
      this.setState({
        formData: { ...this.state.formData, leaveTypeId: option.key as number },
        selectedLeaveType,
        errors: []
      });
    }
  };

  private onStartDateChange = (date: Date | null | undefined): void => {
    if (date) {
      const endDate = this.state.formData.endDate < date ? date : this.state.formData.endDate;
      this.setState({
        formData: { ...this.state.formData, startDate: date, endDate },
        errors: []
      });
    }
  };

  private onEndDateChange = (date: Date | null | undefined): void => {
    if (date) {
      this.setState({
        formData: { ...this.state.formData, endDate: date },
        errors: []
      });
    }
  };

  private onPartialDayChange = (event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    this.setState({
      formData: {
        ...this.state.formData,
        isPartialDay: !!checked,
        partialDayHours: checked ? 4 : undefined
      },
      errors: []
    });
  };

  private onPartialDayHoursChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const hours = newValue ? parseFloat(newValue) : undefined;
    this.setState({
      formData: { ...this.state.formData, partialDayHours: hours },
      errors: []
    });
  };

  private onCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({
      formData: { ...this.state.formData, comments: newValue || '' },
      errors: []
    });
  };

  private onAttachmentUrlChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({
      formData: { ...this.state.formData, attachmentUrl: newValue || '' },
      errors: []
    });
  };

  private validateForm(): string[] {
    const errors: string[] = [];
    const { formData, selectedLeaveType } = this.state;

    if (!formData.leaveTypeId) {
      errors.push('Please select a leave type');
    }

    if (!formData.startDate) {
      errors.push('Please select a start date');
    }

    if (!formData.endDate) {
      errors.push('Please select an end date');
    }

    if (formData.startDate && formData.endDate && formData.endDate < formData.startDate) {
      errors.push('End date must be after start date');
    }

    if (formData.isPartialDay) {
      if (!formData.partialDayHours || formData.partialDayHours <= 0 || formData.partialDayHours > 8) {
        errors.push('Partial day hours must be between 0.5 and 8 hours');
      }
    }

    if (selectedLeaveType?.RequiresDocumentation && !formData.attachmentUrl) {
      errors.push(`Documentation is required for ${selectedLeaveType.Title}`);
    }

    return errors;
  }

  private onSubmit = async (): Promise<void> => {
    const validationErrors = this.validateForm();
    if (validationErrors.length > 0) {
      this.setState({ errors: validationErrors });
      return;
    }

    this.setState({ isSubmitting: true, errors: [], successMessage: '' });

    try {
      const request: ILeaveRequestCreate = {
        LeaveTypeId: this.state.formData.leaveTypeId,
        StartDate: this.state.formData.startDate,
        EndDate: this.state.formData.endDate,
        IsPartialDay: this.state.formData.isPartialDay,
        PartialDayHours: this.state.formData.partialDayHours,
        RequestComments: this.state.formData.comments,
        AttachmentURL: this.state.formData.attachmentUrl
      };

      // Validate against business rules
      const validation = await this.sharePointService.validateLeaveRequest(request);
      if (!validation.isValid) {
        this.setState({ errors: validation.errors, isSubmitting: false });
        return;
      }

      // Create the leave request
      await this.sharePointService.createLeaveRequest(request);

      this.setState({
        successMessage: 'Leave request submitted successfully!',
        isSubmitting: false,
        formData: {
          leaveTypeId: 0,
          startDate: new Date(),
          endDate: new Date(),
          isPartialDay: false,
          partialDayHours: undefined,
          comments: '',
          attachmentUrl: ''
        },
        selectedLeaveType: null
      });
    } catch (error) {
      console.error('Error submitting leave request:', error);
      this.setState({
        errors: ['Failed to submit leave request. Please try again.'],
        isSubmitting: false
      });
    }
  };

  private onReset = (): void => {
    this.setState({
      formData: {
        leaveTypeId: 0,
        startDate: new Date(),
        endDate: new Date(),
        isPartialDay: false,
        partialDayHours: undefined,
        comments: '',
        attachmentUrl: ''
      },
      selectedLeaveType: null,
      errors: [],
      successMessage: ''
    });
  };

  public render(): React.ReactElement<ILeaveRequestFormProps> {
    const { hasTeamsContext, userDisplayName } = this.props;
    const {
      leaveTypes,
      formData,
      isLoading,
      isSubmitting,
      errors,
      successMessage,
      selectedLeaveType
    } = this.state;

    if (isLoading) {
      return (
        <section className={`${styles.leaveRequestForm} ${hasTeamsContext ? styles.teams : ''}`}>
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Loading leave types..." />
          </Stack>
        </section>
      );
    }

    const leaveTypeOptions: IDropdownOption[] = LeaveTypeUtils.toDropdownOptions(leaveTypes);
    const totalDays = formData.startDate && formData.endDate ?
      CommonUtils.calculateBusinessDays(formData.startDate, formData.endDate) : 0;

    return (
      <section className={`${styles.leaveRequestForm} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="xxLarge" as="h1">Leave Request Form</Text>
            <Text variant="medium">Welcome, {escape(userDisplayName)}! Submit your leave request below.</Text>
          </Stack.Item>

          {errors.length > 0 && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline>
              <ul style={{ margin: 0, paddingLeft: 20 }}>
                {errors.map((error, index) => (
                  <li key={index}>{error}</li>
                ))}
              </ul>
            </MessageBar>
          )}

          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success}>
              {successMessage}
            </MessageBar>
          )}

          <Stack tokens={{ childrenGap: 15 }}>
            <Dropdown
              label="Leave Type *"
              placeholder="Select a leave type"
              options={leaveTypeOptions}
              selectedKey={formData.leaveTypeId || undefined}
              onChange={this.onLeaveTypeChange}
              disabled={isSubmitting}
              required
            />

            {selectedLeaveType && (
              <Stack.Item>
                <Text variant="small" styles={{ root: { color: '#666' } }}>
                  {selectedLeaveType.Description}
                  {selectedLeaveType.MaxDaysPerRequest &&
                    ` (Max: ${selectedLeaveType.MaxDaysPerRequest} days per request)`}
                  {selectedLeaveType.RequiresDocumentation && ' - Documentation required'}
                </Text>
              </Stack.Item>
            )}

            <Stack horizontal tokens={{ childrenGap: 15 }}>
              <Stack.Item grow>
                <DatePicker
                  label="Start Date *"
                  value={formData.startDate}
                  onSelectDate={this.onStartDateChange}
                  disabled={isSubmitting}
                  isRequired
                  minDate={new Date()}
                />
              </Stack.Item>
              <Stack.Item grow>
                <DatePicker
                  label="End Date *"
                  value={formData.endDate}
                  onSelectDate={this.onEndDateChange}
                  disabled={isSubmitting}
                  isRequired
                  minDate={formData.startDate}
                />
              </Stack.Item>
            </Stack>

            {totalDays > 0 && (
              <Stack.Item>
                <Label>Total Business Days: {totalDays}</Label>
              </Stack.Item>
            )}

            <Checkbox
              label="Partial Day Request"
              checked={formData.isPartialDay}
              onChange={this.onPartialDayChange}
              disabled={isSubmitting}
            />

            {formData.isPartialDay && (
              <TextField
                label="Hours *"
                type="number"
                value={formData.partialDayHours?.toString() || ''}
                onChange={this.onPartialDayHoursChange}
                disabled={isSubmitting}
                suffix="hours"
                min={0.5}
                max={8}
                step={0.5}
                required
              />
            )}

            <TextField
              label="Comments"
              multiline
              rows={3}
              value={formData.comments}
              onChange={this.onCommentsChange}
              disabled={isSubmitting}
              placeholder="Optional comments about your leave request"
            />

            {selectedLeaveType?.RequiresDocumentation && (
              <TextField
                label="Documentation URL *"
                value={formData.attachmentUrl}
                onChange={this.onAttachmentUrlChange}
                disabled={isSubmitting}
                placeholder="URL to supporting documentation"
                required
              />
            )}

            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text={isSubmitting ? "Submitting..." : "Submit Request"}
                onClick={this.onSubmit}
                disabled={isSubmitting}
                iconProps={isSubmitting ? { iconName: 'Sync' } : { iconName: 'Send' }}
              />
              <DefaultButton
                text="Reset Form"
                onClick={this.onReset}
                disabled={isSubmitting}
                iconProps={{ iconName: 'Clear' }}
              />
            </Stack>
          </Stack>
        </Stack>
      </section>
    );
  }
}
