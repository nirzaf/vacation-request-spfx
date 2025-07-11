# Setup SharePoint Lists for Vacation Request Solution
# This script creates the required SharePoint lists and configures their fields

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeTestData
)

# Import required modules
Import-Module PnP.PowerShell -Force

Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Green
Connect-PnPOnline -Url $SiteUrl -Interactive

try {
    # Create Leave Types List
    Write-Host "Creating Leave Types list..." -ForegroundColor Yellow
    
    $leaveTypesList = Get-PnPList -Identity "LeaveTypes" -ErrorAction SilentlyContinue
    if ($null -eq $leaveTypesList) {
        New-PnPList -Title "LeaveTypes" -Template GenericList -Url "Lists/LeaveTypes"
        
        # Add custom fields to Leave Types list
        Add-PnPField -List "LeaveTypes" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "RequiresApproval" -InternalName "RequiresApproval" -Type Boolean -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "MaxDaysPerRequest" -InternalName "MaxDaysPerRequest" -Type Number -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "RequiresDocumentation" -InternalName "RequiresDocumentation" -Type Boolean -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "ColorCode" -InternalName "ColorCode" -Type Text -AddToDefaultView
        Add-PnPField -List "LeaveTypes" -DisplayName "PolicyURL" -InternalName "PolicyURL" -Type URL -AddToDefaultView
        
        Write-Host "Leave Types list created successfully!" -ForegroundColor Green
    } else {
        Write-Host "Leave Types list already exists." -ForegroundColor Yellow
    }

    # Create Leave Requests List
    Write-Host "Creating Leave Requests list..." -ForegroundColor Yellow
    
    $leaveRequestsList = Get-PnPList -Identity "LeaveRequests" -ErrorAction SilentlyContinue
    if ($null -eq $leaveRequestsList) {
        New-PnPList -Title "LeaveRequests" -Template GenericList -Url "Lists/LeaveRequests"
        
        # Add custom fields to Leave Requests list
        Add-PnPField -List "LeaveRequests" -DisplayName "Requester" -InternalName "Requester" -Type User -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "EmployeeID" -InternalName "EmployeeID" -Type Text -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "Department" -InternalName "Department" -Type Text -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "Manager" -InternalName "Manager" -Type User -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "LeaveType" -InternalName "LeaveType" -Type Lookup -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "StartDate" -InternalName "StartDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "EndDate" -InternalName "EndDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "TotalDays" -InternalName "TotalDays" -Type Calculated -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "IsPartialDay" -InternalName "IsPartialDay" -Type Boolean -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "PartialDayHours" -InternalName "PartialDayHours" -Type Number -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "RequestComments" -InternalName "RequestComments" -Type Note -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "ApprovalStatus" -InternalName "ApprovalStatus" -Type Choice -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "ApprovalDate" -InternalName "ApprovalDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "ApprovalComments" -InternalName "ApprovalComments" -Type Note -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "SubmissionDate" -InternalName "SubmissionDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "AttachmentURL" -InternalName "AttachmentURL" -Type URL -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "WorkflowInstanceID" -InternalName "WorkflowInstanceID" -Type Text -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "NotificationsSent" -InternalName "NotificationsSent" -Type Boolean -AddToDefaultView
        Add-PnPField -List "LeaveRequests" -DisplayName "CalendarEventID" -InternalName "CalendarEventID" -Type Text -AddToDefaultView
        
        # Configure lookup field
        Set-PnPField -List "LeaveRequests" -Identity "LeaveType" -Values @{LookupList="LeaveTypes"; LookupField="Title"}
        
        # Configure choice field
        Set-PnPField -List "LeaveRequests" -Identity "ApprovalStatus" -Values @{Choices=@("Pending","Approved","Rejected","Cancelled"); DefaultValue="Pending"}
        
        # Configure calculated field for TotalDays
        Set-PnPField -List "LeaveRequests" -Identity "TotalDays" -Values @{Formula="=[EndDate]-[StartDate]+1"}
        
        Write-Host "Leave Requests list created successfully!" -ForegroundColor Green
    } else {
        Write-Host "Leave Requests list already exists." -ForegroundColor Yellow
    }

    # Create Leave Balances List
    Write-Host "Creating Leave Balances list..." -ForegroundColor Yellow
    
    $leaveBalancesList = Get-PnPList -Identity "LeaveBalances" -ErrorAction SilentlyContinue
    if ($null -eq $leaveBalancesList) {
        New-PnPList -Title "LeaveBalances" -Template GenericList -Url "Lists/LeaveBalances"
        
        # Add custom fields to Leave Balances list
        Add-PnPField -List "LeaveBalances" -DisplayName "Employee" -InternalName "Employee" -Type User -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "LeaveType" -InternalName "LeaveType" -Type Lookup -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "TotalAllowance" -InternalName "TotalAllowance" -Type Number -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "UsedDays" -InternalName "UsedDays" -Type Number -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "RemainingDays" -InternalName "RemainingDays" -Type Calculated -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "CarryOverDays" -InternalName "CarryOverDays" -Type Number -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "EffectiveDate" -InternalName "EffectiveDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "LeaveBalances" -DisplayName "ExpirationDate" -InternalName "ExpirationDate" -Type DateTime -AddToDefaultView
        
        # Configure lookup field
        Set-PnPField -List "LeaveBalances" -Identity "LeaveType" -Values @{LookupList="LeaveTypes"; LookupField="Title"}
        
        # Configure calculated field for RemainingDays
        Set-PnPField -List "LeaveBalances" -Identity "RemainingDays" -Values @{Formula="=[TotalAllowance]+[CarryOverDays]-[UsedDays]"}
        
        Write-Host "Leave Balances list created successfully!" -ForegroundColor Green
    } else {
        Write-Host "Leave Balances list already exists." -ForegroundColor Yellow
    }

    # Add default Leave Types if requested
    if ($IncludeTestData) {
        Write-Host "Adding default leave types..." -ForegroundColor Yellow
        
        $defaultLeaveTypes = @(
            @{Title="Annual Leave"; Description="Standard annual vacation leave"; IsActive=$true; RequiresApproval=$true; MaxDaysPerRequest=30; RequiresDocumentation=$false; ColorCode="#4CAF50"},
            @{Title="Sick Leave"; Description="Medical leave for illness"; IsActive=$true; RequiresApproval=$false; MaxDaysPerRequest=5; RequiresDocumentation=$true; ColorCode="#FF9800"},
            @{Title="Personal Leave"; Description="Personal time off"; IsActive=$true; RequiresApproval=$true; MaxDaysPerRequest=10; RequiresDocumentation=$false; ColorCode="#2196F3"},
            @{Title="Maternity/Paternity Leave"; Description="Family leave for new parents"; IsActive=$true; RequiresApproval=$true; MaxDaysPerRequest=90; RequiresDocumentation=$true; ColorCode="#E91E63"},
            @{Title="Emergency Leave"; Description="Urgent personal matters"; IsActive=$true; RequiresApproval=$false; MaxDaysPerRequest=3; RequiresDocumentation=$false; ColorCode="#F44336"}
        )
        
        foreach ($leaveType in $defaultLeaveTypes) {
            $existingItem = Get-PnPListItem -List "LeaveTypes" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($leaveType.Title)</Value></Eq></Where></Query></View>"
            
            if ($existingItem.Count -eq 0) {
                Add-PnPListItem -List "LeaveTypes" -Values $leaveType
                Write-Host "Added leave type: $($leaveType.Title)" -ForegroundColor Green
            } else {
                Write-Host "Leave type already exists: $($leaveType.Title)" -ForegroundColor Yellow
            }
        }
    }

    # Create custom views
    Write-Host "Creating custom views..." -ForegroundColor Yellow
    
    # Leave Requests views
    $myRequestsView = Get-PnPView -List "LeaveRequests" -Identity "My Requests" -ErrorAction SilentlyContinue
    if ($null -eq $myRequestsView) {
        Add-PnPView -List "LeaveRequests" -Title "My Requests" -Query "<Where><Eq><FieldRef Name='Requester'/><Value Type='Integer'><UserID/></Value></Eq></Where>" -Fields "Title","LeaveType","StartDate","EndDate","ApprovalStatus","SubmissionDate"
    }
    
    $pendingApprovalsView = Get-PnPView -List "LeaveRequests" -Identity "Pending Approvals" -ErrorAction SilentlyContinue
    if ($null -eq $pendingApprovalsView) {
        Add-PnPView -List "LeaveRequests" -Title "Pending Approvals" -Query "<Where><Eq><FieldRef Name='ApprovalStatus'/><Value Type='Choice'>Pending</Value></Eq></Where>" -Fields "Title","Requester","LeaveType","StartDate","EndDate","SubmissionDate"
    }

    # Leave Balances views
    $myBalancesView = Get-PnPView -List "LeaveBalances" -Identity "My Balances" -ErrorAction SilentlyContinue
    if ($null -eq $myBalancesView) {
        Add-PnPView -List "LeaveBalances" -Title "My Balances" -Query "<Where><Eq><FieldRef Name='Employee'/><Value Type='Integer'><UserID/></Value></Eq></Where>" -Fields "LeaveType","TotalAllowance","UsedDays","RemainingDays","ExpirationDate"
    }

    Write-Host "SharePoint lists setup completed successfully!" -ForegroundColor Green
    Write-Host "Created lists:" -ForegroundColor Cyan
    Write-Host "- LeaveTypes" -ForegroundColor White
    Write-Host "- LeaveRequests" -ForegroundColor White
    Write-Host "- LeaveBalances" -ForegroundColor White

} catch {
    Write-Host "Error occurred during setup: $($_.Exception.Message)" -ForegroundColor Red
    throw
} finally {
    Disconnect-PnPOnline
}
