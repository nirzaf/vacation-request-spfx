# Team Leave & Vacation Request SPFx Solution

## Summary

A comprehensive SharePoint Framework (SPFx) solution for managing team leave requests, vacation planning, and HR administration. This solution provides a complete workflow for leave management with modern React-based web parts, Microsoft Graph integration, and advanced analytics.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/)
- [Office UI Fabric React](https://developer.microsoft.com/en-us/fluentui)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- SharePoint Online environment
- Node.js (v16 or later)
- SharePoint Framework development environment
- PnP PowerShell module
- Appropriate permissions for SharePoint and Microsoft Graph

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| vacation-request-spfx | M.F.M Fazrin ([@nirzaf](https://github.com/nirzaf)) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | December 2024   | Initial release with complete leave management system |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## 🌟 Features

### 📝 Leave Request Form
- **Intuitive Request Submission**: Easy-to-use form with validation and real-time feedback
- **Dynamic Leave Type Selection**: Configurable leave types with specific rules and policies
- **Manager Auto-Detection**: Automatic manager assignment using Microsoft Graph
- **Partial Day Support**: Handle partial day leave requests with hour tracking
- **Document Attachments**: Support for required documentation
- **Real-time Validation**: Business rule validation and conflict detection

### 📅 Team Calendar
- **FullCalendar Integration**: Rich calendar interface with multiple view modes
- **Team Visibility**: View all team leave requests in calendar format
- **Filtering & Search**: Advanced filtering by leave type, department, and status
- **Conflict Detection**: Visual identification of overlapping requests
- **Export Capabilities**: Export calendar data to CSV

### 📊 Leave History & Tracking
- **Personal Dashboard**: Track individual leave request history and status
- **Leave Balance Monitoring**: Real-time balance tracking with expiration warnings
- **Request Modification**: Edit pending requests with approval workflow
- **Status Tracking**: Complete audit trail of request lifecycle

### 🏢 Administration Dashboard
- **Bulk Approval Workflows**: Efficient processing of multiple requests
- **Advanced Analytics**: Comprehensive reporting and trend analysis
- **Team Management**: Department-wise leave tracking and planning
- **Policy Management**: Configure leave types, rules, and allowances
- **Notification System**: Automated email notifications via Microsoft Graph

## Minimal Path to Awesome

1. **Clone this repository**
   ```bash
   git clone https://github.com/nirzaf/vacation-request-spfx.git
   cd vacation-request-spfx
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the solution**
   ```bash
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```

4. **Deploy to SharePoint**
   ```powershell
   # Run the deployment script
   .\scripts\Deploy-Solution.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -AppCatalogUrl "https://yourtenant.sharepoint.com/sites/appcatalog" -IncludeTestData
   ```

5. **Setup SharePoint Lists**
   ```powershell
   # Create required lists and sample data
   .\scripts\Setup-SharePointLists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -IncludeTestData
   ```

## Web Parts Included

| Web Part | Description | Target Users |
|----------|-------------|--------------|
| **LeaveRequestForm** | Submit and manage leave requests | All employees |
| **TeamCalendar** | View team leave requests in calendar format | All employees, Managers |
| **LeaveHistory** | Personal leave tracking and balance management | All employees |
| **LeaveAdministration** | HR and manager dashboard with analytics | HR, Managers |

## SharePoint Lists

| List | Purpose | Key Fields |
|------|---------|------------|
| **LeaveRequests** | Central repository for all leave requests | Requester, LeaveType, StartDate, EndDate, ApprovalStatus |
| **LeaveTypes** | Configurable leave type definitions | Title, RequiresApproval, MaxDaysPerRequest, ColorCode |
| **LeaveBalances** | Employee leave balance tracking | Employee, LeaveType, TotalAllowance, UsedDays, RemainingDays |

## Configuration

### API Permissions
Approve the following permissions in SharePoint Admin Center:
- Microsoft Graph: User.Read
- Microsoft Graph: User.ReadBasic.All
- Microsoft Graph: Calendars.ReadWrite
- Microsoft Graph: Mail.Send
- Microsoft Graph: Directory.Read.All

### Setup Steps
1. Deploy the solution package to your App Catalog
2. Install the solution on your target site
3. Approve API permissions in SharePoint Admin Center
4. Run the SharePoint list setup script
5. Configure leave types and employee balances
6. Add web parts to your SharePoint pages

## Architecture

### Services Layer
- **SharePointService** - SharePoint list operations and data management
- **GraphService** - Microsoft Graph API integration for calendar and notifications
- **NotificationService** - Email notification management
- **ValidationService** - Business rule validation and conflict detection

### Models
- **ILeaveRequest** - Leave request data structure
- **ILeaveType** - Leave type configuration
- **ILeaveBalance** - Employee balance tracking
- **Common utilities** - Shared functionality and validation

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Office UI Fabric React](https://developer.microsoft.com/en-us/fluentui)
- [FullCalendar Documentation](https://fullcalendar.io/docs)
- [PnP PowerShell](https://pnp.github.io/powershell/)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)
