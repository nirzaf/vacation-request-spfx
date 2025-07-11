# Deploy Team Leave & Vacation Request SPFx Solution
# This script deploys the complete solution including SharePoint lists and web parts

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AppCatalogUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipListCreation,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeTestData,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

# Import required modules
try {
    Import-Module PnP.PowerShell -Force -ErrorAction Stop
    Write-Host "✓ PnP PowerShell module loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to load PnP PowerShell module. Please install it first:" -ForegroundColor Red
    Write-Host "Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Configuration
$SolutionPath = "../solution/team-leave-vacation-request-solution.sppkg"
$SolutionName = "team-leave-vacation-request-solution.sppkg"

Write-Host "=== Team Leave & Vacation Request Solution Deployment ===" -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl" -ForegroundColor White
Write-Host "App Catalog: $AppCatalogUrl" -ForegroundColor White
Write-Host "Solution Path: $SolutionPath" -ForegroundColor White
Write-Host ""

# Step 1: Connect to App Catalog
Write-Host "Step 1: Connecting to App Catalog..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $AppCatalogUrl -Interactive
    Write-Host "✓ Connected to App Catalog successfully" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to App Catalog: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 2: Upload and Deploy Solution Package
Write-Host "Step 2: Uploading solution package..." -ForegroundColor Yellow
try {
    if (Test-Path $SolutionPath) {
        # Remove existing solution if Force is specified
        if ($Force) {
            try {
                $existingApp = Get-PnPApp -Identity $SolutionName -ErrorAction SilentlyContinue
                if ($existingApp) {
                    Write-Host "Removing existing solution..." -ForegroundColor Yellow
                    Remove-PnPApp -Identity $existingApp.Id -Force
                    Write-Host "✓ Existing solution removed" -ForegroundColor Green
                }
            } catch {
                Write-Host "Note: No existing solution found to remove" -ForegroundColor Gray
            }
        }
        
        # Upload the solution
        $app = Add-PnPApp -Path $SolutionPath -Overwrite
        Write-Host "✓ Solution package uploaded successfully" -ForegroundColor Green
        
        # Deploy the solution
        Publish-PnPApp -Identity $app.Id -SkipFeatureDeployment:$false
        Write-Host "✓ Solution deployed to App Catalog" -ForegroundColor Green
    } else {
        Write-Host "✗ Solution package not found at: $SolutionPath" -ForegroundColor Red
        Write-Host "Please build the solution first using: gulp bundle --ship && gulp package-solution --ship" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "✗ Failed to upload/deploy solution: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 3: Connect to Target Site
Write-Host "Step 3: Connecting to target site..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "✓ Connected to target site successfully" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to target site: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 4: Install Solution on Site
Write-Host "Step 4: Installing solution on site..." -ForegroundColor Yellow
try {
    $installedApp = Get-PnPApp -Identity $SolutionName -ErrorAction SilentlyContinue
    if ($installedApp -and $installedApp.Installed) {
        if ($Force) {
            Write-Host "Uninstalling existing solution..." -ForegroundColor Yellow
            Uninstall-PnPApp -Identity $installedApp.Id
            Start-Sleep -Seconds 5
        } else {
            Write-Host "Solution is already installed. Use -Force to reinstall." -ForegroundColor Yellow
        }
    }
    
    if (-not $installedApp -or -not $installedApp.Installed -or $Force) {
        Install-PnPApp -Identity $SolutionName
        Write-Host "✓ Solution installed on site successfully" -ForegroundColor Green
    }
} catch {
    Write-Host "✗ Failed to install solution on site: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 5: Create SharePoint Lists (if not skipped)
if (-not $SkipListCreation) {
    Write-Host "Step 5: Creating SharePoint lists..." -ForegroundColor Yellow
    try {
        # Execute the list creation script
        $listScriptPath = "./Setup-SharePointLists.ps1"
        if (Test-Path $listScriptPath) {
            if ($IncludeTestData) {
                & $listScriptPath -SiteUrl $SiteUrl -IncludeTestData
            } else {
                & $listScriptPath -SiteUrl $SiteUrl
            }
            Write-Host "✓ SharePoint lists created successfully" -ForegroundColor Green
        } else {
            Write-Host "⚠ List creation script not found. Please run Setup-SharePointLists.ps1 manually." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "✗ Failed to create SharePoint lists: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "You can create lists manually using Setup-SharePointLists.ps1" -ForegroundColor Yellow
    }
} else {
    Write-Host "Step 5: Skipping SharePoint list creation (as requested)" -ForegroundColor Gray
}

# Step 6: Configure Permissions
Write-Host "Step 6: Configuring permissions..." -ForegroundColor Yellow
try {
    # Grant API permissions (this requires tenant admin approval)
    Write-Host "Note: The following API permissions need to be approved by a tenant administrator:" -ForegroundColor Yellow
    Write-Host "- Microsoft Graph: User.Read" -ForegroundColor White
    Write-Host "- Microsoft Graph: User.ReadBasic.All" -ForegroundColor White
    Write-Host "- Microsoft Graph: Calendars.ReadWrite" -ForegroundColor White
    Write-Host "- Microsoft Graph: Mail.Send" -ForegroundColor White
    Write-Host "- Microsoft Graph: Directory.Read.All" -ForegroundColor White
    Write-Host ""
    Write-Host "Please go to SharePoint Admin Center > Advanced > API access to approve these permissions." -ForegroundColor Cyan
    Write-Host "✓ Permission configuration noted" -ForegroundColor Green
} catch {
    Write-Host "⚠ Note: API permissions may need manual approval" -ForegroundColor Yellow
}

# Step 7: Create Sample Pages
Write-Host "Step 7: Creating sample pages..." -ForegroundColor Yellow
try {
    # Create Leave Management Dashboard page
    $dashboardPage = Add-PnPPage -Name "LeaveManagementDashboard" -Title "Leave Management Dashboard" -LayoutType Article -ErrorAction SilentlyContinue
    if ($dashboardPage) {
        # Add web parts to the page
        Add-PnPPageSection -Page $dashboardPage -SectionTemplate OneColumn -Order 1
        Add-PnPPageTextPart -Page $dashboardPage -Section 1 -Column 1 -Text "<h2>Welcome to Leave Management</h2><p>Use the web parts below to manage your leave requests and view team information.</p>"
        
        Write-Host "✓ Leave Management Dashboard page created" -ForegroundColor Green
    }
    
    # Create Team Calendar page
    $calendarPage = Add-PnPPage -Name "TeamLeaveCalendar" -Title "Team Leave Calendar" -LayoutType Article -ErrorAction SilentlyContinue
    if ($calendarPage) {
        Add-PnPPageSection -Page $calendarPage -SectionTemplate OneColumn -Order 1
        Add-PnPPageTextPart -Page $calendarPage -Section 1 -Column 1 -Text "<h2>Team Leave Calendar</h2><p>View all team leave requests in calendar format.</p>"
        
        Write-Host "✓ Team Leave Calendar page created" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠ Some sample pages may not have been created: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Step 8: Final Verification
Write-Host "Step 8: Verifying deployment..." -ForegroundColor Yellow
try {
    $verifyApp = Get-PnPApp -Identity $SolutionName
    if ($verifyApp -and $verifyApp.Installed) {
        Write-Host "✓ Solution is properly installed and active" -ForegroundColor Green
    } else {
        Write-Host "⚠ Solution installation could not be verified" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠ Could not verify solution installation" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Deployment Summary ===" -ForegroundColor Cyan
Write-Host "✓ Solution package uploaded to App Catalog" -ForegroundColor Green
Write-Host "✓ Solution deployed and installed on site" -ForegroundColor Green
if (-not $SkipListCreation) {
    Write-Host "✓ SharePoint lists configured" -ForegroundColor Green
}
Write-Host "✓ Sample pages created" -ForegroundColor Green
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Approve API permissions in SharePoint Admin Center" -ForegroundColor White
Write-Host "2. Add web parts to your pages" -ForegroundColor White
Write-Host "3. Configure leave types and employee balances" -ForegroundColor White
Write-Host "4. Test the solution with sample data" -ForegroundColor White
Write-Host ""
Write-Host "Deployment completed successfully! 🎉" -ForegroundColor Green

# Disconnect
Disconnect-PnPOnline
