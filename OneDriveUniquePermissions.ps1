# Import PnP PowerShell Module
Import-Module PnP.PowerShell

# Define variables for authentication
$TenantId = "<Your-Tenant-ID>"
$ClientId = "<Your-Client-ID>"
$CertificatePath = "<Path-To-Your-PFX-Certificate>"
$CertificatePassword = "<Certificate-Password>"

# Authenticate using app-only authentication
Connect-PnPOnline -Tenant $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force)

# Retrieve all OneDrive sites (filtering by OneDrive URL pattern)
$OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites $true | Where-Object { $_.Url -like "https://*-my.sharepoint.com/personal/*" }

# Path to the output CSV file
$ReportFilePath = "OneDriveAllListsUniquePermissionsReport.csv"

# Initialize the CSV file with headers
@(
    "SiteUrl,Type,ListName,Url"
) | Out-File -FilePath $ReportFilePath -Encoding UTF8

# Function to Log Results to CSV
function Log-ToCsv {
    param (
        [Parameter(Mandatory = $true)]
        [object]$LogData
    )

    $LogData | ForEach-Object {
        "$($_.SiteUrl),$($_.Type),$($_.ListName),$($_.Url)" | Out-File -FilePath $ReportFilePath -Append -Encoding UTF8
    }
}

# Function to Check Unique Permissions for All Lists
function Check-UniquePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    # Connect to the OneDrive site
    Set-PnPContext -Url $SiteUrl

    # Initialize a local report for this OneDrive site
    $LocalReport = @()

    # Get all lists in the OneDrive site
    $Lists = Get-PnPList

    foreach ($List in $Lists) {
        # Check if the list has unique permissions
        if ($List.HasUniqueRoleAssignments) {
            $LocalReport += [PSCustomObject]@{
                "SiteUrl"   = $SiteUrl
                "Type"      = "List"
                "ListName"  = $List.Title
                "Url"       = $List.DefaultViewUrl
            }
        }

        # Get all items in the list (if it supports items)
        $Items = Get-PnPListItem -List $List -ErrorAction SilentlyContinue
        foreach ($Item in $Items) {
            if ($Item.HasUniqueRoleAssignments) {
                $LocalReport += [PSCustomObject]@{
                    "SiteUrl"   = $SiteUrl
                    "Type"      = "Item/Document"
                    "ListName"  = $List.Title
                    "Url"       = $Item.FieldValues["FileRef"]
                }
            }
        }
    }

    # Log the results to the CSV file
    if ($LocalReport.Count -gt 0) {
        Log-ToCsv -LogData $LocalReport
    }
}

# Loop through all OneDrive sites
foreach ($Site in $OneDriveSites) {
    try {
        # Check unique permissions for all lists in the OneDrive site
        Check-UniquePermissions -SiteUrl $Site.Url
    }
    catch {
        Write-Error "Error accessing OneDrive site: $($Site.Url) - $_"
    }
}

# Disconnect from PnP
Disconnect-PnPOnline

Write-Output "OneDrive Unique Permissions Report for All Lists generated incrementally: $ReportFilePath"
