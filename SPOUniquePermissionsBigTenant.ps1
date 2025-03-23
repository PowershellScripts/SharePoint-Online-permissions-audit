# Created 2025 by Arleta Wanat
# This report searches for unique permissions across the entire SharePoint Online tenant
# Depending on size of your sites, it may take a few minutes per each site collection

# Import PnP PowerShell Module
Import-Module PnP.PowerShell

# Variables for app-only authentication
$AdminUrl = "https://<YourTenantName>-admin.sharepoint.com"
$TenantId = "<Your-Tenant-ID>"
$ClientId = "<Your-Client-ID>"
$CertificatePath = "<Path-To-Your-Certificate>"
$CertificatePassword = "<Certificate-Password>"

# Authenticate using app-only authentication
Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force)

# Retrieve all site collections
$SiteCollections = Get-PnPTenantSite -IncludeOneDriveSites $false

# Path to the output CSV file
$ReportFilePath = "UniquePermissionsReport.csv"

# Initialize the CSV file with headers
@(
    "SiteUrl,Type,Name,Url"
) | Out-File -FilePath $ReportFilePath -Encoding UTF8

# Function to Log Results to CSV
function Log-ToCsv {
    param (
        [Parameter(Mandatory = $true)]
        [object]$LogData
    )

    $LogData | ForEach-Object {
        "$($_.SiteUrl),$($_.Type),$($_.Name),$($_.Url)" | Out-File -FilePath $ReportFilePath -Append -Encoding UTF8
    }
}

# Function to Check Unique Permissions
function Check-UniquePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Web
    )

    # Initialize a local report for this site collection
    $LocalReport = @()

    # Get all lists in the web
    $Lists = Get-PnPList -Web $Web

    foreach ($List in $Lists) {
        # Check if the list has unique permissions
        if ($List.HasUniqueRoleAssignments) {
            $LocalReport += [PSCustomObject]@{
                "SiteUrl"   = $Web.Url
                "Type"      = "List"
                "Name"      = $List.Title
                "Url"       = $List.DefaultViewUrl
            }
        }

        # Get all list items or documents
        $Items = Get-PnPListItem -List $List -Web $Web
        foreach ($Item in $Items) {
            if ($Item.HasUniqueRoleAssignments) {
                $LocalReport += [PSCustomObject]@{
                    "SiteUrl"   = $Web.Url
                    "Type"      = "Item/Document"
                    "Name"      = $Item.FieldValues["FileLeafRef"]
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

# Loop through all site collections
foreach ($Site in $SiteCollections) {
    try {
        # Connect to each site collection using app-only authentication
        Connect-PnPOnline -Url $Site.Url -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force)

        # Get the root web of the site collection
        $RootWeb = Get-PnPWeb

        # Initialize a local report for the site collection
        $SiteReport = @()

        # Check unique permissions for the root web
        if ($RootWeb.HasUniqueRoleAssignments) {
            $SiteReport += [PSCustomObject]@{
                "SiteUrl"   = $RootWeb.Url
                "Type"      = "Site Collection Root"
                "Name"      = $RootWeb.Title
                "Url"       = $RootWeb.Url
            }
        }

        # Recursively check all subsites
        $Subsites = Get-PnPSubWeb -Recurse
        foreach ($Subsite in $Subsites) {
            if ($Subsite.HasUniqueRoleAssignments) {
                $SiteReport += [PSCustomObject]@{
                    "SiteUrl"   = $Subsite.Url
                    "Type"      = "Subsite"
                    "Name"      = $Subsite.Title
                    "Url"       = $Subsite.Url
                }
            }

            # Check unique permissions in lists and items of the subsite
            Check-UniquePermissions -Web $Subsite
        }

        # Check unique permissions in lists and items of the root web
        Check-UniquePermissions -Web $RootWeb

        # Log the site collection report to the CSV
        if ($SiteReport.Count -gt 0) {
            Log-ToCsv -LogData $SiteReport
        }
    }
    catch {
        Write-Error "Error accessing site: $($Site.Url) - $_"
    }
}

# Disconnect from PnP
Disconnect-PnPOnline

Write-Output "Unique Permissions Report generated incrementally: $ReportFilePath"
