# Created 2025 by Arleta Wanat
# This report searches for unique permissions across the entire SharePoint Online tenant
# Depending on size of your sites, it may take about a minute per each site collection
# If your tenant has a lot of sites and a lot of unique permissions, you may want to switch to another script 


# Import PnP PowerShell Module
Import-Module PnP.PowerShell

# Connect to SharePoint Admin Center
$AdminUrl = "https://<YourTenantName>-admin.sharepoint.com"
Connect-PnPOnline -Url $AdminUrl -Interactive

# Retrieve all site collections
$SiteCollections = Get-PnPTenantSite -IncludeOneDriveSites $false

# Initialize a report to store unique permissions details
$Report = @()

# Function to Check Unique Permissions
function Check-UniquePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Web
    )

    # Get all lists in the web
    $Lists = Get-PnPList -Web $Web

    foreach ($List in $Lists) {
        # Check if the list has unique permissions
        if ($List.HasUniqueRoleAssignments) {
            $Report += [PSCustomObject]@{
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
                $Report += [PSCustomObject]@{
                    "SiteUrl"   = $Web.Url
                    "Type"      = "Item/Document"
                    "Name"      = $Item.FieldValues["FileLeafRef"]
                    "Url"       = $Item.FieldValues["FileRef"]
                }
            }
        }
    }
}

# Loop through all site collections
foreach ($Site in $SiteCollections) {
    try {
        # Connect to each site collection
        Connect-PnPOnline -Url $Site.Url -Interactive

        # Get the root web of the site collection
        $RootWeb = Get-PnPWeb

        # Check unique permissions for the root web
        if ($RootWeb.HasUniqueRoleAssignments) {
            $Report += [PSCustomObject]@{
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
                $Report += [PSCustomObject]@{
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
    }
    catch {
        Write-Error "Error accessing site: $($Site.Url) - $_"
    }
}

# Export the results to a CSV file
$Report | Export-Csv -Path "UniquePermissionsReport.csv" -NoTypeInformation -Encoding UTF8

# Disconnect from PnP
Disconnect-PnPOnline

Write-Output "Unique Permissions Report generated: UniquePermissionsReport.csv"
