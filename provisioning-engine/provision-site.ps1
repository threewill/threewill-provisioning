<#     
    .SYNOPSIS
    Applies provisioning template to a site or team based on the provided parameters

    .DESCRIPTION
    This script will apply a provisioning template to a site based on the parameters passed in.
    The script allows for pre- and post-provisioning tasks to be applied, as well as defining
    a provisioning template to invoke for the site.  Alternatively, a provisioning script can 
    be provided to handle the complete provisioning process.

    .PARAMETER ConfigFile
    Relative path of the configuration file to use (e.g. config\boinga.json)

    .EXAMPLE
    .\provision-site.ps1 -ConfigFile config\boinga.json -Site "pwc-templatetest" -SiteTitle "PwC Template Test" -SiteDescription "Test Description"

    .NOTES
    - Dependencies: 
        PnP.PowerShell cmdlets, version 1.9 or higher (Install-Module -Name PnP.PowerShell)
        Teams cmdlets, version 1.01 or higher (Install-Module -Name MicrosoftTeams)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,
    [Parameter(Mandatory = $true)]
    [string]$Site,    
    [Parameter(Mandatory = $true)]
    [string]$SiteTitle,
    [Parameter(Mandatory = $true)]
    [ValidateSet("IntranetSpokeSite", "CommunicationSite", "TeamSite", "MSTeam")]    
    [string]$EntityType,
    [switch]$SkipGetCredentials,
    [switch]$BatchMode
)

# Make the helper functions available to script
. "helpers/helper-functions.ps1"

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"

# only write out stuff if we aren't being called by another script
if ($BatchMode.IsPresent -eq $false)
{
    Write-Output "".PadRight(50, "=")
    Write-Output "provision-site.ps1"
    Write-Output "  params: Site:$($Site)"
    Write-Output "  params: ConfigFile:$($ConfigFile),Site:$($Site),EntityType:$($EntityType)"
    Write-Output "".PadRight(50, "=")
}

# Get configuration
$config = Get-Content $ConfigFile | Out-String | ConvertFrom-Json

# only write out stuff if we aren't being called by another script
if ($BatchMode.IsPresent -eq $false)
{
    Write-Output "Start Config Values ".PadRight(50, "*")
    Write-Output $config
    Write-Output "End Config Values ".PadRight(50, "*")
}

# this allows us to set $cred before executing script and not be prompted
if ($SkipGetCredentials.IsPresent -eq $false)
{
    Write-Output "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
}

$siteUrl = Get-UrlByEntityType $EntityType $Site $config

# only check if the site exists if we are not called from another script
# otherwise assume the parent script has verified the site exists
if ($BatchMode.IsPresent -eq $false)
{
    Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred

    try
    {
        # Check if Site Already Exists
        $existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore

        if ($null -eq $existingSite)
        {
            Write-Warning "The site '$siteUrl' does not exist.  Cannot apply provisioning."
            return
        }
    }
    finally
    {
        Disconnect-PnPOnline
    }
}

$entityTypeConfigKey = switch ($EntityType)
{
    "CommunicationSite" { "communicationSiteProvisioning" }
    "TeamSite" { "teamSiteProvisioning" }
    "IntranetSpokeSite" { "intranetSpokeSiteProvisioning" }
    Default { "msTeamsProvisioning" }
}

$provisioningScript = Get-NestedMember $config "plugins.$entityTypeConfigKey.provisioningScript"
if ($null -ne $provisioningScript)
{
    # Apply provisioning script if we have it
    . $provisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl
}
else 
{ 
    # Get connection to the site
    if ($EntityType -ne "MSTeam")
    {
        Write-Output "Connect-PnpOnline"
        Connect-PnPOnline -Url $siteUrl -Credentials $global:cred
    }
    else
    {
        # Connect to Teams
        Write-Output "Connect-MicrosoftTeams"
        Connect-MicrosoftTeams -Credential $global:cred
    }

    #Pre-Provisioning
    $preProvisioningScript = Get-NestedMember $config "plugins.$entityTypeConfigKey.preProvisioningTask"
    if ($null -ne $preProvisioningScript)
    {
        if (Test-Path -Path $preProvisioningScript)
        {
            Write-Host "Running pre-provisioning script '$preProvisioningScript'"

            # Run the pre-provisioning script, if there is one
            . $preProvisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl -ConfigFile $ConfigFile
        }
        else
        {
            Write-Warning "Could not find pre-provisioning script '$preProvisioningScript'"
        }

        
    }

    #Automatic Template Provisioning
    if ($EntityType -ne "MSTeam")
    {
        $provisioningTemplate = Get-NestedMember $config "plugins.$entityTypeConfigKey.provisioningTemplate"
        if ($null -ne $provisioningTemplate)
        {
            if (Test-Path -Path $provisioningTemplate)
            {
                Write-Host "Inovking PnP template '$provisioningTemplate'"

                # Invoke the provisioning template, if there is one
                Invoke-PnPSiteTemplate -Path $provisioningTemplate
            }
            else
            {
                Write-Warning "Could not find provisioning template '$provisioningTemplate'"
            }
            
        }
    }

    #Post-Provisioning
    $postProvisioningScript = Get-NestedMember $config "plugins.$entityTypeConfigKey.postProvisioningTask"
    if ($null -ne $postProvisioningScript)
    {
        if (Test-Path -Path $postProvisioningScript)
        {
            Write-Host "Running post-provisioning script '$postProvisioningScript'"

            # Run the post-provisioning script, if there is one
            . $postProvisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl -ConfigFile $ConfigFile
        }
        else
        {
            Write-Warning "Could not find post-provisioning script '$postProvisioningScript'"
        }
    }

    #Permissions defaults
    if ($EntityType -ne "MSTeam")
    {
        $permissions = Get-NestedMember $config "plugins.$entityTypeConfigKey.permissions"
        if ($null -ne $permissions)
        {
            $connection = Connect-PnPOnline -Url $siteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop

            $owners = $permissions.owners
            if ($null -ne $owners)
            {
                $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup
                foreach ($o in $owners) {
                    Add-PnPGroupMember -EmailAddress $o -Group $ownerGroup.Id -Connection $connection
                    Write-Host "$o added to Site Owners"
                }
            }

            $members = $permissions.members
            if ($null -ne $members)
            {
                $memberGroup = Get-PnPGroup -AssociatedMemberGroup
                foreach ($m in $members) {
                    Add-PnPGroupMember -EmailAddress $m -Group $memberGroup.Id -Connection $connection
                    Write-Host "$m added to Site Members"
                }
            }

            $visitors = $permissions.visitors
            if ($null -ne $visitors)
            {
                $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup
                foreach ($v in $visitors) {
                    Add-PnPGroupMember -EmailAddress $v -Group $visitorGroup.Id -Connection $connection
                    Write-Host "$v added to Site Visitors"
                }
            }
        }
    }
}

# Make sure all open connections are closed
Disconnect-OpenConnections
 