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
    [ValidateSet("IntranetSpokeSite", "CommunicationSite", "TeamSite", "TeamSiteWithoutM365Group", "MSTeam")]    
    [string]$EntityType,
    [switch]$SkipGetCredentials,
    [switch]$BatchMode,
    [string]$SiteType
)

# Make the helper functions available to script
. "helpers/helper-functions.ps1"

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"

# only write out stuff if we aren't being called by another script
if ($BatchMode.IsPresent -eq $false)
{
    Write-Log "".PadRight(50, "=")
    Write-Log "provision-site.ps1"
    Write-Log "  params: Site:$($Site)"
    Write-Log "  params: ConfigFile:$($ConfigFile),Site:$($Site),EntityType:$($EntityType)"
    Write-Log "".PadRight(50, "=") -WriteNewLine
}

# Get configuration
$config = Get-Content $ConfigFile | Out-String | ConvertFrom-Json

# only write out stuff if we aren't being called by another script
if ($BatchMode.IsPresent -eq $false)
{
    Write-Log "Start Config Values ".PadRight(50, "*")
    Write-Log $config
    Write-Log "End Config Values ".PadRight(50, "*") -WriteNewLine
}

# this allows us to set $cred before executing script and not be prompted
if ($SkipGetCredentials.IsPresent -eq $false)
{
    Write-Log "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
	global:siteConn = Connect-PnPOnline -Url $siteUrl -Interactive -ReturnConnection -ErrorAction Stop
}

$siteUrl = Get-UrlByEntityType $EntityType $Site $config
$global:siteConn = Connect-PnPOnline -Url $siteUrl -Interactive -ReturnConnection -ErrorAction Stop

# only check if the site exists if we are not called from another script
# otherwise assume the parent script has verified the site exists
if ($BatchMode.IsPresent -eq $false)
{
    #Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred

    try
    {
        # Check if Site Already Exists
        $existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore

        if ($null -eq $existingSite)
        {
            Write-Log "The site '$siteUrl' does not exist.  Cannot apply provisioning." -Level Warn
            return
        }
    }
    finally
    {
        #Disconnect-PnPOnline
    }
}

$entityTypeConfigKey = switch ($EntityType)
{
    "CommunicationSite" { "communicationSiteProvisioning" }
    "TeamSite" { "teamSiteProvisioning" }
    "TeamSiteWithoutM365Group" { "teamSiteWithoutM365GroupProvisioning" }
    "IntranetSpokeSite" { "intranetSpokeSiteProvisioning" }
    Default { "msTeamsProvisioning" }
}

$provisioningScript = Get-NestedMember $config "plugins.$entityTypeConfigKey.provisioningScript"
if ($null -ne $provisioningScript)
{
    # Apply provisioning script if we have it
    . $provisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl -SiteType $SiteType -ConfigFile $ConfigFile
}
else 
{ 
    # Get connection to the site
    if ($EntityType -ne "MSTeam")
    {
        Write-Log "[$siteUrl] Connect-PnpOnline" -WriteToHost
        #Connect-PnPOnline -Url $siteUrl -Credentials $global:cred
    }
    else
    {
        # Connect to Teams
        Write-Log "[$siteUrl] Connect-MicrosoftTeams" -WriteToHost
        Connect-MicrosoftTeams -Credential $global:cred
    }

    #Pre-Provisioning
    $preProvisioningScript = Get-NestedMember $config "plugins.$entityTypeConfigKey.preProvisioningTask"
    if ($null -ne $preProvisioningScript)
    {
        if (Test-Path -Path $preProvisioningScript)
        {
            Write-Log "[$siteUrl] Running pre-provisioning script '$preProvisioningScript'" -WriteToHost

            # Run the pre-provisioning script, if there is one
            . $preProvisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl -ConfigFile $ConfigFile -SiteType $SiteType
        }
        else
        {
            Write-Log "[$siteUrl] Could not find pre-provisioning script '$preProvisioningScript'" -Level Warn
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
                Write-Log "[$siteUrl] Inovking PnP template '$provisioningTemplate'" -WriteToHost

                # Invoke the provisioning template, if there is one
                Invoke-PnPSiteTemplate -Path $provisioningTemplate -Connectiion $global:siteConn
            }
            else
            {
                Write-Log "[$siteUrl] Could not find provisioning template '$provisioningTemplate'" -Level Warn
            }
            
        }
    }

    #Post-Provisioning
    $postProvisioningScripts = Get-NestedMember $config "plugins.$entityTypeConfigKey.postProvisioningTask"
    if ($null -ne $postProvisioningScripts)
    {
        foreach ($postProvisioningScript in $postProvisioningScripts)
        {
            if (Test-Path -Path $postProvisioningScript)
            {
                Write-Log "[$siteUrl] Running post-provisioning script '$postProvisioningScript'" -WriteToHost

                # Run the post-provisioning script, if there is one
                . $postProvisioningScript -TenantUrl $config.rootSiteUrl -SitePath $Site -SiteTitle $SiteTitle -FullSiteUrl $siteUrl -ConfigFile $ConfigFile -SiteType $SiteType
            }
            else
            {
                Write-Log "[$siteUrl] Could not find post-provisioning script '$postProvisioningScript'" -Level Warn
            }
        }
    }

    #Permissions defaults
    if ($EntityType -ne "MSTeam")
    {
        $permissions = Get-NestedMember $config "plugins.$entityTypeConfigKey.permissions"
        if ($null -ne $permissions)
        {
            #$connection = Connect-PnPOnline -Url $siteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop

            $owners = $permissions.owners
            if ($null -ne $owners)
            {
                $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -Connection $global:siteConn
                foreach ($o in $owners) {
                    Add-PnPGroupMember -EmailAddress $o -Group $ownerGroup.Id -Connection $global:siteConn
                    Write-Log "[$siteUrl] $o added to Site Owners" -WriteToHost
                }
            }

            $members = $permissions.members
            if ($null -ne $members)
            {
                $memberGroup = Get-PnPGroup -AssociatedMemberGroup -Connection $global:siteConn
                foreach ($m in $members) {
                    Add-PnPGroupMember -EmailAddress $m -Group $memberGroup.Id -Connection $global:siteConn
                    Write-Log "[$siteUrl] $m added to Site Members" -WriteToHost
                }
            }

            $visitors = $permissions.visitors
            if ($null -ne $visitors)
            {
                $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -Connection $global:siteConn
                foreach ($v in $visitors) {
                    Add-PnPGroupMember -EmailAddress $v -Group $visitorGroup.Id -Connection $global:siteConn
                    Write-Log "[$siteUrl] $v added to Site Visitors" -WriteToHost
                }
            }
        }
    }
}


$webpartFiles = $config.webparts.files | Where-Object { $_.deployToTenant -eq $false }
if($webpartFiles)
{
    Write-Log "[$siteUrl] Installing Webparts" -WriteToHost
    . "./install-webpart.ps1" -SiteUrl $newSiteUrl -Credentials $global:cred -ConfigFile $ConfigFile -SkipGetCredentials
}

# Make sure all open connections are closed
#Disconnect-OpenConnections
 