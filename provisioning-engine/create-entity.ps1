<#     
    .SYNOPSIS
    Create a site or team based on the provided parameters

    .DESCRIPTION
    This script will create a site or team based on the parameters passed in.  If the SharePoint Url
    for the site exists it will be skipped.

    .PARAMETER ConfigFile
    Relative path of the configuration file to use (e.g. config\boinga.json)

    .EXAMPLE
    .\create-entity.ps1 -ConfigFile config\boinga.json -Site "pwc-templatetest" -SiteTitle "PwC Template Test" -SiteDescription "Test Description"

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
    [string]$SiteDescription,    
    [Parameter(Mandatory = $true)]
    [ValidateSet("IntranetSpokeSite", "CommunicationSite", "TeamSite", "TeamSiteWithoutM365Group", "MSTeam")]    
    [string]$EntityType,
    [ValidateSet("", "Public", "Private")]    
    [string]$Visibility,
    [switch]$SkipGetCredentials,
    [switch]$BatchMode,
    [string]$AuthMode,
    [string]$ClientId
)

# Make the helper functions available to script
. "helpers/helper-functions.ps1"

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"

# only write out stuff if we aren't being called by another script
if ($BatchMode.IsPresent -eq $false)
{
    Write-Log "".PadRight(50, "=")
    Write-Log "create-entity.ps1"
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

$disconnectWhenDone = $true
# this allows us to set $cred before executing script and not be prompted
if ($SkipGetCredentials.IsPresent -eq $false)
{
    Write-Log "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."

    # Connect to SharePoint
    Write-Log "Connect-PnpOnline"
    #Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred #-Scopes Group.ReadWrite.All
	If($AuthMode -eq "Interactive")
    {
        $global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
    }
    else{
        $global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop
    }

    # Connect to Teams
    Write-Log "Connect-MicrosoftTeams"
    Connect-MicrosoftTeams -Credential $global:cred
}
else
{
    $disconnectWhenDone = $false
}

$siteUrl = Get-UrlByEntityType $EntityType $Site $config

# Check if Site Already Exists
$existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore -Connection $global:tenantConn

$newSiteUrl = $null

if ($EntityType -eq "CommunicationSite")
{
    $provisioningScript = Get-NestedMember $config "plugins.communicationSiteProvisioning.provisioningScript"
    if ($null -eq $provisioningScript)
    {
        if ($null -eq $existingSite)
        {
            # if there is no provisioning script, create the new site
            #     
            # https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/new-pnpsite?view=sharepoint-ps    
            #
            Write-Log "New-PnPSite -Type CommunicationSite -Title $($SiteTitle) -Url $($siteUrl)"
            $newSiteUrl = New-PnPSite -Connection $global:tenantConn -Type CommunicationSite `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription `
                -SiteDesign "Topic"
        }
        else 
        {
            # Site exists and not created
            Write-Log "Site already exists so it wasn't created. $($existingSite.Url)" -Level Warn
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}    
elseif ($EntityType -eq "TeamSite")
{
    $provisioningScript = Get-NestedMember $config "plugins.teamSiteProvisioning.provisioningScript"
    if ($null -eq $provisioningScript)
    {
        if ($null -eq $existingSite)
        {
            # if there is no provisioning script, create the new site
            Write-Log "New-PnPSite -Type TeamSite -Title $($SiteTitle) -Alias $($Site) -Description $($SiteDescription)"
            $newSiteUrl = New-PnPSite -Connection $global:tenantConn -Type TeamSite `
                -Title $SiteTitle `
                -Alias $Site `
                -Description $SiteDescription
        }
        else 
        {
            # Site exists and not created
            Write-Log "Site already exists so it wasn't created. $($existingSite.Url)" -Level Warn
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}
elseif ($EntityType -eq "TeamSiteWithoutM365Group")
{
    $provisioningScript = Get-NestedMember $config "plugins.teamSiteWithoutM365GroupProvisioning.provisioningScript"
    if ($null -eq $provisioningScript)
    {
        if ($null -eq $existingSite)
        {
            # if there is no provisioning script, create the new site
            Write-Log "New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $($SiteTitle) -Url $($siteUrl) -Description $($SiteDescription)"
            $newSiteUrl = New-PnPSite -Connection $global:tenantConn -Type TeamSiteWithoutMicrosoft365Group `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription
        }
        else 
        {
            # Site exists and not created
            Write-Log "Site already exists so it wasn't created. $($existingSite.Url)" -Level Warn
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}
elseif ($EntityType -eq "MSTeam")
{
    $provisioningScript = Get-NestedMember $config "plugins.msTeamsProvisioning.provisioningScript"
    if ($null -eq $provisioningScript)
    {
        if ($null -eq $existingSite)
        {
            Write-Log "New-Team -MailNickName $($Site) -DisplayName $($SiteTitle) -Description $($SiteDescription) -Visibility $($Visibility)"
            # The MailNickName should be the URL
            $newSiteUrl = New-Team -MailNickName $Site `
                -DisplayName $SiteTitle `
                -Description $SiteDescription `
                -Visibility $Visibility
        }
        else 
        {
            # Site exists and not created
            Write-Log "Team already exists so it wasn't created. $($existingSite.Url)" -Level Warn
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}
elseif ($EntityType -eq "IntranetSpokeSite")
{
    $provisioningScript = Get-NestedMember $config "plugins.intranetSpokeSiteProvisioning.provisioningScript"
    if ($null -eq $provisioningScript)
    {
        if ($null -eq $existingSite)
        {
            # if there is no provisioning script, create the new site
            Write-Log "New-PnPSite -Type CommunicationSite -Title $SiteTitle -Url $siteUrl -Description $SiteDescription -SiteDesign Topic"
            $newSiteUrl = New-PnPSite -Connection $global:tenantConn -Type CommunicationSite `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription `
                -SiteDesign "Topic"
        }
        else 
        {
            # Site exists and not created
            Write-Log "Site already exists so it wasn't created. $($existingSite.Url)" -Level Warn
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}

if ($disconnectWhenDone -eq $true)
{
    # Disconnect from PnPOnline & SPOService
    Write-Log "Disconnect from SharePoint" -WriteToHost
    #Disconnect-PnPOnline
}