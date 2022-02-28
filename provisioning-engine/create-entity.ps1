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
    [ValidateSet("IntranetSpokeSite", "CommunicationSite", "TeamSite", "MSTeam")]    
    [string]$EntityType,
    [ValidateSet("", "Public", "Private")]    
    [string]$Visibility,
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
    Write-Output "create-entity.ps1"
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

$disconnectWhenDone = $true
# this allows us to set $cred before executing script and not be prompted
if ($SkipGetCredentials.IsPresent -eq $false)
{
    Write-Output "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."

    # Connect to SharePoint
    Write-Output "Connect-PnpOnline"
    Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred #-Scopes Group.ReadWrite.All

    # Connect to Teams
    Write-Output "Connect-MicrosoftTeams"
    Connect-MicrosoftTeams -Credential $global:cred
}
else
{
    $disconnectWhenDone = $false
}

$siteUrl = Get-UrlByEntityType $EntityType $Site $config

# Check if Site Already Exists
$existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore

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
            Write-Output "New-PnPSite -Type CommunicationSite -Title $($SiteTitle) -Url $($siteUrl)"
            $newSiteUrl = New-PnPSite -Type $EntityType `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription `
                -SiteDesign "Topic"
        }
        else 
        {
            # Site exists and not created
            Write-Warning "Site already exists so it wasn't created. $($existingSite.Url)"
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
            $newSiteUrl = New-PnPSite -Type TeamSiteWithoutMicrosoft365Group `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription


            # Write-Output "New-PnPSite -Type TeamSite -Title $($SiteTitle) -Alias $($Site) -Url $($siteUrl)"
            # $newSiteUrl = New-PnPSite -Type TeamSite `
            #     -Title $SiteTitle `
            #     -Alias $Site `
            #     -Description $SiteDescription
        }
        else 
        {
            # Site exists and not created
            Write-Warning "Site already exists so it wasn't created. $($existingSite.Url)"
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
            Write-Output "New-Team -MailNickName $($Site) -DisplayName $($SiteTitle) -Description $($SiteDescription) -Visibility $($Visibility)"
            # The MailNickName should be the URL
            $team = New-Team -MailNickName $Site `
                -DisplayName $SiteTitle `
                -Description $SiteDescription `
                -Visibility $Visibility
        }
        else 
        {
            # Site exists and not created
            Write-Warning "Team already exists so it wasn't created. $($existingSite.Url)"
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
            $newSiteUrl = New-PnPSite -Type "CommunicationSite" `
                -Title $SiteTitle `
                -Url $siteUrl `
                -Description $SiteDescription `
                -SiteDesign "Topic"
        }
        else 
        {
            # Site exists and not created
            Write-Warning "Site already exists so it wasn't created. $($existingSite.Url)"
            $newSiteUrl = $existingSite.Url
        
            Write-Output "$($existingSite.Url)" >> existingsites.log
        }
    }
}

$webpartFiles = $config.webparts.files | Where-Object { $_.deployToTenant -eq $false }
if($webpartFiles)
{
    Write-Host "Installing Webparts"
    . "./install-webpart.ps1" -SiteUrl $newSiteUrl -Credentials $global:cred -ConfigFile $ConfigFile
}

if ($disconnectWhenDone -eq $true)
{
    # Disconnect from PnPOnline & SPOService
    Write-Output "Disconnect from SharePoint"
    Disconnect-PnPOnline
}