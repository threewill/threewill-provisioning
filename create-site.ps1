<#     
    .SYNOPSIS
    Create s site or team based on the provided parameters

    .DESCRIPTION
    This script will create a site or team based on the parameters passed in.  If the SharePoint Url
    for the site exists it will be skipped.

    .PARAMETER configFile
    Relative path of the configuration file to use (e.g. config\contoso.json)

    .EXAMPLE
    .\create-site.ps1 -configFile config\contoso.json -site "contoso-templatetest" -siteTitle "Contoso Template Test" -siteDescription "Test Description"

    .NOTES
    - Dependencies: 
        SharePointPnPPowerShellOnline cmdlets, version 3.12.1908.1 or higher (August 2019 Intermediate Release 1)
        Teams cmdlets, version 1.01 or higher (Install-Module -Name MicrosoftTeams)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$configFile,
    [Parameter(Mandatory = $true)]
    [int]$placeID,    
    [Parameter(Mandatory = $true)]
    [string]$site,    
    [Parameter(Mandatory = $true)]
    [string]$siteTitle,
    [Parameter(Mandatory = $true)]
    [string]$siteDescription,    
    [Parameter(Mandatory = $true)]
    [ValidateSet("CommunicationSite", "TeamSite", "MSTeam")]    
    [string]$siteType,
    [ValidateSet("", "Public", "Private")]    
    [string]$visibility,
    [switch]$skipGetCredentials,
    [switch]$batchMode
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"

# only write out stuff if we aren't being called by another script
if ($batchMode.IsPresent -eq $false) {
    Write-Output "".PadRight(50, "=")
    Write-Output "create-site.ps1"
    Write-Output "  params: site:$($site)"
    Write-Output "  params: configFile:$($configFile),site:$($site),siteType:$($siteType)"
    Write-Output "".PadRight(50, "=")
}

# Get configuration
$config = Get-Content $configFile | Out-String | ConvertFrom-Json

# only write out stuff if we aren't being called by another script
if ($batchMode.IsPresent -eq $false) {
    Write-Output "Start Config Values ".PadRight(50, "*")
    Write-Output $config
    Write-Output "End Config Values ".PadRight(50, "*")
}

$disconnectWhenDone = $true
# this allows us to set $cred before executing script and not be prompted
if ($skipGetCredentials.IsPresent -eq $false) {
    Write-Output "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."

    # Connect to SharePoint
    Write-Output "Connect-PnpOnline"
    Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred #-Scopes Group.ReadWrite.All

    # Connect to Teams
    Write-Output "Connect-MicrosoftTeams"
    Connect-MicrosoftTeams -Credential $global:cred
}
else {
    $disconnectWhenDone = $false
}

$siteUrl = $null
if ($siteType -eq "CommunicationSite") {
    $siteUrl = "$($config.rootSiteUrl)/$($config.communicationSiteDefaultPath)/$($site)"
}
else {
    $siteUrl = "$($config.rootSiteUrl)/$($config.teamSiteDefaultPath)/$($site)"
}

# Check if Site Already Exists
$existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore

$newSiteUrl = $null
if ($null -eq $existingSite) {

    if ($siteType -eq "CommunicationSite") {
        #     
        # https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/new-pnpsite?view=sharepoint-ps    
        #
        Write-Output "New-PnPSite -Type CommunicationSite -Title $($siteTitle) -Url $($siteUrl)"
        $newSiteUrl = New-PnPSite -Type $siteType `
            -Title $siteTitle `
            -Url $siteUrl `
            -Description $siteDescription `
            -SiteDesign "Topic"
    }    
    elseif ($siteType -eq "TeamSite") {
        Write-Output "New-PnPSite -Type TeamSite -Title $($siteTitle) -Url $($siteUrl)"
        $newSiteUrl = New-PnPSite -Type TeamSite `
            -Title $siteTitle `
            -Alias $site `
            -Description $siteDescription
    }
    elseif ($siteType -eq "MSTeam") {

        Write-Output "New-Team -MailNickName $($site) -DisplayName $($siteTitle) -Description $($siteDescription) -Visibility $($visibility)"
        # The MailNickName should be the URL
        New-Team -MailNickName $site `
            -DisplayName $siteTitle `
            -Description $siteDescription `
            -Visibility $visibility | Out-Null
    }
}
else {
    # Site exists and not created
    Write-Warning "Site or Team already exists so it wasn't created. $($existingSite.Url)"
    $newSiteUrl = $existingSite.Url

    Write-Output "$($existingSite.Url)" >> existingsites.log
}

$webpartFiles = $config.webparts.files | Where-Object { $_.deployToTenant -eq $false }
if($webpartFiles){
    Write-Host "Installing Webparts"
    . "./install-webpart.ps1" -SiteUrl $newSiteUrl -Credentials $global:cred -ConfigFile $configFile
}

if ($disconnectWhenDone -eq $true) {
    # Disconnect from PnPOnline & SPOService
    Write-Output "Disconnect from SharePoint"
    Disconnect-PnPOnline
}