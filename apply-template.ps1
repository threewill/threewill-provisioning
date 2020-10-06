<#     
    .SYNOPSIS
    Apply PnP Provisioning Template to an existing site

    .DESCRIPTION
    This script will use the PnP PowerShell Commandlets and SharePoint Online Commandlets to apply templates to an existing site collection.  It will apply a 'common' base template and then the specific one based on the provided site type.  Additionally, it will turn off the NoScript option on a modern site that is set by default so that we can apply web properties.

    .PARAMETER configFile
    Relative path of the configuration file to use (e.g. config\contoso.json)
    .PARAMETER siteUrl
    Full url for the site to apply the template to
    .PARAMETER  templateType
    Identifier for the type of template to use from the set of valid types available.
    Valid site types are "SoelutionsGroup", "Brand", "SharedService", "Department", "Office", "Community"

    .EXAMPLE
    .\apply-template.ps1 -configFile config\contoso.json -siteUrl https://contoso.sharepoint.com/sites/contoso-templatetest

    .NOTES 
    Dependencies: 
        Microsoft.Online.SharePoint.PowerShell Module, version 16.0.7618.0 or higher
        SharePointPnPPowerShellOnline Module, version 3.12.1908.1 or higher (August 2019 Intermediate Release 1)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$configFile,    
    [Parameter(Mandatory = $true)]
    [string]$siteUrl,
    [switch]$skipGetCredentials,
    [switch]$batchMode
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"

# only write out stuff if we aren't being called by another script
if ($batchMode.IsPresent -eq $false) {
    Write-Output "".PadRight(50, "=")
    Write-Output "apply-template.ps1"
    Write-Output "  params: siteUrl:$($siteUrl)"
    Write-Output "  params: configFile:$($configFile)"
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
# this allows us to set $global:cred before executing script and not be prompted
if ($skipGetCredentials.IsPresent -eq $false) {
    Write-Output "Prompt for SharePoint Credentials"
    $global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
}
else {
    $disconnectWhenDone = $false
}

if ($null -ne $config.adminSiteUrl -and $null -ne $config.rootSiteUrl) {
    # Connect to SharePoint Root Site 
    Connect-PnPOnline $config.rootSiteUrl -Credentials $global:cred
    
    # Connect to Tenant Admin Site
    Connect-PnPOnline -Url $config.adminSiteUrl -Credentials $global:cred

    # Check if Site Exists
    $exists = Get-PnPTenantSite -Url $siteUrl -ErrorAction Ignore

    if ($null -ne $exists) {
        # Connect to the site
        Connect-PnPOnline -Url $siteUrl -Credentials $global:cred

        Set-PnPTenantSite -Url $siteUrl -DenyAddAndCustomizePages:$false
        
        $templatePath = "..\templates\contoso-site-template.xml"
        $templateExists = Test-Path $templatePath
        if ($templateExists -eq $true) {
            try {
                Write-Output "Applying '$($templatePath)' template to '$($siteUrl)'"
                Apply-PnPProvisioningTemplate $templatePath
            }
            catch {
                "ERROR [$($_.Exception.Message)] [$(get-date)]: - $($siteurl)" | Out-File -FilePath "ErrorSites.txt" -Append
                Write-Error "ERROR: $($_.Exception.Message)" -ErrorAction Continue
            }            
        }
        else {
            Write-Error "Template not applied. Could not find template at path '$($templatePath)'" -ErrorAction Continue
        }

        # pnp:theme node doesn't seem to work in site template so doing it with pnp command
        #Get-PnPTenantTheme -Name "Contoso Orange" | Set-PnPWebTheme

        # Set the switch back to NoScript
        Set-PnPTenantSite -Url $siteUrl -DenyAddAndCustomizePages:$true
    }
    else {
        Write-Error "Site at url '$($siteUrl) does not exist" -ErrorAction Continue
    }

    if ($disconnectWhenDone -eq $true) {
        # Disconnect from PnPOnline
        Write-Output "Disconnect from SharePoint"
        Disconnect-PnPOnline
    }
}