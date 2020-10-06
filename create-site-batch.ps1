<#     
    .SYNOPSIS
    Create sites and teams based on a csv file

    .DESCRIPTION
    This script will create sites and teams based on the provided csv file.  It depends on the create-site.ps1
    script that is invoked when this loops through the rows in the csv and constructs the call

    .PARAMETER configFile
    Relative path of the configuration file to use (e.g. config\coxautoinc.json)
    .PARAMETER sitesFile
    Relative path of a csv file containing a collection of site information

    .EXAMPLE
    .\create-site-batch.ps1 -configFile config\contoso.json -sitesFile config\test-template-sites.csv

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
    [string]$sitesFile    
)


$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"
$scriptTime = [System.Diagnostics.Stopwatch]::StartNew()

Write-Output "".PadRight(50, "=")
Write-Output "apply-template-batch.ps1"
Write-Output "  params: configFile:$($configFile),sitesFile:$($sitesFile)"
Write-Output "".PadRight(50, "=")

# Get configuration
$config = Get-Content $configFile | Out-String | ConvertFrom-Json

# Write out configuration information
Write-Output "Start Config Values ".PadRight(50, "*")
Write-Output $config
Write-Output "End Config Values ".PadRight(50, "*")

$global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
$global:urls = @()

# Connect to SharePoint
Write-Output "Connect-PnpOnline"
Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred #-Scopes Group.ReadWrite.All

# Connect to Teams
Write-Output "Connect-MicrosoftTeams"
Connect-MicrosoftTeams -Credential $global:cred

if ($null -ne $global:cred) {

    $sitesData = Import-CSV (Join-Path $script:PSScriptRoot $sitesFile)

    foreach ($line in $sitesData) {
        if ($line.PlaceID.StartsWith("#")) {
            # This CSV line is "commented out" - so skip it
            continue
        }

        $description = $line.Description
        if($line.Description -eq "") {
            $description = $line.Title
        }                

        Write-Output "Start Invoke-Expression".PadRight(50, "*")
        Write-Output ".\create-site.ps1"
        Write-Output "-configFile '$($configFile)'"
        Write-Output "-placeID $($line.PlaceID)"
        Write-Output "-site '$($line.Site)'"
        Write-Output "-siteTitle '$($line.Title)'"
        Write-Output "-siteDescription '$($description)'"
        Write-Output "-siteType '$($line.SiteType)'"
        Write-Output "-visibility '$($line.Visibility)'"
        Write-Output "-skipGetCredentials -batchMode"
        Write-Output ""
        
        Invoke-Expression ".\create-site.ps1 -configFile '$($configFile)' -placeID '$($line.PlaceID)' -site '$($line.Site)' -siteTitle '$($line.Title)' -siteDescription '$($description)' -siteType '$($line.SiteType)' -visibility '$($line.Visibility)' -skipGetCredentials -batchMode"
        
        Write-Output "End Invoke-Expression".PadRight(50, "*")
        Write-Output ""

        # just doing this so we have an easy list of urls after we are done
        if ($line.SiteType -eq "CommunicationSite") {
            $global:urls += "$($config.rootSiteUrl)/$($config.communicationSiteDefaultPath)/$($line.Site)"
        }
        else {
            $global:urls += "$($config.rootSiteUrl)/$($config.teamSiteDefaultPath)/$($line.Site)"
        }
    }
}

# doing this last because I don't want to have to maintain multiple PnP Contexts (admin for create sites and this for each new site)
if($config.additionalSiteCollectionAdmins -gt "") {
    Write-Output "Adding additional site collection admins"

    $admins = @()
    foreach ($admin in $config.additionalSiteCollectionAdmins.Split(",")) {
        $admins += $admin.Trim()
    }
    foreach ($siteUrl in $global:urls) {
        Connect-PnPOnline -Url $siteUrl -Credentials $global:cred
        Write-Output "Add-PnPSiteCollectionAdmin -Owners $($admins) for site $($siteUrl)"
        try
        {
            Add-PnPSiteCollectionAdmin -Owners $admins
        }
        catch
        {
            Write-Warning "ERROR: $($_.Exception.Message)"
        }
    }    
}

$global:urls | Out-File "create-site-batch.txt"

Write-Output "".PadRight(50, "=")
Write-Output "Execution Time: [$(get-date)] [$($scriptTime.Elapsed.ToString())]"