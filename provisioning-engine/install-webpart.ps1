<#     
    .SYNOPSIS
    Installs Webparts to a tenant or site app catalog.

    .DESCRIPTION
    This script can be used to deploy apps to either a Tenant or Site app catalog. It requires a valid configuration file, which has been documented in the readme.
    Deploying to the tenant app requires the SharePoint admin role and will make the webpart available for use in all sites in the tenant. 

    Deploying to the site level requires the Site Admin permission and will make the webpart available only for the specified sites. 

    .PARAMETER ConfigFile
    REQUIRED. Relative path the JSON configuration file to be used.

    .PARAMETER SiteURL
    OPTIONAL. The URLs of sites to deploy site scoped apps to. URLs can be relative (/sites/your-site) or full (https://contoso.sharepoint.com/sites/your-site)
    
    If one or more URLS are provided, only webparts not scoped to the tenant will be deployed. This parameter accepts input from the pipeline.
    
    .PARAMETER Credentials
    OPTIONAL. Can be used to pass a PSCredential that represents the M365 login credentials of the user running the script. If provided, the script will not prompt the user to login. 

    .PARAMETER UserName
    OPTIONAL. If passed, will only prompt the user for a password (unless the Credentials parameter was used).

    .EXAMPLE
    Deploy all tenant scoped apps to the tenant using the specified credentials object.
        ./create-site-batch.ps1 -ConfigFile './config/willdev.json' -Credentials $myCredentials
    
    Deploy all site scoped apps to the specified site. Will prompt user for credentials.
        ./create-site-batch.ps1 -ConfigFile './config/willdev.json' -SiteUrl 'https://contoso.sharepoint.com/sites/test-site'

    Deploy all site scoped apps to multiple sites using the specified credentials.
        ./create-site-batch.ps1 -ConfigFile './config/willdev.json' -SiteUrl '/sites/test-site-1', '/sites/test-site-2' -Credentials $myCredentials

    Deploy all site scoped apps to multiple sites piped in using the specified credentials.
        '/sites/test-site-1', '/sites/test-site-2' | ./create-site-batch.ps1 -ConfigFile './config/willdev.json' -Credentials $myCredentials

    .NOTES
    - Dependencies: 
        SharePointPnPPowerShellOnline cmdlets, version 3.12.1908.1 or higher (August 2019 Intermediate Release 1)
        Teams cmdlets, version 1.01 or higher (Install-Module -Name MicrosoftTeams)
#>
[CmdletBinding(DefaultParameterSetName="TENANT")]
PARAM(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,

    [Parameter(Position=0, Mandatory=$true, ParameterSetName="SITE", ValueFromPipeline=$true)]
    [ValidatePattern("^(?:https:\/\/\S+\.sharepoint\.com\/|\/?)(?:sites|teams)\/\S+$")]
    [string[]]$SiteURL,
  
    [Parameter(Mandatory=$false)]
    [PSCredential]
    $Credentials,

    [Parameter(Mandatory=$false)]
    [string]
    $UserName,
    [switch]$SkipGetCredentials
)
BEGIN
{
    $isTenantScoped = $PSCmdlet.ParameterSetName -eq "TENANT"

    # Load and validate config file
    $config = Get-Content $configFile | Out-String | ConvertFrom-Json
    $webpartFiles = $config.webparts.files | Where-Object { $_.deployToTenant -eq $isTenantScoped }
    
    if($null -eq $webpartFiles)
    {
        Write-Log "No webparts found to deploy at $($PSCmdlet.ParameterSetName) scope. Exiting." -Level Warn
        exit
    }

    # Get Credentials
    if($null -eq $Credentials )
    {
        if([String]::IsNullOrEmpty($UserName))
        {
            $credentials = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
			$global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Interactive -ReturnConnection -ErrorAction Stop
        }
        else
        {
            $credentials = Get-Credential -UserName $UserName -Message "Please provide the password for $UserName"
			$global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Interactive -ReturnConnection -ErrorAction Stop
        }
    }

    function Update-SiteUrl{
        PARAM(
            [Parameter(Mandatory=$true)]
            [string]$url
        )
        if($url -match "^\/?(?:sites|teams)\/\w+$")
        {
            if($url.startsWith('/'))
            {
                $url = $config.rootSiteUrl + $url;
            }
            else
            {
                $url = $config.rootSiteUrl + '/' + $url;
            }
        }
        return $url
    }
}
PROCESS
{
    if($isTenantScoped)
    { 
        # We're just installing apps to the Tenant
        $url = $config.adminSiteUrl


        try
        {
            # Connect to the Tenant Admin site
            Write-Log "[$url] Connecting" -WriteToHost
            #Connect-PnPOnline -Url $url -Credentials $credentials -ErrorAction Stop
        }
        catch
        {
            Write-Log $_ -Level Error
            exit
        }

        foreach($file in $webpartFiles)
        {
            Write-Log "[$url] Uploading $($file.fileName)" -WriteToHost
            $app = Add-PnPApp -path "$($config.webparts.pathToFolder)/$($file.fileName)" -Connection $global:tenantConn -Scope Tenant -Overwrite -Publish -SkipFeatureDeployment
        }
    }
    else
    {
        # We're only installing site scoped webparts            
        foreach($url in $SiteUrl)
        {
            $url = Update-SiteUrl $url

            if ($SkipGetCredentials.IsPresent -eq $false)
            {
                try
                {
                    # Connect to the current Site
                    Write-Log "[$url] Connecting" -WriteToHost
                    #Connect-PnPOnline -Url $url -Credentials $credentials -ErrorAction Stop
                }
                catch
                {
                    Write-Log $_ -Level Error
                    exit
                }   
            }

            Write-Log "[$url] Preparing app catalog" -WriteToHost
            Add-PnPSiteCollectionAppCatalog -Site $url -ErrorAction SilentlyContinue -Connection $global:tenantConn

            # Upload app to app catalog
            foreach($file in $webpartFiles)
            {
                Write-Log "[$url] Uploading $($file.fileName)" -WriteToHost
                $app = Add-PnPApp -path "$($config.webparts.pathToFolder)/$($file.fileName)" -Scope Site -Overwrite -Publish -Connection $global:tenantConn

                # Install the app to the site.
                if($null -ne $app)
                {
                    $installedApp = Get-PnPApp -Identity $app -ErrorAction SilentlyContinue -Connection $global:tenantConn
                    if($installedApp)
                    {
                        Write-Log "[$url] $($file.fileName) already installed." -WriteToHost
                    }
                    else
                    {
                        Write-Log "[$url] Installing $($file.fileName)" -WriteToHost
                        Install-PnPApp -Identity $app -Scope Site -Connection $global:tenantConn
                    }
                }
            }
            #Disconnect-PnPOnline
        }
    }
}
END
{
    
}
