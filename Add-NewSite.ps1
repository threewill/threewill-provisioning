[CmdletBinding(DefaultParameterSetName="PROVISION")]
PARAM(
    [Parameter(Mandatory=$true)]
    [string]$Title,
    [Parameter(Mandatory=$true)]
    [string]$Url,
    [Parameter(Mandatory=$true)]
    [string]$OwnerEmail,
    [Parameter(Mandatory=$true)]
    [string]$TemplatePath,
    [Parameter(Mandatory=$false)]
    [int]$LCID = 1033,
    [Parameter(Mandatory=$false)]
    [string[]]$AdditionalAdmins,
    [switch]$TeamSite    
)
BEGIN{
    Import-Module "$PSScriptRoot\libs\SharePointPnPPowerShellOnline\SharePointPnPPowerShellOnline.psd1" -DisableNameChecking -Verbose:$false -Force
    $siteType = if($TeamSite){ return TeamSite } else{ return CommunicationSite }
}
PROCESS{    
    try{
        # Connect to Tenant Administration, Migration site, and Teams - May Require Logon
        Connect-PnPOnline -Url $tenantAdminUrl -ErrorAction Stop -UseWebLogin -ReturnConnection

        # Create the Site - (see: https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/new-pnpsite?view=sharepoint-ps)
        New-PnPSite -Title $Title -Url $Url -Owner $OwnerEmail -Lcid $LCID -Type $siteType -ErrorAction Stop

        # Wait for provisioning to complete and then connect to site.
        Connect-PnPOnline -Url $URL -ErrorAction Stop -ReturnConnection -UseWebLogin

        # Add Additonal Site Collection Admins
        if($null -ne $AdditionalAdmins -and $AdditionalAdmins.length -gt 0){
            Add-PnPSiteCollectionAdmin -Owners $AdditionalAdmins -ErrorAction Stop
        }

        Apply-PnPProvisioningTemplate -Path $TemplatePath
    }
    catch{
        Write-Error $_
    }            
}
END{
    Disconnect-PnPOnline
}
