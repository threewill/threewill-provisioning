# Batch site provisioning

## Pre-requisites

### PowerShell
This script is intended to run on [PowerShell Core](https://github.com/PowerShell/PowerShell) (tested on version 7.2.1).  To verify that PowerShell Core is installed, run command ``$PSVersionTable`` and verify that the PowerShell version is >= 6.

### PnP PowerShell Module
This script requires [PnP PowerShell](https://pnp.github.io/powershell/) (tested on version 1.9.0).  This can be installed with the following command:  
``Install-Module -Name "PnP.PowerShell"``

In addition,You will have to consent / register the PnP Management Shell Multi-Tenant Azure AD Application in your destination tenant.
Run this command to grant consent:  
``Register-PnPManagementShellAccess``

### Microsoft Teams PowerShell Module
This script requires [Microsoft Teams PowerShell Module](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-overview) (tested on version 3.1.1).  This can be installed with the following command:  
``Install-Module -Name MicrosoftTeams -Force -AllowClobber``

### Execution Policy
If these scripts were downloaded from the internet, they may be blocked from running by Window's execution policy.  To get around this, you can either change the Execution Policy or unblock the file.

To change the [Execution Policy](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.2), run:  
``Set-ExecutionPolicy -ExecutionPolicy Unrestricted``

Alternatively, unblock the file  
``Unblock-File <path to file>``


## Configuration

### Config script
This script relies on an external json file for configuration.  To reduce errors in configuration, this config file must conform to a json schema definition (./config/provisioning-config-schema.json).

#### Plug-ins
This configuration script allows defining custom PowerShell scripts that function as plug-ins that allow for adding custom tasks.  

These PowerShell scripts take a uniform function signature:  
```
param(  
    [Parameter(Mandatory = $true)]  
    [string]$TenantUrl,  
    [Parameter(Mandatory = $true)]  
    [string]$SitePath,  
    [Parameter(Mandatory = $true)]  
    [string]$SiteTitle,  
    [Parameter(Mandatory = $true)]  
    [string]$ConfigFile,
    [string]$SiteType  
)
```

The exception is the MSTeam provisioning script, which takes an additional Visibility parameter:  
```
param(  
    [Parameter(Mandatory = $true)]  
    [string]$TenantUrl,  
    [Parameter(Mandatory = $true)]  
    [string]$SitePath,  
    [Parameter(Mandatory = $true)]  
    [string]$SiteTitle,  
    [Parameter(Mandatory = $true)]  
    [string]$ConfigFile  
    [ValidateSet("", "Public", "Private")]  
    [string]$Visibility  
)
```

For credentialing, these scripts have access to the parents ``$global:creds`` variable.  However, for all files except for the provisioning script, we can expect that a connection to the site collection has already been open.  For the provisioning script, the user will be responsible for creating the connection and disconnecting.

##### Pre- and Post- provisioning tasks
For each entity type, you may define a pre-provisioning and post-provisioning tasks.  For these tasks, we can expect that a connection to the site connection has already been established.

##### Template
For each entity type (excluding MS Teams), you may define a template file, which will be invoked after the site has been created.  This is mutually exclusive to the provisioning script, enforced through the json schema.

##### Provisioning
For each entity type, you may define a provisioning script.  It is assumed that this script will handle both site creating and site provisioning.  To reference a template file in the same folder will have to be refrenced using ``"$PSScriptRoot/<template-name>.xml"``. This is mutually exclusive to providing a provisioning template and pre- and post-provisioning tasks, enforced through the json schema.  The user will be responsible for connecting and disconnecting to SharePoint as appropriate.


### Batch CSV
The list of sites to provision is provided in CSV format.  The CSV is expected to have the following format.

| PlaceID | Site | Title | Description | EnttiyType | SiteType | IsHub | AssociateToHub | Visibility |

## Execution
The parent script takes two mandatory parameters and has two optional parameters:  
``ConfigFile`` - The relative path to the config file  
``SitesFile`` - The relative path to the batch sites CSV
``UseHistory`` - (optional) A switch to tell the script to use previous run status data
``StatusFile`` - (optional) A relative path to the status file the user wishes to reference

### Run execution history
The provisioning engine keeps historical records of previous runs and outcomes.  These records are stored as .csv in a folder called "status", with a subfolder of the datetime of the provisioning run.  The user can utilize these historical records by specifying the ``-UseHistory`` parameter when calling the batch script.  By default, this flag will cause the script to pick the most recent status output file as a record and attempt to not replicate any work that has already been done.  Alternatively, the user can supply a ``-StatusFile`` parameter to a specific file if the default is not applicable.  If no status folder and file is found, the the script will run as if the ``-UseHistory`` flag was not present.  If a file is specified but not found, the user will be prompted if they wish to continue without status.

## Example call

```powershell
cd provisioning-engine
.\run-provisioning.ps1 -ConfigFile ".\config\devConfig.json" -SitesFile ".\config\default-batch.csv" [-UseHistory]
```