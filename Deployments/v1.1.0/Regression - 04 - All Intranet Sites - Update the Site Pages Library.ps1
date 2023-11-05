<#
    SCRIPT OVERVIEW:
    REGRESSION SCRIPT FOR: 01 - DEFRA Intranet - Update the Site Pages Library.ps1
    This script uninstalls our custom column(s) from the "Site Page" libraries in the Defra and ALB Intranet sites.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Full Control (minimum requirements) 
#>

$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

Import-Module SharePointPnPPowerShellOnline

if ($null -ne $psISE)
{
    $logfileName = $($psISE.CurrentFile.FullPath.Split('\'))[$psISE.CurrentFile.FullPath.Split('\').Count-1]
    $logfileName = $logfileName.Replace(".ps1",".txt")

    $global:scriptPath = Split-Path -Path $psISE.CurrentFile.FullPath

    Import-Module "$global:scriptPath\PSModules\Configuration.psm1" -Force
    Import-Module "$global:scriptPath\PSModules\Helper.psm1" -Force
}
else
{
    Clear-Host

    $logFileName = $MyInvocation.MyCommand.Name
    $global:scriptPath = "."

    Import-Module "./PSModules/Configuration.psm1" -Force
    Import-Module "./PSModules/Helper.psm1" -Force
}

$logfileName = $logfileName.Replace(".ps1",".txt")
Start-Transcript -path "$global:scriptPath/Logs/$logfileName" -append | Out-Null

Invoke-Configuration

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "A configuration entry could not be found for '$($global:environment)', '$($global:environment)' is not configured correctly or the rules querying the configuration are returning no result"
}

foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
    Write-Host ""

    $fieldNames = @("Content_x0020_Owner_x0020__x002d__x0020_Team")

    foreach($fieldName in $fieldNames)
    {
        # We do not want to remove the field from the EA site as that's already established and being used. WARNING: Removal of this field would cause data lose.
        if($site.Abbreviation -ne "EA" -and $fieldName -ne "Content_x0020_Owner_x0020__x002d__x0020_Team")
        {
            $field = Get-PnPField -Identity $fieldName -ErrorAction SilentlyContinue

            if($null -ne $field)
            {
                Remove-PnPField -Identity $fieldName -Force
                Write-Host "SITE COLUMN REMOVED: $fieldName" -ForegroundColor Cyan
            }
        }
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript