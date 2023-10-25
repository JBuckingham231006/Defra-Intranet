<#
    SCRIPT OVERVIEW:
    REGRESSION SCRIPT FOR: 01 - DEFRA Intranet - Update the Site Pages Library.ps1
    This script uninstalls our custom column(s) from the "Site Page" libraries in the Intranet sites, except the EA site.

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

# WARNING: We do NOT want to uninstall the column from the EA Intranet as this column is in-use already within production. We're only want to remove our introduction of this column to the other sites
$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 -and $_.Abbreviation -ne "EA" } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'DEFRA Intranet' or is not configured correctly"
}

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"
    Connect-PnPOnline -Url $fullURL -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $fullURL" -ForegroundColor Cyan

    $fieldNames = @("Content_x0020_Owner_x0020__x002d__x0020_Team")

    foreach($fieldName in $fieldNames)
    {
        $field = Get-PnPField -Identity $fieldName -ErrorAction SilentlyContinue

        if($null -ne $field)
        {
            Remove-PnPField -Identity $fieldName -Force
            Write-Host "SITE COLUMN REMOVED: $fieldName" -ForegroundColor Cyan
        }
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript