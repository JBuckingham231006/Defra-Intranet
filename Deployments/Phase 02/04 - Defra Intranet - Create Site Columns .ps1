<#
    SCRIPT OVERVIEW:
    This script creates the site columns used by our custom lists and libraries within the Defra and ALB SharePoint sites, and site columns for the existing lists and libraries.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the Defra and ALB Intranet SharePoint sites
    OR
    - Access to the SharePoint Tenant Administration site
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

$site = $global:sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
Write-Host ""

# SITE PAGE FIELDS 
# "Organisation (Intranets)" column
$displayName = "Organisation (Intranets)"
$field = Get-PnPField | Where-Object { $_.InternalName -eq "OrganisationIntranetsContentEditorInput" }
$termSetPath = $global:termSetPath

if($null -eq $field)
{
    $field = Add-PnPTaxonomyField -DisplayName $displayName -InternalName "OrganisationIntranetsContentEditorInput" -TermSetPath $termSetPath -MultiValue
    Write-Host "SITE COLUMN INSTALLED: $($displayName)" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $($displayName)" -ForegroundColor Yellow
}

# "Organisation (Intranets)" column
$displayName = "Approval Information"
$field = Get-PnPField | Where-Object { $_.InternalName -eq "PageApprovalInfo" }

if($null -eq $field)
{
    $field = Add-PnPField -DisplayName $displayName -InternalName "PageApprovalInfo" -Type URL
    Write-Host "SITE COLUMN INSTALLED: $($displayName)" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $($displayName)" -ForegroundColor Yellow
}

Write-Host ""

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript