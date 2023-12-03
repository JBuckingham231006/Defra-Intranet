<#
    SCRIPT OVERVIEW:
    This script creates our custom lists required within the Defra Intranet

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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "Entries could not be found in the configuration module that matches the requirements for this script to run. The Defra Intranet and all associated ALB intranets are required."
}

# LIST - Create the "News Article Approval Information" list in which of the Intranet sites
$displayName = "News Article Approval Information"
$listURL = "Lists/SPAI"

$fieldNames = @("AssociatedSitePage","NewsArticleTitle","ContentSubmissionStatus","DateOfApprovalRequest")

Write-Host "`nCREATING THE '$displayName' LIST" -ForegroundColor Green

$site = $sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

$ctx = Get-PnPContext

$list = Get-PnPList -Identity $listURL -ErrorAction SilentlyContinue

if($null -eq $list)
{
    $list = New-PnPList -Template GenericList -Title $displayName -Url $listURL -Hidden
    Write-Host "LIST CREATED: $displayName (URL: $listURL)" -ForegroundColor Green
}
else
{
    Write-Host "THE LIST '$displayName' ALREADY EXISTS" -ForegroundColor Yellow
}

# FIELDS - ADD OUR CUSTOM FIELDS TO THE LIST 
Write-Host "`nADDING OUR FIELDS TO THE LIST" -ForegroundColor Green

foreach($fieldName in $fieldNames)
{
    $field = Get-PnPField -List $list -Identity $fieldName -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        Add-PnPField -List $list -Field $fieldName
        Write-Host "FIELD ADDED TO THE '$displayName' LIST: $fieldName" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$fieldName' ALREADY EXISTS IN THE LIST '$displayName'" -ForegroundColor Yellow 
    }
}

Set-PnPField -List $list -Identity "Title" -Values @{
    Title = "Approving ALB"
    Required = $false
}

Write-Host "CUSTOMISED THE 'Title' FIELD" -ForegroundColor Green

# UPDATE VIEW INFORMATION
$view = Get-PnPView -List $list -Identity "All Items"

if($null -ne $view)
{
    $view = Set-PnPView -List $list -Identity $view.Title -Fields @("AssociatedSitePage","NewsArticleTitle","Title","ContentSubmissionStatus","DateOfApprovalRequest")

    $view.ViewQuery = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="DateOfApprovalRequest" Ascending="FALSE" /></GroupBy><OrderBy><FieldRef Name="DateOfApprovalRequest" Ascending="FALSE" /></OrderBy>'
    $view.Update()
    $ctx.ExecuteQuery()

    Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
}
else
{
    Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' DOES NOT EXIST" -ForegroundColor Yellow
}

# LIST SETTING AND PERMISSION UPDATES
Set-PnPList -Identity $list -EnableAttachments 0
Write-Host "LIST ATTACHMENTS DISABLED" -ForegroundColor Green

# Break Permission Inheritance of the List and set the new permissions for the members
Set-PnPList -Identity $list -BreakRoleInheritance

$membersGroup = Get-PnPGroup | Where-Object { $_.Title -like "* Members"}
$ownersGroup = Get-PnPGroup | Where-Object { $_.Title -like "* Owners"}

Set-PnPListPermission -Identity $list -AddRole "Read" -Group $membersGroup
Set-PnPListPermission -Identity $list -AddRole "Read" -Group $ownersGroup

Write-Host "LIST PERMISSIONS UPDATED" -ForegroundColor Green

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript