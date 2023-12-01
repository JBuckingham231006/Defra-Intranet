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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "Entries could not be found in the configuration module that matches the requirements for this script to run. The Defra Intranet and all associated ALB intranets are required."
}

foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
    Write-Host ""

    # Custom Field values (per-site)
    switch ($site.Abbreviation)
    {
        "EA" {
            $contentTypeOptions = "Alert","Blog or Online Diary", "Guidance Page","News Story - Highlight","News Story - Top Story"
        }

        default {
            $contentTypeOptions = "Blog or Online Diary","Form","Guidance Page","News Story","Office Notice"
        }
    }

    # "Alternative Contact" column
    $displayName = "Alternative Contact"
    $field = Get-PnPField -Identity "AltContact" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "User" -InternalName "AltContact" -DisplayName $displayName
        Set-PnPField -Identity $field.Id -Values @{
            SelectionMode=0;
            Description = "Please provide the name of someone else we can contact about this request should you be out of the office."
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else 
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Content Type" column
    $displayName = "Content Types"
    $field = Get-PnPField -Identity "ContentTypes" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Choice" -InternalName "ContentTypes" -DisplayName $displayName -Required -Choices $contentTypeOptions

        Set-PnPField -Identity $field.Id -Values @{
            Description = "Please select what kind of content you are submitting:";
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Line Manager" column
    $displayName = "Line Manager"
    $field = Get-PnPField -Identity "LineManager" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "User" -InternalName "LineManager" -DisplayName $displayName -Required
        Set-PnPField -Identity $field.Id -Values @{
            SelectionMode = 0;
            Description = "Please let us know which senior management provided the final sign-off on this content."
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else 
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "When do you need this published?" column
    $displayName = "When do you need this published?"
    $field = Get-PnPField -Identity "PublishBy" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "DateTime" -InternalName "PublishBy" -DisplayName $displayName -Required
        Set-PnPField -Identity $field.Id -Values @{
            FriendlyDisplayFormat = [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType]::Disabled;
            Description="Please let us know when you want your content published and any reason for that date e.g., a policy launch or awareness day. We aim to publish on the requested dates, but this may not be possible if it is short notice or there are competing internal announcements."
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Status" column
    $displayName = "Status"
    $field = Get-PnPField -Identity "ContentSubmissionStatus" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Choice" -InternalName "ContentSubmissionStatus" -DisplayName $displayName -Choices "Pending Approval","Approved","Rejected"

        Set-PnPField -Identity $field.Id -Values @{
            DefaultValue ="Pending Approval"
            Description = "Status of the content."; 
            CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Pending Approval"]},"sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont sp-field-fontSizeSmall sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved"]},"sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Rejected"]},"sp-css-backgroundColor-BgDustRose sp-css-borderColor-DustRoseFont sp-field-fontSizeSmall sp-css-color-DustRoseFont",""]}]}]}},"txtContent":"[$ContentSubmissionStatus]"}]}'
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Text" column
    $displayName = "Text"
    $field = Get-PnPField -Identity "ContentSubmissionDescription" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Note" -InternalName "ContentSubmissionDescription" -DisplayName $displayName -Required
        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else 
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    Write-Host ""

    # EVENT CONTENT-TYPE FIELDS 
    # "Event Date/Time" column
    $displayName = "Event Date/Time"
    $field = Get-PnPField -Identity "EventDateTime" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "DateTime" -InternalName "EventDateTime" -DisplayName $displayName -Required

        Set-PnPField -Identity $field.Id -Values @{
            FriendlyDisplayFormat = [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType]::Disabled;
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Venue and Joining Details" column
    $displayName = "Venue and Joining Details"
    $field = Get-PnPField -Identity "EventVenueAndJoiningDetails" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Note" -InternalName "EventVenueAndJoiningDetails" -DisplayName $displayName -Required
        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Event Details" column
    $displayName = "Details about the event"
    $field = Get-PnPField -Identity "EventDetails" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Note" -InternalName "EventDetails" -DisplayName $displayName -Required
        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

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

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript