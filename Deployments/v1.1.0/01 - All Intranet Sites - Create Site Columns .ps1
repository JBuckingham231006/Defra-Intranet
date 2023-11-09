<#
    SCRIPT OVERVIEW:
    This script creates the site columns required by our custom list(s) and libraries within the Defra and ALB SharePoint sites

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
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}

foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
    Write-Host ""

    # "Alternative Contact" column
    $displayName = "Alternative Contact"
    $field = Get-PnPField -Identity "AltContact" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "User" -InternalName "AltContact" -DisplayName $displayName -Required
        Set-PnPField -Identity $field.Id -Values @{SelectionMode=0}

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
        $field = Add-PnPField -Type "Choice" -InternalName "ContentTypes" -DisplayName $displayName -Required -Choices "Blog or Online Diary","Form","Guidance Page","News Story","Office Notice"

        Set-PnPField -Identity $field.Id -Values @{
            Description = "Please select what kind of content you are submitting:"; 
            CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]",""]},"",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","News story"]},"sp-css-backgroundColor-successBackground50 sp-css-color-green",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Blog or online diary"]},"sp-css-backgroundColor-warningBackground50 sp-css-color-neutralPrimaryAlt",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Office notice"]},"sp-css-backgroundColor-BgLightPurple sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-LightPurpleFont sp-css-color-LightPurpleFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Guidance Page"]},"sp-css-backgroundColor-BgCyan sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-CyanFont sp-css-color-CyanFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Blog or Online Diary"]},"sp-css-backgroundColor-BgGold sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-GoldFont sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Form"]},"sp-css-backgroundColor-BgDustRose sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-DustRoseFont sp-css-color-DustRoseFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","News Story"]},"sp-css-backgroundColor-BgMintGreen sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-MintGreenFont sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Office Notice"]},"sp-css-backgroundColor-BgLightPurple sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-LightPurpleFont sp-css-color-LightPurpleFont",""]}]}]}]}]}]}]}]}]}},"txtContent":"[$ContentTypes]"}],"templateId":"BgColorChoicePill"}'
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
            Description = "Please tell us who in your senior management provided the final sign-off on this story and confirm all relevant stakeholders have been informed:"
        }

        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else 
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Preferred Timing" column
    $displayName = "Preferred Timing"
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

    # "Stakeholders Informed" column
    $displayName = "Stakeholders Informed"
    $field = Get-PnPField -Identity "StakeholdersInformed" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Choice" -InternalName "StakeholdersInformed" -DisplayName $displayName -Required -Choices "Yes","No"
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

    # "Booking" column
    $displayName = "Booking"
    $field = Get-PnPField -Identity "EventBooking" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Note" -InternalName "EventBooking" -DisplayName $displayName -Required
        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }

    # "Further information for the reader" column
    $displayName = "Further information for the reader"
    $field = Get-PnPField -Identity "EventFurtherInformation" -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        $field = Add-PnPField -Type "Note" -InternalName "EventFurtherInformation" -DisplayName $displayName -Required
        Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript