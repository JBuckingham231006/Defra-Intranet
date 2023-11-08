<#
    SCRIPT OVERVIEW:
    This script creates our custom submission list

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

    $ctx = Get-PnPContext

    # Create new "Submission" list
    $displayName = "Internal Comms Intranet Content Submissions"
    $listURL = "Lists/ICICS"

    # Site-specific variable configuration.
    switch ($site.Abbreviation)
    {
        "Defra" { 
            $fieldNames = @("AltContact","ContentTypes","OrganisationIntranets","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription")
        }
        default { 
            $fieldNames = @("AltContact","ContentTypes","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription")
        }
    }

    Write-Host "`nCREATING THE LIST" -ForegroundColor Green

    # LIST - LIST CREATION
    $list = Get-PnPList -Identity $listURL

    if($null -eq $list)
    {
        $list = New-PnPList -Template GenericList -Title $displayName -Url $listURL -EnableVersioning -EnableContentTypes
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

    # SITE-SPECIFIC FIELDS
    switch ($site.Abbreviation)
    {
        "Defra" { 
            # Customise the OrganisationIntranets column for this list
            $field = Get-PnPField -List $list -Identity "OrganisationIntranets" -ErrorAction SilentlyContinue

            if($null -ne $field)
            {
                Set-PnPField -List $list -Identity $field.Id -Values @{
                    Title = "Content Relevant To"
                    Description = "Select whether the content is relevant to the whole of the Defra group or specific departments or functions"
                    Required = $true
                }
            }
            else
            {
                Write-Host "THE FIELD 'OrganisationIntranets' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
            }
        }
    }

    # LIST-LEVEL FIELD CUSTOMISATION
    Write-Host "`nCUSTOMISING FIELDS" -ForegroundColor Green

    # Customise the "ContentSubmissionStatus" column for this list
    $field = Get-PnPField -List $list -Identity "ContentSubmissionStatus" -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $list -Identity $field.Id -Values @{
            Hidden = $true
            CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Pending Approval"]},"sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont sp-field-fontSizeSmall sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved"]},"sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Rejected"]},"sp-css-backgroundColor-BgDustRose sp-css-borderColor-DustRoseFont sp-field-fontSizeSmall sp-css-color-DustRoseFont",""]}]}]}},"txtContent":"[$ContentSubmissionStatus]"}]}'
        }

        Write-Host "THE FIELD '$($field.Title)' HAS BEEN CUSTOMISED FOR THE LIST '$displayName'" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "THE FIELD 'ContentSubmissionStatus' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
    }

    # CONTENT TYPES
    Write-Host "`nCUSTOMISING CONTENT TYPES" -ForegroundColor Green

    $CTsToHide = New-Object System.Collections.ArrayList

    # Content Submission Request - Stage 2
    $ctName = "Content Submission Request - Stage 2"
    $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

    if($null -eq $listCT)
    {
        $ct = Get-PnPContentType -Identity $ctName

        if($null -ne $ct)
        {
            Add-PnPContentTypeToList -List $displayName -ContentType $ct
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName
            Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
        }
        else
        {
            throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'Create Content Types.ps1' script then try again." 
        }
    }
    else
    {
        Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
    }

    $ctx.Load($list.ContentTypes)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    # We'll hide this CT from the New menu, as it's only needed by Power Automate.
    $CTsToHide.Add($listCT.Id.StringValue) | Out-Null

    # Event Submission Request
    $ctName = "Event Submission Request"
    $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

    if($null -eq $listCT)
    {
        $ct = Get-PnPContentType -Identity $ctName

        if($null -ne $ct)
        {
            Add-PnPContentTypeToList -List $displayName -ContentType $ct
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName
            Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
        }
        else
        {
            throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'Create Content Types.ps1' script then try again." 
        }
    }
    else
    {
        Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
    }

    $ctx.Load($list.ContentTypes)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    # Event Submission Request - Stage 2
    $ctName = "Event Submission Request - Stage 2"
    $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

    if($null -eq $listCT)
    {
        $ct = Get-PnPContentType -Identity $ctName

        if($null -ne $ct)
        {
            Add-PnPContentTypeToList -List $displayName -ContentType $ct
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName
            Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
        }
        else
        {
            throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'Create Content Types.ps1' script then try again." 
        }
    }
    else
    {
        Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
    }

    $ctx.Load($list.ContentTypes)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    # We'll hide this CT from the New menu, as it's only needed by Power Automate.
    $CTsToHide.Add($listCT.Id.StringValue) | Out-Null
   
    if($null -eq $list.RootFolder.UniqueContentTypeOrder)
    {
        $contentTypesInPlace = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]'
        
        foreach($ct in $list.ContentTypes | where {$CTsToHide -notcontains $_.Id.StringValue -and $_.Name -ne "Folder"})
        {
            Write-Host "$($ct.Name) added the 'New' menu" -ForegroundColor Cyan
            $contentTypesInPlace.Add($ct.Id)
        }
    }
    else 
    {
        $contentTypesInPlace = [System.Collections.ArrayList] $list.RootFolder.UniqueContentTypeOrder
        $contentTypesInPlace = $contentTypesInPlace | where {$_.StringValue -ne $ct.Id.StringValue}
    }

    $list.RootFolder.UniqueContentTypeOrder = [System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]] $contentTypesInPlace
    $list.RootFolder.Update()             
    Invoke-PnPQuery

    # Rename default "Item" content type to "Content Submission Request"
    $ct = Get-PnPContentType -List $list -Identity "Item" -ErrorAction SilentlyContinue

    if($null -ne $ct)
    {
        $ctx = Get-PnPContext
        $ctx.Load($ct)
        $ctx.ExecuteQuery()

        try
        {
            $ct.ReadOnly = $false
            $ct.Update($false)
            $ctx.ExecuteQuery()

            $ct.Name = "Content Submission Request"
            $ct.Update($false)
            $ctx.ExecuteQuery()

            Write-Host "`nList default content type 'Item' renamed to 'Content Submission Request'" -ForegroundColor Green
        }
        finally
        {
            $ct.ReadOnly = $true
            $ct.Update($false)
            $ctx.ExecuteQuery()
        }
    }

    # VIEWS - Setup custom list views
    Write-Host "`nCUSTOMISING LIST VIEWS" -ForegroundColor Green

    switch ($site.Abbreviation)
    {
        "Defra" { 
            $fieldNames = @("AltContact","ContentTypes","OrganisationIntranets","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription","AssignedTo")
            $viewFields = "Attachments","LinkTitle","AssignedTo","ContentSubmissionDescription","Author","ContentSubmissionStatus","PublishBy","ContentTypes","AltContact","OrganisationIntranets","LineManager","StakeholdersInformed"
        }
        default { 
            $fieldNames = @("AltContact","ContentTypes","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription","AssignedTo")
            $viewFields = "Attachments","LinkTitle","AssignedTo","ContentSubmissionDescription","Author","ContentSubmissionStatus","PublishBy","ContentTypes","AltContact","LineManager","StakeholdersInformed"
        }
    }

    $viewConfiguration = @(
        [PSCustomObject]@{
            'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq></Where>'
            'TargetSite' = ''
            'Title' = 'All Pending Submissions'
        },
        [PSCustomObject]@{
            'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><Or><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Approved</Value></Eq><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Eq></Or></Where>'
            'TargetSite' = ''
            'Title' = 'All Processed Submissions'
        },
        [PSCustomObject]@{
            'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy>'
            'TargetSite' = ''
            'Title' = 'By Assigned To'
        },
        [PSCustomObject]@{
            'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="ContentTypes" /></GroupBy>'
            'TargetSite' = ''
            'Title' = 'By Content Types'
        },
        [PSCustomObject]@{
            'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="PublishBy" /></OrderBy><Where><And><Geq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today /></Value></Geq><Leq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And></Where>'
            'TargetSite' = ''
            'Title' = 'Due in the next 7 days'
        }
    )

    foreach($viewConfig in $viewConfiguration)
    {
        # If this is a view for a specific-site and the site we're on is not that site then we skip
        if($viewConfig.TargetSite.Length -gt 0 -and $viewConfig.TargetSite -ne $site.Abbreviation)
        {
            continue;
        }

        $title = $viewConfig.viewTitle
        $view = Get-PnPView -List $list -Identity $viewConfig.Title -ErrorAction SilentlyContinue

        if($null -eq $view)
        {
            $view = Add-PnPView -List $list -Title $viewConfig.Title -Fields $viewFields -Query $viewConfig.Query
            Write-Host "VIEW '$($viewConfig.Title)' ADD TO THE LIST" -ForegroundColor Green
        }
        else
        {
            Write-Host "THE VIEW '$($viewConfig.Title)' ALREADY EXISTS" -ForegroundColor Yellow
        }
    }

    $view = Get-PnPView -List $list -Identity "All Items"

    if($null -ne $view)
    {
        $view = Set-PnPView -List $list -Identity $view.Title -Fields $viewFields
        Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
    }
  
    # Set unique permissions for the list so anyone on the site can add an item
    if($null -ne $site.GroupPrefix -and $site.GroupPrefix.Length -gt 0)
    {
        Write-Host "`nCUSTOMISING LIST PERMISSIONS" -ForegroundColor Green
        Set-PnpList -Identity $list -BreakRoleInheritance
    
        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Owners" -AddRole "Full Control"
        Write-Host "'$($site.GroupPrefix) Owners' given Full Control" -ForegroundColor Yellow
    
        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Members" -AddRole "Edit"
        Write-Host "'$($site.GroupPrefix) Members' given Edit permissions to the list" -ForegroundColor Yellow

        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Visitors" -AddRole "Contribute"
        Write-Host "'$($site.GroupPrefix) Visitors' given Contribute permissions to the list" -ForegroundColor Yellow
    }

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript