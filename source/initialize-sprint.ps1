function Write-Log
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    param 
    (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info","Log")]
        [string]$Level="Log",

        [Parameter(Mandatory=$false)]
        [switch][bool]$NoNewLine=$false,

        [Parameter(Mandatory=$false)]
        [switch][bool]$Quiet=$false,

        [Parameter(Mandatory=$false)]
        [switch][bool]$OutToFile=$false,

        [Parameter(Mandatory=$false)]
        [string]$ForegroundColor = "Gray"
    )

    Begin
    {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'

        if ($null -eq $ForegroundColor -or $ForegroundColor -eq "")
        {
            $ForegroundColor = "Gray"
        }
    }
    
    Process 
    {
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $FormattedShortDate = Get-Date -Format "yyyy-MM-dd"
        $logPath = Join-Path $PSScriptRoot "log_$FormattedShortDate.log"

        switch ($Level) 
        {
            "Error"
                {
                    Write-Error $Message
                    $LevelText = "ERROR:"
                }
            "Warn"
                {
                    Write-Warning $Message
                    $LevelText = "WARNING:"
                }
            "Info"
                {
                    Write-Verbose $Message
                    $LevelText = "INFO:"
                }
            "Log"
                {
                    if (!$Quiet)
                    {
                        Write-Host "$FormattedDate - $LevelText $Message" -ForegroundColor:$ForegroundColor
                    }
                }
        }

        if ($OutToFile)
        {
            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logPath -Append
        }
    }
    End 
    {

    }
}

function Get-Url
{
    param
    (
       [string]$Url,
       [hashtable]$Header,
       [string]$AreaId
    )
    
    $orgResourceUrl = "$Url/_apis/resourceAreas/$($AreaId)"
    $res = Invoke-RestMethod -Uri $orgResourceUrl -Headers $header

    if ($null -ne $res) 
    {
        return $res.locationUrl
    } 
    return $Url
}

function Get-Iteration
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [string]$ParentPath
    )
    $testIterationUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/classificationnodes/Iterations/$($ParentPath)?{0}=3&api-version=$($ProjectConfig.ApiVersion)"
    $testIterationUrl = [string]::Format($testIterationUrl, '$depth')
    $iterations = Invoke-RestMethod -Uri $testIterationUrl -Method Get -ContentType "application/json" -Headers $($ProjectConfig.Header)
    
    if ($iterations.hasChildren)
    {
        foreach ($iteration in $iterations.children)
        {
            if ($iteration.name -eq $Name)
            {
                return $iteration
            }
        }
    }
    return $null
}

function Test-Iteration
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [string]$ParentPath
    )
    
    $testIterationUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/classificationnodes/Iterations/$($ParentPath)?{0}=3&api-version=$($ProjectConfig.ApiVersion)"
    $testIterationUrl = [string]::Format($testIterationUrl, '$depth')
    
    $iterations = Invoke-RestMethod -Uri $testIterationUrl -Method Get -ContentType "application/json" -Headers $($ProjectConfig.Header)
    if ($iterations.hasChildren)
    {
        foreach ($iteration in $iterations.children)
        {
            if ($iteration.name -eq $Name)
            {
                return $true
            }
        }
    }
    return $false
}

function New-Iteration
{
    param
    (
        [object]$ProjectConfig,
        [string]$ParentPath,
        [object]$IterationConfig
    )

    $newIterationUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/classificationnodes/iterations/$($ParentPath)?api-version=$($ProjectConfig.ApiVersion)"
    $iteration = "{""name"":""$($IterationConfig.IterationName)"", ""StructureType"": ""iteration"", ""hasChildren"":false, ""attributes"":{ ""startDate"":""$($IterationConfig.StartDate)T00:00:00Z"", ""finishDate"":""$($IterationConfig.FinishDate)T00:00:00Z""}}"
        
    [void](Invoke-RestMethod -Uri $newIterationUrl -Method Post -ContentType "application/json" -Headers $($ProjectConfig.Header) -Body $iteration)
}

function Add-IterationTeam
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [string]$ParentPath
    )

    $iteration = Get-Iteration -ProjectConfig $ProjectConfig -Name $Name -ParentPath $ParentPath

    if ($iteration)
    {
        $addIterationTeamUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/work/teamsettings/iterations/?api-version=$($ProjectConfig.ApiVersion)" 
        $iterationBody = "{""id"":""$($iteration.identifier)""}"
        try 
        {
            [void](Invoke-RestMethod -Uri $addIterationTeamUrl -Method Post -ContentType "application/json" -Headers $Header -Body $iterationBody)
            return $true
        }
        catch 
        {
            return $false    
        }
    }    
    return $false
}

function Test-WorkItem
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [string]$WorkItemTypeLabel
    )
    
    $workItems = Get-WorkItem -ProjectConfig $ProjectConfig -Name $Name -WorkItemTypeLabel $WorkItemTypeLabel

    if ($workItems)
    {
        return $true
    }
    return $false
}

function Get-WorkItem
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [string]$WorkItemTypeLabel
    )
    
    $getWorkItemUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/wiql?api-version=$($ProjectConfig.ApiVersion)"
    $query = "{ ""Query"": ""Select [System.Id] From WorkItems Where [System.WorkItemType] = '$WorkItemTypeLabel' AND [State] <> 'Removed' AND [System.Title] = '$Name'"" }"
    $workItems = Invoke-RestMethod -Uri $getWorkItemUrl -Method Post -ContentType "application/json" -Headers $($ProjectConfig.Header) -Body $query

    return $workItems.workItems
}

function Get-WorkItemById
{
    param
    (
        [object]$ProjectConfig,
        [int]$Id
    )
    
    $getWorkItemUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/workitems/$($Id)?api-version=$($ProjectConfig.ApiVersion)"
    $workItems = Invoke-RestMethod -Uri $getWorkItemUrl -Method Get -ContentType "application/json" -Headers $($ProjectConfig.Header)

    return $workItems
}

function New-WorkItem
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [object]$IterationConfig,
        [string]$WorkItemType,
        [string]$State
    )

    $newWorkitemUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/workitems/{0}?api-version=5.1" #$($ProjectConfig.ApiVersion)"
    $newWorkitemUrl = [string]::Format($newWorkitemUrl, [string]::Concat("$", $WorkItemType))

    $iterationPath = "$($ProjectConfig.Project)\\$($IterationConfig.ParentIterationName)\\$($IterationConfig.IterationName)"
    $areaPath = $IterationConfig.AreaPath -replace "/", "\\"

    $workItemDefinition = "[
        { ""op"": ""add"",""path"": ""/fields/System.Title"",""from"": null,""value"": ""$Name"" },
        { ""op"": ""add"",""path"": ""/fields/System.AreaPath"",""from"": null,""value"": ""$areaPath"" },
        { ""op"": ""add"",""path"": ""/fields/System.State"",""from"": null,""value"": ""$State"" },
        { ""op"": ""add"",""path"": ""/fields/System.IterationPath"",""from"": null,""value"": ""$iterationPath"" }
    ]"
        
    try {
        Invoke-RestMethod -Uri $newWorkitemUrl -Method Post -ContentType "application/json-patch+json" -Headers $($ProjectConfig.Header) -Body $workItemDefinition
        return $true
    }
    catch {
        Write-Log -level "Error" -Message $error
        return $false
    }
}

function New-TaskLink
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [int]$ParentWorkItemId,
        [object]$IterationConfig
    )

    $newWorkitemUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/workitems/{0}?api-version=5.1" #$($ProjectConfig.ApiVersion)"
    $newWorkitemUrl = [string]::Format($newWorkitemUrl, [string]::Concat("$", "Task"))

    $iterationPath = "$($ProjectConfig.Project)\\$($IterationConfig.ParentIterationName)\\$($IterationConfig.IterationName)"
    $areaPath = $IterationConfig.AreaPath -replace "/", "\\"
    $baseUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)"

    $workItemDefinition = "[
        { ""op"": ""add"",""path"": ""/id"",""from"": null,""value"": ""-1"" },
        { ""op"": ""add"",""path"": ""/fields/System.Title"",""from"": null,""value"": ""$Name"" },
        { ""op"": ""add"",""path"": ""/fields/System.AreaPath"",""from"": null,""value"": ""$areaPath"" },
        { ""op"": ""add"",""path"": ""/fields/System.State"",""from"": null,""value"": ""To Do"" },
        { ""op"": ""add"",""path"": ""/fields/System.IterationPath"",""from"": null,""value"": ""$iterationPath"" },
        {
            ""op"": ""add"",
            ""path"": ""/relations/-"",
            ""value"": {
              ""rel"": ""System.LinkTypes.Hierarchy-Reverse"",
              ""url"": ""$baseUrl/_apis/wit/workItems/$ParentWorkItemId"",
              ""attributes"": {
                ""comment"": ""Making a new link for the dependency""
              }
            }
        }
    ]"
    
    try {
        
        [void](Invoke-RestMethod -Uri $newWorkitemUrl -Method Post -ContentType "application/json-patch+json" -Headers $($ProjectConfig.Header) -Body $workItemDefinition)
        return $true
    }
    catch {
        Write-Log -level "Error" -Message $error
        return $false
    }
}

function Test-TaskLink
{
    param
    (
        [object]$ProjectConfig,
        [string]$Name,
        [int]$ParentWorkItemId
    )
    
    $relationsUrl = "$($ProjectConfig.AzDOsUrl)$($ProjectConfig.Project)/_apis/wit/workitems/{0}?{1}&api-version=5.1"
    $relationsUrl = [string]::Format($relationsUrl, $ParentWorkItemId, [string]::Concat("$", "expand=Relations"))
    $taskLinks = Invoke-RestMethod -Uri $relationsUrl -Method Get -ContentType "application/json" -Headers $($ProjectConfig.Header)
    
    if ($taskLinks.relations.count -lt 1)
    {
        return $false
    }
    
    foreach ($taskLink in $taskLinks.relations) 
    {
        if ($taskLink.rel -eq "System.LinkTypes.Hierarchy-Forward")
        {
            $childId = Split-Path $taskLink.url -Leaf
            $children = Get-WorkItemById -ProjectConfig $ProjectConfig -Id $childId
            $child = $children.fields | Where-Object -Property "System.Title" -eq $Name
            if ($child)
            {
                return $true
            }
        }
    }
    return $false
}

function New-Sprint
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param 
    (
        [Parameter(Mandatory=$false)]
        [switch][bool]$Quiet=$false
    )

    Begin
    {
        ## GLOBAL
        $orgUrl = "https://dev.azure.com/xyz"
        $coreAreaId = "79134c72-4a58-4b42-976c-04e7115f32bf"
        $pat = "xxxxxxxxxxxxxxxxxxxx"
        $token = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$($pat)"))
        $header = @{authorization = "Basic $token"}
        $azDOsBaseUrl = Get-Url -Url $orgUrl -Header $header -AreaId $coreAreaId
        ##
        
        # inputs
        $project = Read-Host "Project"
        if ($null -eq $project -or $project -eq "")
        {
            $project = "YourProject"
        }

        $team = Read-Host "Team"
        if ($null -eq $team -or $team -eq "")
        {
            $team = "YourTeam"
        }

        $areaPath = Read-Host "Area Path"
        if ($null -eq $areaPath -or $areaPath -eq "")
        {
            $areaPath = "YourArea"
        }
        
        $parentIterationName = Read-Host "Parent iteration name"
        if ($null -eq $parentIterationName -or $parentIterationName -eq "")
        {
            $parentIterationName = ""
        }

        $iterationName = Read-Host "Iteration name"
        if ($null -eq $iterationName -or $iterationName -eq "")
        {
            $iterationName = "YourIteration"
        }

        $prevIterationName = Read-Host "Previous iteration name"
        if ($null -eq $prevIterationName -or $prevIterationName -eq "")
        {
            $prevIterationName = $iterationName
        }

        $startDate = Read-Host "Start date (YYYY-MM-DD)"
        $finishDate = Read-Host "Finish date (YYYY-MM-DD)"
        
        $objectProperty = [hashtable]@{
            AzDOsUrl = $azDOsBaseUrl
            Header = $header
            Project = $project
            Team = $team
            ApiVersion = "5.0"
        }
        $GlobalConfig = New-Object -TypeName psobject -Property $objectProperty
        
        $objectProperty = [hashtable]@{
            AreaPath = $areaPath
            ParentIterationName = $parentIterationName
            IterationName = $iterationName
            PreviousIterationName = $prevIterationName
            StartDate = $startDate
            FinishDate = $finishDate
        }
        $IterationConfig = New-Object -TypeName psobject -Property $objectProperty
    }

    Process {

        $iterationExists = Test-Iteration -ProjectConfig $GlobalConfig -Name $($IterationConfig.IterationName) -ParentPath $parentIterationName
        if (!$iterationExists)
        {
            Write-Log -Level "Log" -Message "Creating iteration ""$($IterationConfig.IterationName)"".." -Quiet:$Quiet
            New-Iteration -ProjectConfig $GlobalConfig -ParentPath $parentIterationName -IterationConfig $IterationConfig
        }
        else 
        {
            Write-Log -Level "Log" -Message """$($IterationConfig.IterationName)"" already exists. No user action required." -Quiet:$Quiet
        }

        Write-Log -Level "Log" -Message "Adding iteration ""$($IterationConfig.IterationName)"" to the ""$($GlobalConfig.Team)"" team.." -Quiet:$Quiet
        $iterationTeamAdded = Add-IterationTeam -ProjectConfig $GlobalConfig -Name $($IterationConfig.IterationName) -ParentPath $parentIterationName
        if (!$iterationTeamAdded)
        {
            Write-Log -Level "Error" -Message "Something went wrong. Quitting.."
            exit
        }

        # PARENT ITEM
        $workItemName = "$($IterationConfig.IterationName) - BUG xyz"
        $workItemExists = Test-WorkItem -ProjectConfig $GlobalConfig -Name $workItemName -WorkItemTypeLabel "Bug"
        if (!$workItemExists)
        {
            Write-Log -Level "Log" -Message "Creating work item ""$workItemName"".." -Quiet:$Quiet

            $workItemAdded = New-WorkItem -ProjectConfig $GlobalConfig -Name $workItemName -IterationConfig $IterationConfig -WorkItemType "Bug" -State "Committed"
            if (!$workItemAdded)
            {
                Write-Log -Level "Error" -Message "Something went wrong. Quitting.."
                exit
            }
        }
        else 
        {
            Write-Log -Level "Log" -Message """$workItemName"" already exists. No user action required." -Quiet:$Quiet
        }

        # TASK
        $childWorkItemName = "Fix"        
        $parentWorkItem = Get-WorkItem -ProjectConfig $GlobalConfig -Name $workItemName -WorkItemTypeLabel "Bug" -IterationConfig $IterationConfig
        if (!$parentWorkItem)
        {
            Write-Log -Level "Error" -Message "Something went wrong. Quitting.."
            exit
        } 

        $workItemExists = Test-TaskLink -ProjectConfig $GlobalConfig -Name $childWorkItemName -ParentWorkItemId $parentWorkItem.id
        if (!$workItemExists)
        {
            Write-Log -Level "Log" -Message "Creating task ""$childWorkItemName"" for ""$workItemName"".." -Quiet:$Quiet
            $workItemAdded = New-TaskLink -ProjectConfig $GlobalConfig -Name $childWorkItemName -ParentWorkItemId $parentWorkItem.id -IterationConfig $IterationConfig
            if (!$workItemAdded)
            {
                Write-Log -Level "Error" -Message "Something went wrong. Quitting.."
                exit
            }
        }
        else 
        {
            Write-Log -Level "Log" -Message """$childWorkItemName"" for ""$workItemName"" already exists. No user action required." -Quiet:$Quiet
        }        
    }

    End {
    
    }
}

Clear-Host
# script folder
New-Sprint