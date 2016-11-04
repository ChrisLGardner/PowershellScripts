function New-TfsTeamProjectDashboardBuildChart
{
    <#  
        .SYNOPSIS
            This function will create a chart for the specified build and add it to the dashboard.

        .DESCRIPTION
            This function will create a chart for the specified build and add it to the dashboard.

            First creates a blank chart object on the dashboard and then updates it to use the specified build definition.
        .PARAMETER WebSession
            Web session object for the target TFS server.

        .PARAMETER Team
            The name of the team

        .PARAMETER Project
            The name of the project under which the team can be found

        .PARAMETER Dashboard
            The name of the dashboard for the team project

        .PARAMETER BuildName
            The build definition name

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            Username to connect to TFS with

        .PARAMETER AccessToken
            AccessToken for VSTS to connect with.

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE 
            New-TfsTeamProjectDashboardBuildChart -WebSession $session -Team 'Engineering' -Project 'Super Product' -Dashboard 'Overview' -BuildName 'Project1.Build.CI'

            This will add a new build chart to the Overview dashboard for the Project1.Build.CI build definition using the specified web session.

        .EXAMPLE
            New-TfsTeamProjectDashboardBuildChart -Team 'Engineering' -Project 'Super Product' -Dashboard 'Overview' -BuildName 'Project1.Build.CI' -Uri 'https://product.visualstudio.com/DefaultCollection' -Username 'MainUser' -AccessToken (Get-Content c:\accesstoken.txt | Out-String)

            This will add a new build chart to the Overview dashboard for the Project1.Build.CI build definition on the target VSTS account using the provided creds.

    #>
    [cmdletbinding()]
    param (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Team,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Dashboard,

        [Parameter(Mandatory)]
        [String]$BuildName,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials

    )
    Process
    {
        $headers = @{'Content-Type'='application/json';'accept'='api-version=2.2-preview.1;application/json'}
        $Parameters = @{}

        #Use Hashtable to create param block for invoke-restmethod and splat it together
        switch ($PsCmdlet.ParameterSetName) 
        {
            'SingleConnection'
            {
                $WebSession = Connect-TfsServer -Uri $uri -Username $Username -AccessToken $AccessToken
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)

            }
            'LocalConnection'
            {
                $WebSession = Connect-TfsServer -uri $Uri -UseDefaultCredentials
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
            }
            'WebSession'
            {
                $Uri = $WebSession.uri
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
                #Connection details here from websession, no creds needed as already there
            }
        }

        $DashboardObject = Get-TfsTeamProjectDashboard -WebSession $WebSession -Team $Team -Project $Project | Where-Object Name -eq $Dashboard
        $BuildObject = Get-TfsBuildDefinition -WebSession $WebSession -Project $Project | Where-Object Name -eq $BuildName

        $Uri = "$($DashboardObject.url)/widgets"
        $Parameters.add('Uri',$uri)

        $BuildWidgetJson = @"
{
    "isEnabled": true,
    "_links": null,
    "contentUri": null,
    "contributionId": "ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget",
    "configurationContributionId": "ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget.Configuration",
    "isNameConfigurable": true,
    "url": null,
    "name": "Chart for Build History",
    "id": null,
    "size": {
        "rowSpan": 1,
        "columnSpan": 2
    },
    "position": {
        "column": 0,
        "row": 0
    },
    "settings": null,
    "typeId": "Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget"
}
"@

        try
        {
          $JsonData = Invoke-RestMethod @Parameters -Method Post -Body $BuildWidgetJson -ErrorAction Stop
    
        }
        catch
        {
          Write-Error "Error was $_"
          $line = $_.InvocationInfo.ScriptLineNumber
          Write-Error "Error was in Line $line"
        }
    
        $Parameters.uri = $JsonData.Url
        
        $UpdateBuild = @"
{
    "url": null,
    "id": "$($jsondata.id)",
    "name": "$BuildName",
    "position": {
        "row": 0,
        "column": 0
    },
    "size": {
        "rowSpan": 1,
        "columnSpan": 2
    },
    "settings": "{\"name\":\"$BuildName\",\"projectId\":\"$($BuildObject.Project.id)\",\"id\":$($BuildObject.id),\"type\":$(if ($BuildObject.Type -eq 'xaml') { '1'} else { '2'}) ,\"uri\":\"$($BuildObject.Uri)\",\"providerName\":\"Team favorites\",\"lastArtifactName\":\"$($BuildObject.Name)\"}",
    "artifactId": "",
    "isEnabled": true,
    "contentUri": null,
    "contributionId": "ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget",
    "typeId": "Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget",
    "configurationContributionId": "ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.BuildChartWidget.Configuration",
    "isNameConfigurable": true,
    "loadingImageUrl": "https://source.blackmarble.co.uk/tfs/_static/Widgets/sprintBurndown-buildChartLoading.png",
    "allowedSizes": [{
        "rowSpan": 1,
        "columnSpan": 2
    }]
}
"@
        try
        {
          $UpdatedBuildJson = Invoke-RestMethod @Parameters -Method Patch -Body $UpdateBuild -ErrorAction Stop
        }
        catch
        {
          Write-Error "Error was $_"
          $line = $_.InvocationInfo.ScriptLineNumber
          Write-Error "Error was in Line $line"
        }

        Write-Output $UpdatedBuildJson
    }
}


function New-TfsTeamProjectDashboardWorkItemQuery
{
    <#  
        .SYNOPSIS
            This function will add a work item query to a dashboard.

        .DESCRIPTION
            This function will add a work item query to a dashboard, either using an existing query or a new query when passed a wiql string.

            First creates a object on the dashboard and then updates it to use the specified query.
        .PARAMETER WebSession
            Web session object for the target TFS server.

        .PARAMETER Team
            The name of the team

        .PARAMETER Project
            The name of the project under which the team can be found

        .PARAMETER Dashboard
            The name of the dashboard for the team project

        .PARAMETER QueryPath
            The path to the query, either existing or new, including folders and seperated with /'s
            
        .PARAMETER Query
            The wiql query string to create and use for the widget

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            Username to connect to TFS with

        .PARAMETER AccessToken
            AccessToken for VSTS to connect with.

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE 
            New-TfsTeamProjectDashboardWorkItemQuery -WebSession $session -Team 'Engineering' -Project 'Super Product' -Dashboard 'Overview' -QueryPath 'Shared Queries/Assigned To You'

            This will add a new query to the Overview dashboard for the Assigned To You work item query using the specified web session.

        .EXAMPLE
            $WiqlString = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State],[System.Tags] FROM WorkItemLinks WHERE ([Source].[System.TeamProject] = @project AND ( [Source].[System.WorkItemType] = 'Product Backlog Item' OR [Source].[System.WorkItemType] = 'Bug' ) AND [Source].[System.State] <> 'Done' AND [Source].[System.State] <> 'Removed' AND [Source].[System.IterationPath] = @currentIteration) AND ([Target].[System.TeamProject] = @project AND [Target].[System.WorkItemType] = 'Task' AND [Target].[System.AssignedTo] = @me AND [Target].[System.State] <> 'Done' AND [Target].[System.State] <> 'Removed' AND [Target].[System.IterationPath] = @currentIteration) mode(MustContain)"
            New-TfsTeamProjectDashboardWorkItemQuery -Team 'Engineering' -Project 'Super Product' -Dashboard 'Overview' -QueryPath 'Shared Queries/Assigned For Sprint' -Query $WiqlString -Uri 'https://product.visualstudio.com/DefaultCollection' -Username 'MainUser' -AccessToken (Get-Content c:\accesstoken.txt | Out-String)

            This will add a new query to the dashboard using the specified query string on the target VSTS account using the provided creds.
    #>
    [cmdletbinding()]
    param (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Team,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Dashboard,

        [Parameter(Mandatory)]
        [String]$QueryPath,

        [String]$Query,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials

    )
    Process
    {
        $headers = @{'Content-Type'='application/json';'accept'='api-version=2.2-preview.1;application/json'}
        $Parameters = @{}

        #Use Hashtable to create param block for invoke-restmethod and splat it together
        switch ($PsCmdlet.ParameterSetName) 
        {
            'SingleConnection'
            {
                $WebSession = Connect-TfsServer -Uri $uri -Username $Username -AccessToken $AccessToken
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)

            }
            'LocalConnection'
            {
                $WebSession = Connect-TfsServer -uri $Uri -UseDefaultCredentials
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
            }
            'WebSession'
            {
                $Uri = $WebSession.uri
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
                #Connection details here from websession, no creds needed as already there
            }
        }
        
        $QueryPath = $QueryPath -replace '\\','/'

        #Either create the new Query or get the details for the existing one
        if ($Query)
        {
            
            New-TfsWorkItemQuery -WebSession $WebSession -Project $Project -Folder $QueryFolder -Name $QueryName -Wiql $Query
            <#            
            $QueryName = ($QueryPath -split '/')[-1]
            $QueryFolder = ($QueryPath -split '/' | Select-Object -SkipLast 1) -join '/'
            $NewQueryBody = @{ Name = $QueryName; wiql = $Query} | ConvertTo-Json -Depth 10

            $QueryUrl = "$($WebSession.Uri)/$Project/_apis/wit/queries/$($QueryFolder -replace ' ','%20')?api-version=1.0"
            try
            {
                $JsonData = Invoke-Restmethod -Uri $QueryUrl @Parameters -Method Post -Body $NewQueryBody -ErrorAction Stop
            
            }
            catch
            {
                Write-Error "Error was $_"
                $line = $_.InvocationInfo.ScriptLineNumber
                Write-Error "Error was in Line $line"
            }#>
        }
        else
        {
            $QueryUrl = "$($WebSession.Uri)/$Project/_apis/wit/queries/$($QueryPath -replace ' ','%20')?api-version=1.0"

            try
            {
                $JsonData = Invoke-Restmethod -Uri $QueryUrl @Parameters -ErrorAction Stop
            
            }
            catch
            {
                Write-Error "Error was $_"
                $line = $_.InvocationInfo.ScriptLineNumber
                Write-Error "Error was in Line $line"
            }
        }
        
        #Add the widget to the dashboard and then update it to use the correct query
        $NewQueryTileJson = @"
{
    "isEnabled":true,
    "_links":null,
    "contentUri":null,
    "contributionId":"ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget",
    "configurationContributionId":"ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget.Configuration",
    "isNameConfigurable":true,
    "url":null,
    "name":"Query Tile",
    "id":null,
    "size":{
        "rowSpan":1,
        "columnSpan":1
    },
    "position":{
        "column":0,
        "row":0
    },
    "settings":null,
    "typeId":"Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget"
}
"@
        $Uri = "{0}/widgets" -f (Get-TfsTeamProjectDashboard -WebSession $WebSession -Team $Team -Project $Project | Where-Object Name -eq $Dashboard | select-object -ExpandProperty url)
        
        try
        {
          $NewWidget = Invoke-RestMethod -uri $uri @parameters -Method Post -Body $NewQueryTileJson
        
        }
        catch
        {
          "Error was $_"
          $line = $_.InvocationInfo.ScriptLineNumber
          "Error was in Line $line"
        }
        
        $UpdateQueryJson = @"
{
    "url":null,
    "id":"$($NewWidget.id)",
    "name":"$($JsonData.Name)",
    "position":{
        "row":0,
        "column":0
    },
    "size":{
        "rowSpan":1,
        "columnSpan":1
    },
    "settings":"{\"queryId\":\"$($JsonData.id)\",\"queryName\":\"$($JsonData.Name)\",\"colorRules\":[{\"isEnabled\":false,\"backgroundColor\":\"#339933\",\"thresholdCount\":10,\"operator\":\"<=\"},{\"isEnabled\":false,\"backgroundColor\":\"#E51400\",\"thresholdCount\":20,\"operator\":\">\"}],\"lastArtifactName\":\"$($JsonData.name)\"}",
    "artifactId":"",
    "isEnabled":true,
    "contentUri":null,
    "contributionId":"ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget",
    "typeId":"Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget",
    "configurationContributionId":"ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.QueryScalarWidget.Configuration",
    "isNameConfigurable":true,
    "loadingImageUrl":"$($Websession.uri)/_static/Widgets/scalarLoading.png",
    "allowedSizes":[{
        "rowSpan":1,
        "columnSpan":1
    }]
}
"@
        try
        {
          $UpdateQuery = Invoke-RestMethod -Uri $NewWidget.url -Method Patch -Body $UpdateQueryJson -Headers @{'Content-Type'='application/json; charset=utf-8; api-version=2.2-preview.1'}
        }
        catch
        {
          "Error was $_"
          $line = $_.InvocationInfo.ScriptLineNumber
          "Error was in Line $line"
        }
    }
}

Function Add-TfsTeamProjectDashboardSprintWidget
{
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Team,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Dashboard,

        [ValidateSet('Burndown','Capacity','Overview')]
        [String]$SprintWidget,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials

    )
    Process
    {
        $headers = @{'Content-Type'='application/json';'accept'='api-version=2.2-preview.1;application/json'}
        $Parameters = @{}

        #Use Hashtable to create param block for invoke-restmethod and splat it together
        switch ($PsCmdlet.ParameterSetName) 
        {
            'SingleConnection'
            {
                $WebSession = Connect-TfsServer -Uri $uri -Username $Username -AccessToken $AccessToken
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)

            }
            'LocalConnection'
            {
                $WebSession = Connect-TfsServer -uri $Uri -UseDefaultCredentials
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
            }
            'WebSession'
            {
                $Uri = $WebSession.uri
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
                #Connection details here from websession, no creds needed as already there
            }
        }

        $SprintWidgetJson = @"
{
    "isEnabled":true,
    "_links":null,
    "contentUri":null,
    "contributionId":"ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.Sprint$($SprintWidget)Widget",
    "configurationContributionId":null,
    "isNameConfigurable":false,
    "url":null,
    "name":"Sprint $SprintWidget",
    "id":null,
    "size":{
        "rowSpan":1,
        "columnSpan":2
    },
    "position":{
        "column":0,
        "row":0
    },
    "settings":null,
    "typeId":"Microsoft.VisualStudioOnline.Dashboards.Sprint$($SprintWidget)Widget"
}
"@

        $DashboardObject = Get-TfsTeamProjectDashboard -WebSession $WebSession -Team $Team -Project $Project

        $Uri = "$($DashboardObject.url)/widgets"
        $Parameters.Add('uri',$uri)

        try
        {
            $JsonData = Invoke-RestMethod @Parameters -Method Post -Body $SprintWidgetJson -ErrorAction Stop
        
        }
        catch
        {
            Write-Error "Error was $_"
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Error "Error was in Line $line"
        }
    }
}

function Get-TfsTeamProjectDashboard
{
    <#  
        .SYNOPSIS
            This function will return all dashboards associated with a specific team project.

        .DESCRIPTION
            This function will return all dashboards associated with a specific team project.

            The function will take either a websession object or a uri and
            credentials. The web session can be piped to the fuction from the
            Connect-TfsServer function.

        .PARAMETER WebSession
            Websession with connection details and credentials generated by Connect-TfsServer function

        .PARAMETER Team
            The name of the team who's dashboards should be returned

        .PARAMETER Project
            The name of the project containing the dashboard

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE
            Get-TfsTeamProjectDashboard -WebSession $Session -Team 'Engineering' -Project 'Super Product'

            This will return any dashboards that are on the Super Product project and linked to the Engineering team 
            using the already established session.

        .EXAMPLE
            Get-TfsTeamProjectDashboard -Uri 'https://test.visualstudio.com/defaultcollection'  -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) -Team 'Engineering' -Project 'Super Product'

            This will return any dashboards that are on the Super Product project and linked to the Engineering team 
            using the provided credentials and uri.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Team,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials

    )
    Process
    {
        $headers = @{'Content-Type'='application/json-patch+json'}
        $Parameters = @{}

        #Use Hashtable to create param block for invoke-restmethod and splat it together
        switch ($PsCmdlet.ParameterSetName) 
        {
            'SingleConnection'
            {
                $WebSession = Connect-TfsServer -Uri $uri -Username $Username -AccessToken $AccessToken
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)

            }
            'LocalConnection'
            {
                $WebSession = Connect-TfsServer -uri $Uri -UseDefaultCredentials
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
            }
            'WebSession'
            {
                $Uri = $WebSession.uri
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
                #Connection details here from websession, no creds needed as already there
            }
        }

        #Construct the uri and add it to paramaters block
        $TeamId = Get-TfsTeam -WebSession $Websession -Project $Project | Where-Object Name -eq "$Team" | Select-Object -ExpandProperty id
        $uri = "$uri/$($Project)/_apis/Dashboard/Groups/$($TeamId)"
        $Parameters.add('Uri',$uri)


        try
        {
            $jsondata = Invoke-restmethod @Parameters -erroraction Stop
        }
        catch
        {
            throw
        }

        Write-Output $jsondata.dashboardentries

    }
}

Function New-TfsTeamProjectDashboard
{
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Team,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Name,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials

    )
    Process
    {
        $headers = @{'Content-Type'='application/json'}
        $Parameters = @{}

        #Use Hashtable to create param block for invoke-restmethod and splat it together
        switch ($PsCmdlet.ParameterSetName) 
        {
            'SingleConnection'
            {
                $WebSession = Connect-TfsServer -Uri $uri -Username $Username -AccessToken $AccessToken
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)

            }
            'LocalConnection'
            {
                $WebSession = Connect-TfsServer -uri $Uri -UseDefaultCredentials
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
            }
            'WebSession'
            {
                $Uri = $WebSession.uri
                $Parameters.add('WebSession',$WebSession)
                $Parameters.add('Headers',$headers)
                #Connection details here from websession, no creds needed as already there
            }
        }

        $Dashboards = Get-TfsTeamProjectDashboard -WebSession $WebSession -Team $Team -Project $Project 
        if ($Dashboards | Where-Object {$_.Name -eq $Name})
        {
            Write-Error "Dashboard $Name already exists."
            break
        }
        else
        {
            $NextDashboardSlot = $Dashboards.Position[-1] + 1
        }

        #Construct the uri and add it to paramaters block
        $TeamId = Get-TfsTeam -WebSession $Websession -Project $Project | Where-Object Name -eq "$Team" | Select-Object -ExpandProperty id
        if ($uri -like '*.visualstudio.com*')
        {
            $uri = "$uri/$($Project)/$($TeamId)/_apis/Dashboard/Dashboards?api-version=3.1-preview.2"
        }
        else
        { 
            $uri = "$uri/$($Project)/_apis/Dashboard/groups/$($TeamId)/dashboards/?api-version=2.2-preview.1"
        }
        $Parameters.add('Uri',$uri)

        $body = @{if = $null; name = $Name; position = $NextDashboardSlot;widgets=$null; refreshInterval = $null; eTag = $null; _links = $null; url = $null} | ConvertTo-Json

        try
        {
            $JsonData = Invoke-RestMethod @Parameters -Method Post -Body $body -ErrorAction Stop
        
        }
        catch
        {
            $ErrorMessage = $_
            $ErrorMessage = ConvertFrom-Json -InputObject $ErrorMessage.ErrorDetails.Message
            if ($ErrorMessage.TypeKey -eq 'QueryItemNotFoundException')
            {
                $JsonData = $null
            }
            Else
            {
                Write-Error "Error was $_"
                $line = $_.InvocationInfo.ScriptLineNumber
                Write-Error "Error was in Line $line"
            }
        }

        Write-output $JsonData
    }
}