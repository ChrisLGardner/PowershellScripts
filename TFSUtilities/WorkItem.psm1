function Get-TfsWorkItemDetail
{
    <#
        .SYNOPSIS
            This function gets the details of the specified work item

        .DESCRIPTION
            This function gets the details of the specified work item from either the target
            server, using the specified credentials, or to the specified WebSession.

            The function will return the raw JSON output as a PS object, if an invalid ID is
            provided then an error will be returned.

        .PARAMETER WebSession
            Websession with connection details and credentials generated by Connect-TfsServer function

        .PARAMETER ID
            ID of work item to look up

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE
            Get-TfsWorkItemDetail -Uri https://test.visualstudio.com/DefaultCollection  -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) -id 1

            This will get the details of the work item with ID 1, connecting to the target TFS server using the specified username 
            and password

        .EXAMPLE
            Get-TfsWorkItemDetail -WebSession $Session -id 1
            
            This will get the details of the work item with ID 1, connecting to the TFS server specified in the web session and using
            the credentials stored there.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$ID,

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
        write-verbose "Getting details for WI $id via $($WebSession.Uri) "

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

        $uri = "$Uri/_apis/wit/workitems/$id"
        $Parameters.add('Uri', $Uri)
        try
        {
            $jsondata = Invoke-RestMethod @Parameters -ErrorAction Stop
        }
        catch
        {
            Throw

        }
        Write-Output $jsondata
    }
}


function Get-TfsWorkItemInIterationWithNoTask
{
    <#
        .SYNOPSIS
            This function gets any work items in the specified Iteration with no tasks linked

        .DESCRIPTION
            This function gets any work items which have no tasks linked to them in the specified iteration.
            It will query TFS for all the tasks in the specified Iteration path with the specified states, 
            and then iterate over them to find which ones are root items, which are parents and which are children.
            Then it can compare the list of root items with the list of parent items and find any which are root items
            but not parents. It will then use the Get-WorkItemDetails function to get all the details of these items
            and return that data to the pipeline.

            The function accepts input from the pipeline in the form of a WebSession object, such as generated by the Connect-TfsServer
            function.

        .PARAMETER WebSession
            Websession with connection details and credentials generated by Connect-TfsServer function

        .PARAMETER IterationPath
            The exact path of the iteration to look up, such as 'test\Sprint 1'

        .PARAMETER States
            String of states to query, for multiple states use double quotes around a comma seperated list of single quoted strings.
            Accepted states are: Approved, Committed, Done, In Test, New, Removed

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE
            Get-TfsWorkItemInIterationWithNoTask -Uri https://test.visualstudio.com/DefaultCollection -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) -IterationPath 'Test\Sprint 1' -States 'New'

            This will get all the work items with no tasks in Sprint 1 of the Test project that are in the New state, using the specified credentials and Uri.

        .EXAMPLE
            Get-TfsWorkItemInIterationWithNoTask -WebSession $Session -IterationPath 'Test\Sprint 4' -States "'New','Approved'"
            
            This will get all the work items with no tasks in Sprint 4 of the Test project with the New or Approved state, using the WebSession object for the Uri and credentials.

        .EXAMPLE
            Connect-TfsServer -Uri "https://test.visualstudio.com/DefaultCollection -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) |  Get-TfsWorkItemInIterationWithNoTask -IterationPath 'Test\Sprint 4' -States "'New','Approved'"
            
            This will connect to the specified TFS server and then pass the WebSession object into the pipeline and get all the work items with no tasks in Sprint 4 of the Test project with the New or Approved state.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$IterationPath,

        [Parameter(Mandatory)]
        [String]$States,

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

        write-verbose "Getting Backlog Items under $iterationpath via $uri that have no child tasks" 

        $queryuri = "$($uri)/_apis/wit/wiql?api-version=1.0"
        $wiq = "SELECT [System.Id], [System.Links.LinkType], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItemLinks WHERE (  [Source].[System.State] IN ($states)  AND  [Source].[System.IterationPath] UNDER '$iterationpath') And ([System.Links.LinkType] <> '') And ([Target].[System.WorkItemType] = 'Task') ORDER BY [System.Id] mode(MayContain)"
        $data = @{query = $wiq } | ConvertTo-Json

        $Parameters.add('Uri', $queryUri)

        try
        {
            $jsondata = Invoke-RestMethod @parameters -Method Post -Body $data   #$wc.UploadString($uri,'POST', $data) | ConvertFrom-Json 
        }
        catch
        {
            Throw
        }
    
        # work out which root items have no child tasks
        # might be a better way to do this
        $rootItems = @()
        $childItems = @()
        $parentItems = @()
    
        foreach($wi in $jsondata.workItemRelations)
        {
            if ($wi.rel -eq $null)
            {
                $rootItems += $wi.target.id
            } else 
            {
                $childItems += $wi.target.id
                $parentItems += $wi.source.id
            }
        }

        $ids = (Compare-Object -ReferenceObject ($rootItems |  Sort-Object) -DifferenceObject ($parentItems | Select-Object -uniq |  Sort-Object)).InputObject
        $retItems = @()

        foreach ($id in $ids)
        {
            if ($WebSession)
            {
                $item = Get-TfsWorkItemDetail -WebSession $WebSession -id $id
            }
            else
            {
                $item = Get-TfsWorkItemDetail -uri $uri -id $id -username $username -password $password
            }
            $retItems += $item | Select-Object id, @{ Name = 'WIT' ;Expression ={$_.fields.'System.WorkItemType'}} , @{ Name = 'Title' ;Expression ={$_.fields.'System.Title'}}

        }
    
        Write-Output $retItems
    }
}


Function Get-TfsWorkItemInIteration
{
    <#
        .SYNOPSIS
            This function gets any work items in the specified Iteration

        .DESCRIPTION
            This function gets any work items in the specified iteration.

            The function accepts input from the pipeline in the form of a WebSession object, such as generated by the Connect-TfsServer
            function.

        .PARAMETER WebSession
            Websession with connection details and credentials generated by Connect-TfsServer function

        .PARAMETER IterationPath
            The exact path of the iteration to look up, such as 'test\Sprint 1'

        .PARAMETER IdOnly
            Switch that will cause function to only return the Ids of the Work Items

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE
            Get-TfsWorkItemInIteration -Uri https://test.visualstudio.com/DefaultCollection -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) -IterationPath 'Test\Sprint 1'

            This will get all the work items in Sprint 1 of the Test project that are in the New state, using the specified credentials and Uri.

        .EXAMPLE
            Get-TfsWorkItemInIteration -WebSession $Session -IterationPath 'Test\Sprint 4'
            
            This will get all the work items in Sprint 4 of the Test project with the New or Approved state, using the WebSession object for the Uri and credentials.

        .EXAMPLE
            Connect-TfsServer -Uri "https://test.visualstudio.com/DefaultCollection  -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) |  Get-TfsWorkItemInIterationWithNoTask -IterationPath 'Test\Sprint 4'
            
            This will connect to the specified TFS server and then pass the WebSession object into the pipeline and get all the work items in Sprint 4 of the Test project.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$IterationPath,

        [switch]$IdOnly,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [Parameter(ParameterSetName='LocalConnection',Mandatory)]
        [String]$uri,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='SingleConnection',Mandatory)]
        [string]$AccessToken,

        [parameter(ParameterSetName='LocalConnection',Mandatory)]
        [switch]$UseDefaultCredentials,

        [switch]$IncludeClosed

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
        
        write-verbose "Getting Backlog Items under $iterationpath via $uri that have no child tasks" 

        $queryuri = "$($uri)/_apis/wit/wiql?api-version=1.0"
        if ($IncludeClosed)
        {
            $wiq = "SELECT [System.Id], [System.Links.LinkType], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItemLinks WHERE ([Source].[System.IterationPath] UNDER '$iterationpath') And ([System.Links.LinkType] <> '') And ([Target].[System.WorkItemType] IN GROUP 'Microsoft.RequirementCategory') ORDER BY [System.Id] mode(MayContain)"
        }
        else
        {
            $wiq = "SELECT [System.Id], [System.Links.LinkType], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItemLinks WHERE ([Source].[System.IterationPath] UNDER '$iterationpath') And ([Source].[System.State] <> 'Done') And ([Source].[System.State] <> 'Removed') And ([System.Links.LinkType] <> '') And ([Target].[System.WorkItemType] IN GROUP 'Microsoft.RequirementCategory') ORDER BY [System.Id] mode(MayContain)"
        }
        $data = @{query = $wiq } | ConvertTo-Json

        $Parameters.add('Uri', $queryUri)

        Try
        {
            $jsondata = Invoke-RestMethod  @parameters -Method Post -Body $data -ErrorAction Stop
        }
        catch
        {
            Throw
        }

        if ($IdOnly)
        {
            Write-Output $jsondata.workItemRelations.target | Select-Object -ExpandProperty Id -Unique
        }
        else
        {
            Write-Output $jsondata
        }
    }
}


function Add-TfsTaskToWorkItem
{
    <#
        .SYNOPSIS
            This function will add the specified task or tasks to a specified work item

        .DESCRIPTION
            This function will add the task or tasks to the work item specified. 

            The function will take either a websession object or a uri and
            credentials. The web session can be piped to the fuction from the
            Connect-TfsServer function.

        .PARAMETER WebSession
            Websession with connection details and credentials generated by Connect-TfsServer function

        .PARAMETER ID
            The ID of the work item to add the task(s) to

        .PARAMETER Task
            The name of the task or tasks to add to the work item

        .PARAMETER WorkRemaining
            The work remaining for the work item to be added, defaults to 0

        .PARAMETER IterationPath
            The iteration path to add the tasks to

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE
            Add-TfsTasskToWorkItem -WebSession $Session -Task 'Code Review' -id 3

            This will add a task named 'Code Review' to the work item with an id of 3.

        .EXAMPLE
            Add-TfsTasskToWorkItem -Uri 'https://test.visualstudio.com/DefaultCollection' -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) -Task 'Code Review' -id 3 -WorkRemaining 4

            This will add a task named 'Code Review' to the work item with an id of 3 and set the work remaining to 4.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Id,

        [Parameter(Mandatory)]
        [String]$IterationPath,

        [Parameter(Mandatory)]
        [String[]]$Task,

        [int]$WorkRemaining = 0,

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
        
        $uri = "$Uri/$($IterationPath.split('\')[0])/_apis/wit/workitems/`$Task?api-version=1.0"
        $Parameters.add('Uri', $Uri)
        $jsondata = @()

        foreach ($TaskToAdd in $task)
        {
            $data = @(@{op = 'add'; path = '/fields/System.Title'; value = "$TaskToAdd" } ; `
                      @{op = 'add'; path = '/fields/System.Description'; value = "$TaskToAdd" };  `
                      @{op = 'add'; path = '/fields/Microsoft.VSTS.Scheduling.RemainingWork'; value = "$WorkRemaining" }  ;  `
                      @{op = 'add'; path = '/fields/System.IterationPath'; value = "$IterationPath" }  ;  `
                      @{op = 'add'; path = '/relations/-'; value = @{ 'rel' = 'System.LinkTypes.Hierarchy-Reverse' ; 'url' = "$($WebSession.Uri)/_apis/wit/workItems/$id"} }   ) | ConvertTo-Json
            
            try
            {
                $jsondata += Invoke-RestMethod @parameters -Method Patch -Body $data -ErrorAction Stop
            }
            catch
            {
                Throw
            }
        }

        Write-Output $jsondata
    }
}


function New-TfsWorkItemQuery
{
    <#  
        .SYNOPSIS
            This function will add a work item query to a folder.

        .DESCRIPTION
            This function will add a work item query to a folder when passed a wiql string.

        .PARAMETER WebSession
            Web session object for the target TFS server.

        .PARAMETER Project
            The name of the project under which the team can be found

        .PARAMETER Folder
            The name of the folder to store the query.

        .PARAMETER Name
            The name of the query to create.

        .PARAMETER Wiql
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
            $WiqlString = "select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] from WorkItems where [System.TeamProject] = @project and [System.WorkItemType] <> '' and [System.State] = 'In Test' order by [System.WorkItemType] desc"
            
            New-TfsWorkItemQuery -WebSession $session -Project 'Super Product' -Folder 'Shared Queries' -Name 'In Test' -Wiql $Wiql 

            This will add a new query to the Shared Queries folder with the specified wiql using the specified web session.

        .EXAMPLE
            $WiqlString = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State],[System.Tags] FROM WorkItemLinks WHERE ([Source].[System.TeamProject] = @project AND ( [Source].[System.WorkItemType] = 'Product Backlog Item' OR [Source].[System.WorkItemType] = 'Bug' ) AND [Source].[System.State] <> 'Done' AND [Source].[System.State] <> 'Removed' AND [Source].[System.IterationPath] = @currentIteration) AND ([Target].[System.TeamProject] = @project AND [Target].[System.WorkItemType] = 'Task' AND [Target].[System.AssignedTo] = @me AND [Target].[System.State] <> 'Done' AND [Target].[System.State] <> 'Removed' AND [Target].[System.IterationPath] = @currentIteration) mode(MustContain)"
            
            New-TfsWorkItemQuery -Project 'Super Product' -Folder 'Shared Queries' -Name 'In Test' -Wiql $WiqlString -Uri 'https://product.visualstudio.com/DefaultCollection' -Username 'MainUser' -AccessToken (Get-Content c:\accesstoken.txt | Out-String)

            This will add a new query to the Shared Queries folder with the specified wiql on the target VSTS account using the provided creds.

    #>
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Folder,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Wiql,

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
        $headers = @{'Content-Type'='application/json';'accept'='api-version=2.2;application/json'}
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

        #Make the variables web safe
        $NameParsed = $Name -replace ' ','%20'
        $FolderParsed = $Folder -replace ' ','%20'

        #Get queries to check if one already exists in that location
        if (Get-TfsWorkItemQuery -WebSession $WebSession -Project $Project -Folder $FolderParsed -Name $NameParsed)
        {
            Write-Error "$name already exists in the location specified. Please try again with a different name."
            break
        }
        else
        {
            Write-verbose -Verbose "Creating query: $name"
            try 
            {
                $uri = "$uri/$project/_apis/wit/queries/$($Folder)?api-version=2.2"
                $Parameters.Add('Uri',$uri)
                $Body = @{
                    name = $Name
                    wiql = $Wiql 
                } | ConvertTo-Json
               
               $JsonData = Invoke-RestMethod @Parameters -Method Post -Body $Body
            }
            catch
            {
                $ErrorMessage = $_
                $ErrorMessage = ConvertFrom-Json -InputObject $ErrorMessage.ErrorDetails.Message
                if ($ErrorMessage.TypeKey -eq 'LegacyQueryItemException')
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

        }
        
        Write-Output $JsonData

    }
}


function Get-TfsWorkItemQuery
{
    <#  
        .SYNOPSIS
            This function will get a work item query from a folder on a project.

        .DESCRIPTION
            This function will get a work item query from a folder on a project.

        .PARAMETER WebSession
            Web session object for the target TFS server.

        .PARAMETER Project
            The name of the project under which the team can be found

        .PARAMETER Folder
            The name of the folder to store the query.

        .PARAMETER Name
            The name of the query to create.

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            Username to connect to TFS with

        .PARAMETER AccessToken
            AccessToken for VSTS to connect with.

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE 
            New-TfsWorkItemQuery -WebSession $session -Project 'Super Product' -Folder 'Shared Queries' -Name 'In Test' -Wiql $Wiql 

            This will add a new query to the Shared Queries folder with the specified wiql using the specified web session.

        .EXAMPLE
            New-TfsTeamProjectDashboardWorkItemQuery -Project 'Super Product' -Folder 'Shared Queries' -Name 'In Test' -Wiql $WiqlString -Uri 'https://product.visualstudio.com/DefaultCollection' -Username 'MainUser' -AccessToken (Get-Content c:\accesstoken.txt | Out-String)

            This will add a new query to the Shared Queries folder with the specified wiql on the target VSTS account using the provided creds.

    #>
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [Parameter(Mandatory)]
        [String]$Project,

        [Parameter(Mandatory)]
        [String]$Folder,

        [Parameter(Mandatory)]
        [string]$Name,

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
        $headers = @{'Content-Type'='application/json';'accept'='api-version=2.2;application/json'}
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

        #Make the variables web safe
        $Name = $Name -replace ' ','%20'
        $Folder = $Folder -replace ' ','%20'

        #Get queries to check if one already exists in that location
        $Uri = "$($WebSession.uri)/$Project/_apis/wit/queries/$Folder/$($Name)?api-version=2.2"
        $Parameters.Add('uri',$uri)

        try
        {
            $JsonData = Invoke-RestMethod @Parameters -Method GET -ErrorAction Stop
        
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