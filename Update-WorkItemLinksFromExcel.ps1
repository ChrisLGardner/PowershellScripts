[cmdletbinding()]
param 
(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter(Mandatory)]
    [string]$Sheet,

    [Parameter(Mandatory)]
    [string]$ParentIdColumn,

    [Parameter(Mandatory)]
    [string]$ChildIdColumn,

    [Parameter(ParameterSetName='SingleConnection',Mandatory)]
    [String]$uri,

    [Parameter(ParameterSetName='SingleConnection',Mandatory)]
    [string]$Username,

    [Parameter(ParameterSetName='SingleConnection',Mandatory)]
    [string]$AccessToken
)

function Add-TfsWorkItemParentChildLink
{
    <#  
        .SYNOPSIS
            This function will add a hyperlink to a work item.

        .DESCRIPTION
            This function will add a hyperlink to a work item.

            Additional comments can be added to the history and the hyperlink itself using the 
            parameters provided. This will auto-increment the revision number as needed based on the existing
            one on the work item.

        .PARAMETER WebSession
            Web session object for the target TFS server.
        
        .PARAMETER ParentId
        ID of the parent work item to add a link to.

        .PARAMETER ChildId
        ID of the child work item to link to.

        .PARAMETER Uri
            Uri of TFS serverm, including /DefaultCollection (or equivilent)

        .PARAMETER Username
            Username to connect to TFS with

        .PARAMETER AccessToken
            AccessToken for VSTS to connect with.

        .PARAMETER UseDefaultCredentails
            Switch to use the logged in users credentials for authenticating with TFS.

        .EXAMPLE 
            Add-TfsWorkItemHyperlink -WebSession $session -Id 123 -Hyperlink 'http://www.bbc.co.uk'

            This will add a hyperlink to bbc.co.uk to the work item using an existing web session for authenticating.

        .EXAMPLE
            Add-TfsWorkItemHyperlink -Id 123 -Hyperlink 'http://www.bbc.co.uk' -Uri 'https://product.visualstudio.com/DefaultCollection' -Username 'MainUser' -AccessToken (Get-Content c:\accesstoken.txt | Out-String)

            This will add a hyperlink to bbc.co.uk to the work item using the provided credentials.

    #>
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName='WebSession', Mandatory,ValueFromPipeline)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,

        [parameter(Mandatory)]
        [int]$ParentId,
        
        [parameter(Mandatory)]
        [int]$ChildId,

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
        $headers = @{'Content-Type'='application/json-patch+json';'accept'='api-version=2.2'}
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


        $ParentWorkItem = Get-TfsWorkItemDetail -WebSession $WebSession -ID $ParentId
        $ChildWorkItem = Get-TfsWorkItemDetail -WebSession $WebSession -ID $ChildId

        $JsonBody = @"
[
    {
        "op": "test",
        "path": "/rev",
        "value": $($ParentWorkItem.Rev)
    },
    {
    "op": "add",
    "path": "/relations/-",
    "value": {
      "rel": "System.LinkTypes.Hierarchy-forward",
      "url": "$($ChildWorkItem.url)",
      "attributes": {
        "comment": "Making a new link for the dependency"
      }
    }
  }
]
"@

        $Parameters.Add('Uri',$ParentWorkItem.url)

        try
        {
            $JsonOutput = Invoke-RestMethod -Method Patch -Body $JsonBody @Parameters -ErrorAction Stop
        }
        catch
        {
            Write-Error "Failed to update work item $ParentId."
        }

        Write-Output $JsonOutput
    }
}

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
            Write-Error "No work item with ID: $Id found in the target instance ($Uri)"
        }
        Write-Output $jsondata
    }
}

function Connect-TfsServer 
{
    <#
        .Synopsis
            Connects to a TFS server and returns a websession object for future connections.
        .DESCRIPTION
            This function allows two options, test the connection to a TFS server to ensure it's a valid endpoint
            or it can connect to a TFS server and authenticate with either the specified username and access token or with the
            default credentials used to log on to the computer.

            To test the connection to a TFS server the script attempts to Invoke-Webrequest to it and then checks the status code
            of the response. If the Invoke fails then the status code is set to 404 and will return '<URL> Not available' to the user.
            If the invoke succeeds and returns a non-20x status code then it's assumed it's unavailable. 

            When connecting to a TFS server there are two authentication methods available, for local servers you can use default credentials
            to connect using the domain credentials that you are logged in as (assuming you are on a domain computer and the TFS server authenticates
            with AD) or for remote servers you can use a personal access token, which can be generated from https://siteuri/_details/security/tokens. 
            Once a connection has been established the function will return a websession object which can be used with other functions in the module, 
            either by storing it in a variable or by piping this cmdlet to another.

        .PARAMETER Uri
            Uri of target TFS server

        .PARAMETER Username
            The username to connect to the remote server with

        .PARAMETER AccessToken
            Access token for the username connecting to the remote server

        .PARAMETER UseDefaultCredentials
            Switch for using local credentials when connecting to on-prem TFS server

        .PARAMETER TestConnection
            Switch to test if TFS server is available to connect

        .EXAMPLE
            Connect-TfsServer -Uri 'https://test.visualstudio.com/DefaultCollection' -TestConnection

            This will attempt to connect to the target TFS server to ensure it's a valid connection. If the server returns a status code other
            than a 20x then the script will report that it's unavailable.

        .EXAMPLE
            $WebSession = Connect-TfsServer -Uri 'https://test.visualstudio.com/DefaultCollection' -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt)

            This will attempt to connect to the remote TFS server hosted on VisualStudio.com using the credentials specified and return a 
            web session object for use with other functions 

        .EXAMPLE
            $WebSession = Connect-TfsServer -Uri 'https://tfs.domain.local/DefaultCollection' -UseDefaultCredentials

            This will attempt to connect to the local TFS server hosted on your domain using your current domain account credentials. It will
            return a web session object for use with the other functions.

        .EXAMPLE
            $OutputData = Connect-TfsServer -Uri 'https://test.visualstudio.com/DefaultCollection' -Username username@email.com -AccessToken (Get-Content C:\AccessToken.txt) | Get-TfsTeams -TeamProject 'TestProject'

            This will connect to the TFS server, create a web request object and then pipe it to the Get-TfsTeams and use it to complete the REST API check there.
    #>
    [cmdletbinding()]
    param(
        
        [parameter(Mandatory)]
        [String]$Uri,

        [Parameter(ParameterSetName='ConnectRemote',Mandatory)]
        [string]$Username,

        [Parameter(ParameterSetName='ConnectRemote',Mandatory)]
        [string]$AccessToken,

        [Parameter(ParameterSetName='ConnectLocal')]
        [switch]$UseDefaultCredentials,

        [Parameter(ParameterSetName='TestConnection')]
        [switch]$TestConnection
    
    )

    $Parameters = @{}

    switch ($PsCmdlet.ParameterSetName) 
    {
        'ConnectRemote'
        {
            $AuthToken = [Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f $Username,$AccessToken))
            $AuthToken = [Convert]::ToBase64String($AuthToken)
            $headers = @{Authorization=("Basic $AuthToken")}

            $Parameters.Add('Uri',$uri)
            $Parameters.Add('UseBasicParsing',$true)
            $Parameters.Add('SessionVariable','WebSession')
            $Parameters.Add('Headers',$headers)

            $Result = Invoke-WebRequest @Parameters
            $WebSession | Add-Member -MemberType NoteProperty -Name 'Uri' -Value $Uri
            Write-Output $WebSession

                
        }

        'ConnectLocal'
        {
            $Parameters.Add('Uri',$uri)
            $Parameters.Add('UseBasicParsing',$true)
            $Parameters.Add('SessionVariable','WebSession')
            $Parameters.Add('UseDefaultCredentials',$True)

            $Result = Invoke-WebRequest @Parameters
            $WebSession | Add-Member -MemberType NoteProperty -Name 'Uri' -Value $Uri
            Write-Output $WebSession
        
        }

        'TestConnection'
        {
            try 
            {
                $WebsiteStatus = (Invoke-WebRequest -UseBasicParsing -Uri $uri -ErrorAction Stop).StatusCode
            }
            catch
            {
                $WebsiteStatus = 404
            }
            if ($WebsiteStatus -like '20*')
            {
                Write-Output "$uri Available"
            }
            else
            {
                Write-Output "$uri Not available"
            }
        }
        
    } 

}
