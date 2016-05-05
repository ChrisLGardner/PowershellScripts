<#
.Synopsis
   Checks the event log of the domain controllers and report any successful login attempts
for a particular user and which computers they logged in to. Needs to run as Administrator to access the log.
.DESCRIPTION
   Checks the event log of the domain controllers and report any successful login attempts
for a particular user and which computers they logged in to.Uses Get-WinEvent and filtering on ID 4624
between the dates provided. Then collates the Eventlogs into XML, which is searched for the username and
that is stored along with the IP and computer name (from DNS lookup) and output to the user.

This script can take quite a while to run due to the number of event log entries that are logged on each server,
testing was showing times of 10+ minutes to query each server when looking at the results for just one day. 
Further investigation required to attempt to streamline the process to reduce this time.

.EXAMPLE
   Get-LoginSuccess -Username "test.user" -StartDate "10/01/2015" -EndDate "11/01/2015"

   This will get all login successes by user test.user on the dates 10/01/2015 and 11/01/2015. If the second date is in the future then
   it will only get them from any days that have actually happened.

.EXAMPLE
   Get-LoginSuccess -Username "test.user"

   This will get all login successes by user test.user for the current day.

.EXAMPLE
   Get-LoginSuccess -ComputerName "test-computer"

   This will get all login successes by each user who logged into test-computer for the current day.

.EXAMPLE
   Get-LoginSuccess -ComputerName "test-computer" -StartDate "10/01/2015"

   This will get all login successes by each user who logged into test-computer from the start date till today.

.EXAMPLE
   Get-LoginSuccess -ComputerName "test-computer" -StartDate "10/01/2015" -EndDate "15/01/2015"

   This will get all login successes by each user who logged into test-computer from the start date till the specified end date.

.EXAMPLE
   Get-LoginSuccess -ComputerName "test-computer" -Username "test.user" -StartDate "10/01/2015" -EndDate "15/01/2015"

   This will get all login successes by specified userlogging into test.computer from the start date till the specified end date.

#>
[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$True,HelpMessage='Computer to check',ValueFromPipelineByPropertyName,
        position=0,ParameterSetName="Computer")]
    [Alias('Computer Name','Computer','Device','Device Name')]
    [String]$ComputerName = "test-comptuer",
	[Parameter(Mandatory=$false,HelpMessage="Username to search for",
        ValueFromPipelineByPropertyName,Position=0,ParameterSetName='IndividualUser')]
    [Parameter(HelpMessage="Username to search for",
        ValueFromPipelineByPropertyName,Position=1,ParameterSetName='Computer')]
    [alias('user','name')]
    [string]$Username,
	[string]$StartDate = $(Get-Date -format dd/MM/yyyy),
	[string]$EndDate
)
Process 
{
    if ($EndDate -eq '' -or $EndDate -eq $StartDate) 
    {
	    $EndDate = Get-Date -date ((Get-date $StartDate).AddDays(1)) -format dd/MM/yyyy
    }
    else
    {
        $EndDate = Get-Date -Date $EndDate -Format dd/MM/yyyy
    }
    
    if ($EndDate -lt $StartDate) 
    {
	    Write-error "End date ($EndDate) must be after start date ($StartDate). Please try again with the correct dates."
        Exit
    }
    
    $EventOutput =@()

    Switch ($PSCmdlet.ParameterSetName)
    {
        "IndividualUser" {
            $Controllers = Get-ADDomainController -filter *
            Foreach ($Controller in $Controllers)
            {
                Write-Verbose "Getting event data for $StartDate until $EndDate from $($Controller.Name)"
                $EventLog = Get-WinEvent -FilterHashTable @{logname='security';id=4624;StartTime=$StartDate;EndTime=$EndDate} -ComputerName "$($Controller.Name)"
                Foreach ($Event in $EventLog)
                {
                    $SingleEventOutput = New-Object -TypeName PSObject -Property @{'Username'='';'Computer'='';'Date'="$($Event.TimeCreated)"}
                    write-Verbose "Converting to XML"
                    $EventLogXML = [XML]$Event.ToXML()
                    Write-Verbose "Parsing XML and finding required entries"
                    $ValidEntry = 0
	                foreach ($Property in $EventLogXML.Event.EventData.Data) 
                    {
		                if ($Property.Name -eq "TargetUserName" -and $Property.'#text' -eq $Username) 
                        {
			                $SingleEventOutput.Username = $Property.'#text'
                            $ValidEntry++
		                }
		                elseif ($Property.Name -eq "IpAddress" -and $Property.'#text' -notmatch "::ffff:<serverIPRange>" -and $Property.'#text' -notmatch "<OtherIPRangeToIgnore>" -and $ValidEntry -eq 1) 
                        {
                            Write-Verbose "Getting computer name from IP Address - ($($Property.'#text'))"
                            Write-Debug "Getting computer name from IP Address - ($($Property.'#text'))"
			                $PCName = nslookup ($Property.'#text')
			                $SingleEventOutput.Computer = (($PCName[3]).substring(9))
                            $ValidEntry++
		                }
	                }
                    if ($ValidEntry -eq 2) {
                        $EventOutput += $SingleEventOutput
                    }
                }
            }
        }

        "Computer" {
            Write-Verbose "Getting event data for $StartDate until $EndDate"
            $EventLog = Get-WinEvent -FilterHashTable @{logname='security';id=4624;StartTime=$StartDate;EndTime=$EndDate} -ComputerName $ComputerName
            Foreach ($Event in $EventLog)
            {
                $SingleEventOutput = New-Object -TypeName PSObject -Property @{'Username'='';'Computer'="$Computername";'Date'="$($Event.TimeCreated)"}
                write-Verbose "Converting to XML"
                $EventLogXML = [XML]$Event.ToXML()
                Write-Verbose "Parsing XML and finding required entries"
                $ValidEntry = 0
	            foreach ($Property in $EventLogXML.Event.EventData.Data) 
                {
		            if ($Username -ne '')
                    {
                        if ($Property.Name -eq "TargetUserName" -and $Property.'#text' -eq $Username)
                        {
			                $SingleEventOutput.Username = $Property.'#text'
                            $ValidEntry++
		                }
                    }
                    else
                    {
                        if ($Property.Name -eq "TargetUserName" -and ($Property.'#text').Trim() -notmatch "$ComputerName" -and $Property.'#text' -ne "SYSTEM" -and $Property.'#text' -ne $env:USERNAME) 
                        {
			                $SingleEventOutput.Username = ($Property.'#text').Trim()
                            $ValidEntry++
		                }
                    }
		            
	            }
                if ($ValidEntry -eq 1) {
                    $EventOutput += $SingleEventOutput
                }
            }
        }
    }

    Write-Output $EventOutput
}