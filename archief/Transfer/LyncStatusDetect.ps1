[CmdletBinding()]
param()

function OffHook-Action ($action)
	{
	if ($action -eq 'offhook')
		{
		InACall
		#Write-Host "Off-hook action executed"
		}
	else
		{
		HungUp
		#Write-Host "On-hook action executed"
		}
	}
    
function MusicSystemCommand($command) {

#Find Arduino COM Port
$PortName = (Get-WmiObject Win32_SerialPort | Where-Object { $_.Name -match "Arduino"}).DeviceID
if ( $PortName -eq $null ) { throw "Arduino Not Found"}
#Create SerialPort and Configure
$port = New-Object System.IO.Ports.SerialPort
$port.PortName = $PortName
$port.BaudRate = "9600"
$port.Parity = "None"
$port.DataBits = 8
$port.StopBits = 1
$port.ReadTimeout = 1000 #Milliseconds
$port.open() #open serial connection
Start-Sleep -Milliseconds 100 #wait 0.1 seconds
$port.Write($command) #write $byte parameter content to the serial connection
try    {

#Check for response
if (($response = $port.ReadLine()) -gt 0)
{ $response }
}
catch [TimeoutException] {
#Timeout
}
finally    {
$port.Close() #close serial connection
}
}

#Region Your dependent function(s) for off-hook actions

$global:GEWISCallStatus = 0
$global:GEWISRingingStatus = 0
$global:CleanUpInitiated = 0
$global:PhoneStatus = 0

function HungUp
	{
	#Write-Host "User hung up" -ForegroundColor Yellow -BackgroundColor Black
	$text = '0';
    $text > 'C:\inetpub\wwwroot\status.txt';
    '1' > 'C:\inetpub\wwwroot\change.txt';
    $global:GEWISCallStatus = 0
    $global:PhoneStatus = 0
    MusicSystemCommand("on")
	}
		
function IncomingCall ($caller)
	{
    #Write-Host "Phone is ringing, caller: $caller" -ForegroundColor Yellow -BackgroundColor Black
    $text = '1';
    $text > 'C:\inetpub\wwwroot\status.txt';
    $caller > 'C:\inetpub\wwwroot\caller.txt';
    '1' > 'C:\inetpub\wwwroot\change.txt';
    '1' > 'C:\inetpub\wwwroot\ringing.txt';
    $global:GEWISRingingStatus = 1
    MusicSystemCommand("off")
	}
    
function InACall
	{
    #Write-Host "In a call"
    $text = '1';
    $text > 'C:\inetpub\wwwroot\status.txt';
    '1' > 'C:\inetpub\wwwroot\change.txt';
    $global:GEWISCallStatus = 1
    MusicSystemCommand("off")
	}
	
$timer2 = New-Object Timers.Timer
$timer3 = New-Object Timers.Timer

function CheckStatus
	{
	if (($global:GEWISRingingStatus -eq 1) -and ($global:GEWISCallStatus -eq 1))
		{
		$global:PhoneStatus = 1
		#Write-Host "Case 1: External call picked up"
		if($global:CleanUpInitiated -eq 0)
			{
			$timer2.Interval = 30000
			$timer2.AutoReset = $false
			$timer2.Enabled = $true

			## $args[0] Timer object
			## $args[1] Elapsed event properties
			Register-ObjectEvent -InputObject $timer2 -EventName Elapsed -SourceIdentifier CleanUpAfterRing  -Action {CleanUpAfterRing}
			
			$global:CleanUpInitiated = 1
			#Write-Host "Cleanup timer initiated for case 1" -ForegroundColor Green
			}
    	}
   if (($global:GEWISRingingStatus -eq 1) -and ($global:GEWISCallStatus -eq 0))
		{
		$global:PhoneStatus = 1
		#Write-Host "Case 2: Phone is ringing"
		if($global:CleanUpInitiated -eq 0)
			{
			$timer3.Interval = 30000
			$timer3.AutoReset = $false
			$timer3.Enabled = $true

			## $args[0] Timer object
			## $args[1] Elapsed event properties
			Register-ObjectEvent -InputObject $timer3 -EventName Elapsed -SourceIdentifier CleanUpAfterRing  -Action {CleanUpAfterRing}
			
			$global:CleanUpInitiated = 1
			#Write-Host "Cleanup timer initiated for case 2" -ForegroundColor Green
			}
    	}
    if (($global:GEWISRingingStatus -eq 0) -and ($global:GEWISCallStatus -eq 1))
		{
		$global:PhoneStatus = 1
		#Write-Host "Case 3: Self-initiated call in progress" 
    	}
    if ($global:PhoneStatus -eq 1)
    	{
    	#Write-Host "GEWIS music system has been muted" -ForegroundColor Red
    	}
	}
	
function CleanUpAfterRing
	{
		if($global:GEWISCallStatus -eq 0)
		{
		#Write-Host "Resetting ring status" -ForegroundColor Red
		$global:GEWISRingingStatus = 0
		$global:CleanUpInitiated = 0
		$global:PhoneStatus = 0
		$text = '0';
    	$text > 'C:\inetpub\wwwroot\status.txt';
    	'1' > 'C:\inetpub\wwwroot\change.txt';
        MusicSystemCommand("on")
    	}
	}
	
$timer = New-Object Timers.Timer

$timer.Interval = 2000
$timer.AutoReset = $true
$timer.Enabled = $true

## $args[0] Timer object
## $args[1] Elapsed event properties
Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier CheckStatus  -Action {CheckStatus}
    
#EndRegion	

#Region Core functions
function Write-VerboseEvent ($text)
	{
	if ($verboseEvent)
		{
		Write-Host "VERBOSE: $((Get-Date).ToLongTimeString()) - $text" -ForegroundColor Yellow -BackgroundColor Black
		}
	}

function Connect-LyncClient
	{
	$i = 1
	do
		{
		#Attach to Lync process, if running
		#Choose 'lync' line for 2013, 'communicator' line for 2010
		$lyncProcess = [System.Diagnostics.Process]::GetProcessesByName('lync')
		#$lyncProcess = [System.Diagnostics.Process]::GetProcessesByName('communicator')
		
		if ($lyncProcess.Length -eq 0) #Process is not running
			{
			if ($i -eq 1) #Report status only on first attempt
				{
				Write-Host "$((Get-Date).ToLongTimeString()) - Waiting for Lync process to start (15-second intervals)..." -ForegroundColor Yellow
				}
			$i++
			Start-Sleep 15
			}
		}
	until ($lyncProcess.Length -eq 1)

	#Register for when Lync process exits
	Register-ObjectEvent -InputObject $lyncProcess[0] -EventName "Exited" -SourceIdentifier "LyncProcessHandler" -Action {LyncProcess-Handler} | Out-Null

	#Wait for client object initialization to complete
	do
		{
		$global:client = [Microsoft.Lync.Model.LyncClient]::GetClient()
		}
	while (-not $client -or $client.State -eq [Microsoft.Lync.Model.ClientState]::Invalid)
	}

function Register-ContactChange
	{
	#Create self object as a contact
	$global:selfContact = $client.Self.Contact
	
	#Register for contact changes
	Register-ObjectEvent -InputObject $selfContact -EventName "ContactInformationChanged" -SourceIdentifier "OffHookHandler" -Action {Offhook-Handler $event} | Out-Null
	
    $conversationMgr = $client.ConversationManager
    
    #Register for incoming call
    Register-ObjectEvent -InputObject $conversationMgr -EventName "ConversationAdded" -SourceIdentifier "NewIncomingConversation" -action { 
    $global:myEvent = $Event 
    $caller = $Event.Sender.Conversations[0].Participants | Select -ExpandProperty Contact | Select -ExpandProperty Uri
    IncomingCall($caller)
    }
        
    #Get initial off-hook status and set state variable
	$activity = $selfContact.GetContactInformation([Microsoft.Lync.Model.ContactInformationType]::Activity)
	if ($activity -eq 'In a call' -or $activity -eq 'In a conference call')
		{
		$global:offhook = $true
		}
	else
		{
		$global:offhook = $false
		}
	}

function Register-ClientStateChange
	{
	#Register for sign-in changes
	Register-ObjectEvent -InputObject $client -EventName "StateChanged" -SourceIdentifier "ClientStateHandler" -Action {ClientState-Handler $event} | Out-Null
	}

function Offhook-Handler ($event)
	{
	#Act if what has changed is activity
	if ($event.SourceEventArgs.ChangedContactInformation -contains 'Activity')
		{
		$newActivity = $selfContact.GetContactInformation([Microsoft.Lync.Model.ContactInformationType]::Activity)
		#Act only on true activity change
		if ($newActivity -ne $currentActivity)
			{
			#Act if off- or on-hook
			if ($newActivity -eq 'In a call' -or $newActivity -eq 'In a conference call')
				{
				if ($offhook -eq $false) #Only run off-hook action if not already on a call
					{
					Write-VerboseEvent $newActivity
					OffHook-Action 'offhook'
					$global:offhook = $true #Stateful tracking of status in successive changes
					}
				}
			else
				{
				if ($offhook)
					{
					OffHook-Action 'onhook'
					Write-VerboseEvent "No longer on the phone ($newActivity)"
					$global:offhook = $false
					}
				else
					{
					Write-VerboseEvent "Non-phone activity change: $newActivity"
					}
				}
			#Global variable provides stateful tracking of activity change
			$global:currentActivity = $newActivity
			}
		}
	}
	
function ClientState-Handler ($event)
	{
	#Get current client state
	$newState = $event.SourceEventArgs.NewState
	if ($newState -eq 'SignedIn')
		{
		Register-ContactChange
		Write-VerboseEvent "Activity changes now being monitored."
		}
	elseif ($newState -eq 'SignedOut') 
		{
		$subscriptionSource = Get-EventSubscriber | Select-Object -ExpandProperty SourceIdentifier
		if ($subscriptionSource -contains "OffHookHandler")
			{
			#If subscription currently is registered, remove it so it can
			#be successfully created again when signed in
			Unregister-Event OffHookHandler
			Write-VerboseEvent "Activity changes will be monitored when the client signs in."
			}
		}
	}

function LyncProcess-Handler
	{
	Write-Host "$((Get-Date).ToLongTimeString()) - Lync client has shut down." -ForegroundColor Yellow
	#Client object is invalid if Lync process stops
	Stop-Monitoring
	
	#Restart connection and registration steps
	Connect-LyncClient
	Initialize-Registration
	}

function Initialize-Registration
	{
	#Register for contact changes if client is already signed in
	if ($client.State -eq [Microsoft.Lync.Model.ClientState]::SignedIn)
		{
		Register-ContactChange
		Write-Host "$((Get-Date).ToLongTimeString()) - Lync phone status for GEWIS is now being monitored. Do not close this PowerShell window." -ForegroundColor Green
		Register-ClientStateChange
		}
	#Register for client changes, which will handle contact change registration
	else
		{
		Register-ClientStateChange
		Write-VerboseEvent "Activity changes will be monitored when the client has signed in."
		}
	}

function Stop-Monitoring
	{
	Unregister-Event OffHookHandler -ErrorAction SilentlyContinue
	Unregister-Event ClientStateHandler -ErrorAction SilentlyContinue
	Unregister-Event LyncProcessHandler -ErrorAction SilentlyContinue
	Write-VerboseEvent "Events unregistered"
	$global:client = $null
	$global:lyncProcess = $null
	$global:currentActivity = $null
	}

#EndRegion

#Region Script body

#Check for dot sourcing
if ($MyInvocation.InvocationName -ne '.')
	{
	Write-Error "Script was not dot-sourced.  This script is designed to be executed by dot sourcing it: . <pathtoscript>" -Category InvalidOperation
	break
	}

#Check for SDK installation
$apiPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Lync.Model\4.0.0.0__31bf3856ad364e35\Microsoft.Lync.Model.dll"
if (Test-Path $apiPath)
	{
	Add-Type -Path $apiPath
	#Connect to local Lync client
	Connect-LyncClient
	}
else
	{
	Write-Error "This script requires the Lync SDK runtime library." -Category NotInstalled
	break
	}

#Check for Verbose parameter for event functions
if ($MyInvocation.BoundParameters['verbose'])
	{
	$global:verboseEvent = $true
	}
else
	{
	$global:verboseEvent = $false
	}
	
#Start event registration
Initialize-Registration

#EndRegion