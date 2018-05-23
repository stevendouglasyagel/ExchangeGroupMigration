#################################################################################
#
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages.
#
#################################################################################
#
# Created by Matbyrd@microsoft.com as a PowerShell Script Start-RobustCloudCommand.ps1
# and last updated 5/10/2016. See https://gallery.technet.microsoft.com/scriptcenter/Start-RobustCloudCommand-69fb349e
#
# Adapted and improved by nilesh.ghodekar@microsoft.com as a PowerShell module RobustCloudCommand.psm1
# Now also uses only the new EXO PowerShell module that has MFA support
#################################################################################

# Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	
	# Get the current date
	[string]$date = Get-Date -Format G
		
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $Global:__LogFile__ -Append
	
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Debug
	}
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)

	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		
		# Sleep 1 second
		start-sleep 1
	}
	
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session
Function New-CleanO365Session {
	
	# Clear out all errors
	$Error.Clear()
	$Global:Error.Clear()
	
	# Create the session
	Write-Log "Creating new PS Session. User: $UserName"
	
	try
	{
		# Always use $UserPrincipalName and disregard any $Credential object passed to allow for implicit remoting next time
		# even though it means an addtional prompt for the first time
		$currentErrorActionPreference = $Global:ErrorActionPreference

		$Global:ErrorActionPreference = 'Stop'
		Connect-EXOPSSession -UserPrincipalName $UserName

		$Global:ErrorActionPreference = $currentErrorActionPreference
	}
	catch
	{
		$Global:ErrorActionPreference = $currentErrorActionPreference
		Write-Log "[ERROR] - Exception caught while invoking command Connect-EXOPSSession. Error: $_"
		
		if ($_.ToString() -match "500 - Internal server error")
		{
			# If we are not aborting then sleep in the hope that the issue is transient
			$sleep =  300 * ($ErrorCount + 1) + (Get-Random -Maximum 180)
			Write-Log "Sleeping $sleep seconds so that issue can potentially be resolved"
			Start-SleepWithProgress -sleeptime $sleep
		}
	}
	
	# Test for errors. For some reason the exception increases the error count in the $Global:Error only. So check $Global:Error.Count as well
	if ($Error.Count -gt 0 -or $Global:Error.Count -gt 0) 
	{	
		Write-Log "[ERROR] - Error while setting up session"
		if ($Error.Count  -gt 0)
		{
			Write-Log $Error
		}
		else
		{
			Write-Log $Global:Error
		}

		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		
		# if we have failed to setup the session > 5 times then we need to abort because we are in a failure state
		if ($ErrorCount -gt 5){
		
			Write-Log "[ERROR] - Failed to setup session after multiple tries"
			Write-Log "[ERROR] - Aborting Script"
			exit
		
		}
		
		# If we are not aborting then sleep 60s in the hope that the issue is transient
		Write-Log "Sleeping 60s so that issue can potentially be resolved"
		Start-SleepWithProgress -sleeptime 60
		
		# Attempt to set up the sesion again
		New-CleanO365Session
	}
	
	# If the session setup worked then we need to set $errorcount to 0
	else {
		$ErrorCount = 0
	}
	
	# Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy
# Goes ahead and resets it every $ResetSeconds number of seconds either way
Function Test-O365Session {
	
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession | Where { $_.ComputerName -eq "outlook.office365.com" }
	
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-Log "[ERROR] - No Session Found"
		Write-Log "Recreating Session"
		New-CleanO365Session
	}	
	# Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-Log "[ERROR] - Session not in Open State"
		Write-Log ($SessionInfo | fl | Out-String )
		Write-Log "Recreating Session"
		New-CleanO365Session
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-Log "Rebuilding Connection"
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		if ($DelayinSeconds -gt 0){
		
			Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		else {
			Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
				
		# new O365 session and reset our object processed count
		New-CleanO365Session
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	
	# If we have a manual throttle value then sleep for that many milliseconds
	if ($ManualThrottle -gt 0){
		Write-Log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

# If the $identifyingProperty has not been set then we attempt to locate a value for tracking modified objects
Function Get-ObjectIdentificationProperty {
	Param($object)
	
	Write-Log "Trying to identify a property for displaying per object progress"
	
	# Common properties to check
	[array]$PropertiesToCheck = "DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"
	
	# Set our counter to 0
	$i = 0
	
	# While we haven't found an ID property continue checking
	while ([string]::IsNullOrEmpty($IdentifyingProperty))
	{
	
		# If we have gone thru the list then we need to throw an error because we don't have Identity information
		# Set the string to bogus just to ensure we will exit the while loop
		if ($i -gt ($PropertiesToCheck.length -1))
		{
			Write-Log "[ERROR] - Unable to find a common identity parameter in the input object"
			
			# Create an error message that has all of the valid property names that we are looking for
			$PropertiesToCheck | foreach { [string]$PropertiesString = $PropertiesString + "`"" + $_ + "`", " }
			$PropertiesString = $PropertiesString.TrimEnd(", ")
			[string]$errorstring = "Objects does not contain a common identity parameter " + $PropertiesString + " please use -IdentifyingProperty to set the identity value"
			
			# Throw error
			Write-Error -Message $errorstring -ErrorAction Stop
		}
		
		# Get the property we are testing out of our array
		[string]$Property = $PropertiesToCheck[$i]
	
		# Check the properties of the object to see if we have one that matches a well known name
		# If we have found one set the value to that property
		if ($object.$Property -ne $null)
		{ 
			Write-Log ("Found " + $Property + " to use for displaying per object progress")
			Set-Variable -Scope script -Name IdentifyingProperty -Value $Property
		}
		
		# Increment our position counter
		$i++
		
	}
}

# Gather and print out information about how fast the script is running
Function Get-EstimatedTimeToCompletion {
	param([int]$ProcessedCount)
	
	# Increment our count of how many objects we have processed
	$ProcessedCount++
	
	# Every 100 we need to estimate our completion time and write that out
	if (($ProcessedCount % 100) -eq 0){
	
		# Get the current date
		$CurrentDate = Get-Date
		
		# Average time per object in seconds
		$AveragePerObject = (((($CurrentDate) - $ScriptStartTime).totalseconds) / $ProcessedCount)
		
		# Write out session stats and estimated time to completion
		Write-Log ("[STATS] - Total Number of Objects:     " + $ObjectCount)
		Write-Log ("[STATS] - Number of Objects processed: " + $ProcessedCount)
		Write-Log ("[STATS] - Average seconds per object:  " + $AveragePerObject)
		Write-Log ("[STATS] - Estimated completion time:   " + $CurrentDate.addseconds((($ObjectCount - $ProcessedCount) * $AveragePerObject)))
	}
	
	# Return number of objects processed so that the variable in incremented
	return $ProcessedCount
}

<#
.SYNOPSIS
Generic wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

Works well with intense operations that may cause throttling

.DESCRIPTION
Wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

It accomplishs this by doing the following:
* Monitors the health of the Remote powershell session and restarts it as needed.
* Restarts the session every X number seconds to ensure a valid connection.
* Attempts to work past session related errors and will skip objects that it can't process.
* Attempts to calculate throttle exhaustion and sleep a sufficient time to allow throttle recovery

.PARAMETER Agree
Verifies that you have read and agree to the disclaimer at the top of the script file.

.PARAMETER AutomaticThrottle
Calculated value based on your tenants powershell recharge rate.
You tenant recharge rate can be calculated using a Micro Delay Warning message.

Look for the following line in your Micro Delay Warning Message
Balance: -1608289/2160000/-3000000 

The middle value is the recharge rate.
Divide this value by the number of milliseconds in an hour (3600000)
And subtract the result from 1 to get your AutomaticThrottle value

1 - (2160000 / 3600000) = 0.4

Default Value is .25

.PARAMETER Credential
Credential object for logging into Exchange Online Shell.
Prompts if there is non provided.

.PARAMETER IdentifyingProperty
What property of the objects we are processing that will be used to identify them in the log file and host
If the value is not set by the user the script will attempt to determine if one of the following properties is present
"DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"

If the value is not set and we are not able to match a well known property the script will generate an error and terminate.

.PARAMETER LogFile
Location and file name for the log file.

.PARAMETER ManualThrottle
Manual delay of X number of milliseconds to sleep between each cmdlets call.
Should only be used if the AutomaticThrottle isn't working to introduce sufficent delay to prevent Micro Delays

.PARAMETER NonInteractive
Suppresses output to the screen.  All output will still be in the log file.

.PARAMETER Recipients
Array of objects to operate on. This can be mailboxes or any other set of objects.
Input must be an array!
Anything comming in from the array can be accessed in the script block using $input.property

.PARAMETER ResetSeconds
How many seconds to run the script block before we rebuild the session with O365.

.PARAMETER ScriptBlock
The script that you want to robustly execute against the array of objects.  The Recipient objects will be provided to the cmdlets in the script block
and can be accessed with $input as if you were pipelining the object.

.LINK
http://EHLO.Link

.OUTPUTS
Creates the log file specified in -logfile.  Logfile contains a record of all actions taken by the script.

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\mbx.csv
$mbx = import-csv c:\temp\mbx.csv
$cred = get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $mbx -logfile C:\temp\out.log -ScriptBlock {Set-Clutter -identity $input.PrimarySMTPAddress.tostring() -enable:$false}

Gets all mailboxes from the service returning only Displayname,Identity, and PrimarySMTPAddress.  Exports the results to a CSV
Imports the CSV into a variable
Gets your O365 Credential
Executes the script setting clutter to off

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\recipients.csv
$recipients = import-csv c:\temp\recipients.csv
$cred = Get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $recipients -logfile C:\temp\out.log -ScriptBlock {Get-MobileDeviceStatistics -mailbox $input.PrimarySMTPAddress.tostring() | Select-Object -Property @{Name = "PrimarySMTPAddress";Expression={$input.PrimarySMTPAddress.tostring()}},DeviceType,LastSuccessSync,FirstSyncTime | Export-Csv c:\temp\stats.csv -Append }

Gets All Recipients and exports them to a CSV (for restartability)
Imports the CSV into a variable
Gets your O365 Credentials
Executs the script to gather EAS Device statistics and output them to a csv file
#>
Function Start-RobustCloudCommand
{
	Param(
		[switch]$Agree,
		#[Parameter(Mandatory=$true)]
		#[string]$LogFile,
		[Parameter(Mandatory=$true)]
		[array]$Recipients,
		[Parameter(Mandatory=$true)]
		[ScriptBlock]$ScriptBlock,
		[Parameter(Mandatory=$false)]
		[Object[]]$ArgumentList ,
		[Parameter(Mandatory=$true)]
		$UserName,
		[int]$ManualThrottle=0,
		[double]$ActiveThrottle=.25,
		[int]$ResetSeconds=870,
		[string]$IdentifyingProperty,
		[Switch]$NonInteractive
	)

	####################
	# Main Function
	####################

	# Force use of at least version 3 of powershell https://technet.microsoft.com/en-us/library/hh847765.aspx
	#Requires -version 3

	# Turns on strict mode https://technet.microsoft.com/library/03373bbe-2236-42c3-bf17-301632e0c428(v=wps.630).aspx
	Set-StrictMode -Version 2

	# Write creation date of script for version information
	Write-Log "Created 05/10/2016"

	# Statement to ensure that you have looked at the disclaimer or that you have removed this line so you don't have too
	if ($Agree -ne $true){ Write-Error "Please run the script with -Agree to indicate that you have read and agreed to the sample script disclaimer at the top of the script file" -ErrorAction Stop }
	else { Write-Log "Agreed to Disclaimer" }

	# Log the script block for debugging purposes
	Write-Log $ScriptBlock

	# Setup our first session to O365
	$ErrorCount = 0
	New-CleanO365Session | Out-Null

	# Get when we started the script for estimating time to completion
	$ScriptStartTime = Get-Date

	# Get the object count and write it out to be used in esitmated time to completion + logging
	[int]$ObjectCount = $Recipients.count
	[int]$ObjectsProcessed = 0

	# If we don't have an identifying property then try to find one
	if ([string]::IsNullOrEmpty($IdentifyingProperty))
	{
		# Call our function for finding an identifying property and pass in the first recipient object
		Get-ObjectIdentificationProperty -object $Recipients[0]
	}

	# Go thru each recipient object and execute the script block
	$LoopIndex = 0
	$MaxCount = $Recipients.Count
	foreach ($object in $Recipients)
	{
		++$LoopIndex
			
		# Set our initial while statement values
		$TryCommand = $true
		$errorcount = 0
	
		# Try the command 3 times and exit out if we can't get it to work
		# Record the error and restart the session each time it errors out
		while ($TryCommand)
		{
			Write-Log ("Running scriptblock ($LoopIndex of $MaxCount) for " + ($object.$IdentifyingProperty).tostring())
		
			# Clear all errors
			$Error.Clear()
			$Global:Error.Clear()
	
			# Test our connection and rebuild if needed
			Test-O365Session
	
			# Invoke the script block
			try
			{
				$currentErrorActionPreference = $Global:ErrorActionPreference

				$Global:ErrorActionPreference = 'Stop'
				Invoke-Command -InputObject $object -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction 'Stop' 2>> ($Global:__LogFile__ + ".err")

				$Global:ErrorActionPreference = $currentErrorActionPreference
			}
			catch
			{
				$Global:ErrorActionPreference = $currentErrorActionPreference
				Write-Log "[ERROR] - Exception caught while invoking command. Error: $_"
			}
		
			# Test for errors. For some reason the exception increases the error count in the $Global:Error only. So check $Global:Error.Count as well
			if ($Error.Count -gt 0 -or $Global:Error.Count -gt 0) 
			{
				# Write that we failed
				Write-Log ("[ERROR] - Failed on object " + ($object.$IdentifyingProperty).tostring())
				if ($Error.Count  -gt 0)
				{
					Write-Log $Error
				}
				else
				{
					Write-Log $Global:Error
				}
			
				# Increment error count
				$errorcount++
			
				# Handle if we keep failing on the object
				if ($errorcount -ge 3)
				{
					Write-Log ("[ERROR] - Oject has failed three times " + ($object.$IdentifyingProperty).tostring())
					Write-Log ("[ERROR] - Skipping Object")
					
					# Increment the object processed count / Estimate time to completion
					$ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
					
					# Set trycommand to false so we abort the while loop
					$TryCommand = $false
				}
				# Otherwise try the command again
				else 
				{
					Write-Log ("Rebuilding session and trying again")
					# Create a new session in case the error was due to a session issue
					New-CleanO365Session | Out-Null 
				}
			}
			else 
			{
				# Since we didn't get an error don't run again
				$TryCommand = $false
			
				# Increment the object processed count / Estimate time to completion
				$ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
			}
		}
	}

	Write-Log "Script Complete Destroying PS Sessions"
	
	# Destroy any outstanding PS Session
	Get-PSSession | Remove-PSSession -Confirm:$false
}

Export-ModuleMember -Function Write-Log
Export-ModuleMember -Function Start-RobustCloudCommand
