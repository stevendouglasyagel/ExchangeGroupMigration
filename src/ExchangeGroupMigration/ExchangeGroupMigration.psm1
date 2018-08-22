#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This is a library of PowerShell cmdlets to help migrating on-prem mail-enabled groups to Exchange Online managed classic groups. 
#
#	Edit the Global Varibales in the top most region appropriately.
#########################################################################################################################################################################

#region Global Variables. Must be edited appropriately

$Global:TestTenant = $false
$Global:GroupForest = "CONTOSO" # "TH-CONTOSO" # "CONTOSOMY" # "CONTOSO-EU"
$Global:SmtpDomains =  "@*" #@("@uk.contosolab.com", "@in.contosolab.com")

$Global:__UPN__ =  if ($Global:TestTenant -eq $true) { "nileshg@contosolab.onmicrosoft.com" } else { "a-th30@global.contoso.org" }
$Global:NewGroupNameSuffix = "#NEW#"
$Global:NewSmtpDomain = if ($Global:TestTenant -eq $true) { "@contosolab.com" } else { "@contoso.com" }
$Global:HybridEmailRoutingDomain = if ($Global:TestTenant -eq $true) { "@contosolab.mail.onmicrosoft.com" } else { "@contoso.mail.onmicrosoft.com" }
$Global:DefaultGroupOwner = "DL_Admin" + $Global:NewSmtpDomain # must be an cloud MESG
$Global:OnpremExchangeUri = switch ($Global:GroupForest) {
		"CONTOSO" { if ($Global:TestTenant -eq $true) { "http://OC-EXCH-01.ukroi.contosolab.net/PowerShell" } else { "http://OC-EXCH-01.ukroi.contoso.org/PowerShell" } }
		"TH-CONTOSO" { if ($Global:TestTenant -eq $true) { "http://TH-EXCH-01.th.contosolab.net/PowerShell" } else { "http://TH-EXCH-01.TH-CONTOSO.ORG/PowerShell" } }
		"CONTOSOMY" { if ($Global:TestTenant -eq $true) { "http://MY-EXCH-01.my.contosolab.net/PowerShell" } else { "http://MY-EXCH-01.my.contoso.com/PowerShell" } }
		"CONTOSO-EU" { if ($Global:TestTenant -eq $true) { "http://CE-EXCH-01.ce.contosolab.net/PowerShell" } else { "http://CE-EXCH-01.contoso-europe.com/Powershell" } }
	}

# The OU which is excluded from syncing to AAD. This is because we don't have "DoNotSync" filter on contacts
$Global:ContactSyncExclusionOU = switch ($Global:GroupForest) {
		"CONTOSO" { if ($Global:TestTenant -eq $true) { "OU=GALSyncContacts,OU=XX,DC=ukroi,DC=contosolab,DC=net" } else { "OU=GALSyncContacts,OU=XX,DC=ukroi,DC=contoso,DC=org" } }
		"TH-CONTOSO" { if ($Global:TestTenant -eq $true) { "OU=GALSyncContacts,OU=XX,DC=th,DC=contosolab,DC=net" } else { "OU=GALSyncContacts,OU=XX,DC=TH-CONTOSO,DC=ORG" } }
		"CONTOSOMY" { if ($Global:TestTenant -eq $true) { "OU=GALSyncContacts,OU=XX,DC=my,DC=contosolab,DC=ne" } else { "OU=GALSyncContacts,OU=XX,DC=my,DC=contoso,DC=com" } }
		"CONTOSO-EU" { if ($Global:TestTenant -eq $true) { "OU=GALSyncContacts,DC=ce,DC=contosolab,DC=net" } else { "OU=GALSyncContacts,DC=contoso-europe,DC=com" } }
	}

# Domain Controllers used by AADC to avoid any replication delays
$Global:DomainControllerLookup = @{
		"DC=UKROI,DC=*" = if ($Global:TestTenant -eq $true) { "OC-ADDS-03.ukroi.contosolab.net" } else {  "OC-ADDS-03.ukroi.contoso.org" };
		"DC=IN,DC=*" =  if ($Global:TestTenant -eq $true) { "OC-ADDS-05.in.contosolab.net" } else { "OC-ADDS-05.in.contoso.org" };
		"DC=TSL,DC=*" =  if ($Global:TestTenant -eq $true) { "OC-ADDS-07.tsl.contosolab.net" } else { "OC-ADDS-07.tsl.contoso.org" };
		"DC=TH-CONTOSO,DC=*" =  if ($Global:TestTenant -eq $true) { "TH-ADDS-01.th.contosolab.net" } else { "TH-ADDS-01.th-contoso.org" };
		"DC=MY,DC=*" =  if ($Global:TestTenant -eq $true) { "MY-ADDS-01.my.contosolab.net" } else { "MY-ADDS-01.my.contoso.com" };
		"DC=CZ,DC=*" =  if ($Global:TestTenant -eq $true) { "CE-ADDS-03.cz.ce.contosolab.net" } else { "CE-ADDS-03.cz.contoso-europe.com" };
		"DC=HU,DC=*" =  if ($Global:TestTenant -eq $true) { "CE-ADDS-05.hu.ce.contosolab.net" } else { "CE-ADDS-05.hu.contoso-europe.com" };
		"DC=PL,DC=*" =  if ($Global:TestTenant -eq $true) { "CE-ADDS-07.pl.ce.contosolab.net" } else { "CE-ADDS-07.pl.contoso-europe.com" };
		"DC=SK,DC=*" =  if ($Global:TestTenant -eq $true) { "CE-ADDS-09.sk.ce.contosolab.net" } else { "CE-ADDS-09.sk.contoso-europe.com" };
		"DC=CONTOSO-EUROPE,DC=*" =  if ($Global:TestTenant -eq $true) { "CE-ADDS-01.ce.contosolab.net" } else { "CE-ADDS-01.contoso-europe.com" };
	}

#endregion Global Variables. Must be edited appropriately

#region Helper Global Variables

$Global:__LogFile__ = ".\Logs\ExchangeGroupMigration_" + [Guid]::NewGuid() + ".log"
$Global:GroupsToMigrateTxtFileName = ".\GroupsToMigrate.txt" # Need only for pilot / test migration
$Global:GroupsToDisableTxtFileName = ".\GroupsToDisable.txt" # Need only for pilot / test migration
$Global:GroupsToDeleteTxtFileName = ".\GroupsToDelete.txt"
$Global:GroupsGoodToDeleteTxtFileName = ".\GroupsGoodToDelete.txt"

$Global:OnpremGroupExportFileName = ".\Exports\Onprem-DL-Data.json"
$Global:OnpremGroupMemberExportFileName = ".\Exports\Onprem-DL-Member-Data.json"
$Global:OnpremDeliveryRestrictionsExportFileName = ".\Exports\Onprem-Delilvery-Restrictions-And-Delegates-Data.json"
$Global:OnpremMailboxExportFileName = ".\Exports\Onprem-Mailbox-Data.json"
$Global:OnpremFullAccessPermissionsExportFileName = ".\Exports\Onprem-FullAccess-Permissions-Data.json"
$Global:OnpremSendAsPermissionsExportFileName = ".\Exports\Onprem-SendAs-Permissions-Data.json"
$Global:OnpremCalendarSchedulingPermissionsExportFileName = ".\Exports\Onprem-CalendarScheduling-Permissions-Data.json"
$Global:OnpremGroupsFailingDeliveryRestrictionsValidationFileName = ".\Exports\Onprem-Groups-Failing-Delilvery-Restrictions-And-Delegates-Validation.json"
$Global:OnpremGroupsFailingFullAccessDelegationsValidationFileName = ".\Exports\Onprem-Groups-Failing-FullAccess-Delegations-Validation.json"
$Global:OnpremGroupsFailingSendAsDelegationsValidationFileName = ".\Exports\Onprem-Groups-Failing-SendAs-Delegations-Validation.json"
$Global:OnpremGroupsFailingCalendarSchedulingDelegationsValidationFileName = ".\Exports\Onprem-Groups-Failing-CalendarScheduling-Delegations-Validation.json"
$Global:OnpremGroupsFailingNestingValidationFileName = ".\Exports\Onprem-Groups-Failing-Nesting-Validation.json"
$Global:OnpremGroupsExcludedFileName = ".\Exports\Onprem-Groups-Excluded-All.json"
$Global:OnpremGroupsToMigrateFileName = ".\Exports\Onprem-Groups-to-Migrate-Data.json"
$Global:OnpremGroupsGoodToMigrateFileName = ".\Exports\Onprem-Groups-Good-to-Migrate-Data.json"
$Global:OnpremGroupsGoodToMigrateWithOwnersFileName = ".\Exports\Onprem-Groups-Good-to-Migrate-Owner-Data.json"
$Global:OnpremGroupsGoodToDisableFileName = ".\Exports\Onprem-Groups-Good-to-Disable-Data.json"
$Global:OnpremGroupsFailingNestingValidationForDisablementFileName = ".\Exports\Onprem-Groups-Failing-Nesting-Validation-For-Disablement.json"
$Global:OnpremAdGroupExportFileName = ".\Exports\Onprem-ADGroup-Data.json"
$Global:OnpremGroupsGoodToDeleteFileName = ".\Exports\Onprem-Groups-Good-to-Delete-Data.json"
$Global:OnpremGroupsFailingNestingValidationForDeletionFileName = ".\Exports\Onprem-Groups-Failing-Nesting-Validation-For-Deletion.json"

$Global:OnlineGroupExportFileName = ".\Exports\Online-DL.json"
$Global:OnlineGroupMemberExportFileName = ".\Exports\Online-DL-Member.json"
$Global:OnlineDeliveryRestrictionsExportFileName = ".\Exports\Online-Recipient-Restrictions-and-Delegates.json"
$Global:OnlineMailboxExportFileName = ".\Exports\Online-Mailbox.json"
$Global:OnlineFullAccessPermissionsExportFileName = ".\Exports\Online-FullAccess-Permissions.json"
$Global:OnlineSendAsPermissionsExportFileName = ".\Exports\Online-SendAs-Permissions.json"
$Global:OnlineCalendarSchedulingPermissionsExportFileName = ".\Exports\Online-CalendarScheduling-Permissions.json"
$Global:OnlineGroupsToMigrateFileName = ".\Exports\Online-Groups-to-Migrate-Data.json"
$Global:OnlineGroupsGoodToMigrateFileName = ".\Exports\Online-Groups-Good-to-Migrate-Data.json"
$Global:OnlineGroupsGoodToMigrateAsSecurityGroupsFileName = ".\Exports\Online-Groups-Good-to-Migrate-As-SecurityGroups-Data.json"

#endregion Helper Global Variables

# Require at least PowerShell 5. Otherwise ConvertFrom-Json cmdlet fails when the input file is larger than 5 MB.
# For a work-around see: https://social.technet.microsoft.com/Forums/windowsserver/en-US/833c99c1-d8eb-400d-bf58-38f7265b4b0e/error-when-converting-from-json?forum=winserverpowershell&prof=required
if ($PSVersionTable.PSVersion.Major -lt 5) {throw "This script needs PowerShell 5 or later."}

$Global:ErrorActionPreference = "Continue"
$Global:DebugPreference = "Continue"
$Global:VerbosePreference = "SilentlyContinue"

$Global:RecipientGuidCache = @{"__RecipientName__"="__RecipientGuid__"}

#### Import the RobustCloudCommand module
if (Get-Module -Name RobustCloudCommand) { Remove-Module -Name RobustCloudCommand }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "RobustCloudCommand.psm1") -Global -ErrorAction Stop

#region Utility Functions

<#
.Synopsis
	Returns the current script name. 
.DESCRIPTION
	Returns the current script name for logging purposes. 
.EXAMPLE
	Get-CurrentScriptName
#>
function Get-CurrentScriptName
{
	[CmdletBinding()]
	param
	(
	)
	
	Split-Path $MyInvocation.PSCommandPath -Leaf
}

<#
.Synopsis
	Returns the onprem Domain Controller FQDN based on the supplied DN. 
.DESCRIPTION
	Returns the onprem Domain Controller FQDN to be used in the subsequent cmdlets based on the DN of the object.
	Returns the Domain Controller from the lookup specified in the global variable $Global:DomainControllerLookup.
.EXAMPLE
	Get-DomainController -DN "CN=TestDL,OU=Groups,OU=Corp,DC=contoso,DC=com"
#>
function Get-DomainController
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateNotNull()]
		[string]
		$dn
	)
	
	$dc = $Global:DomainControllerLookup.Keys | % { if ($dn -match $_) { $Global:DomainControllerLookup[$_] } }

	if ($dc.Count -gt 1)
	{
		# return root domain only if the child domain is not present
		if (($dc[0] -split ".").Count -gt ($dc[1] -split ".").Count)
		{
			$dc[0]
		}
		else
		{
			$dc[1]
		}
	}
	else
	{
		$dc
	}
}

function Get-ShadowGroupName ([string] $GroupName, [string] $ShadowGroupSuffix = $Global:NewGroupNameSuffix)
{
	if ($GroupName.EndsWith($ShadowGroupSuffix))
	{
		$GroupName
		return
	}

	if ($GroupName.Length -gt 64 - $ShadowGroupSuffix.Length)
	{
        $GroupName = $GroupName -replace '\s',''
        if ($GroupName.Length -gt 64 - $ShadowGroupSuffix.Length)
        {
            $GroupName.Substring(0, 64 - $ShadowGroupSuffix.Length) + $ShadowGroupSuffix 
        }
        else
        {
            $GroupName + $ShadowGroupSuffix
        }
	}
	else
	{
		$GroupName + $ShadowGroupSuffix
	}
}

function Get-ShadowGroupAlias ([string] $GroupAlias, [string] $ShadowGroupSuffix = $Global:NewGroupNameSuffix)
{
	if ($GroupAlias.EndsWith($ShadowGroupSuffix))
	{
		$GroupAlias
		return
	}

	if ($GroupAlias.Length -gt 64 - $ShadowGroupSuffix.Length)
	{
        $GroupAlias = $GroupAlias -replace '\s',''
        if ($GroupAlias.Length -gt 64 - $ShadowGroupSuffix.Length)
        {
            $GroupAlias.Substring(0, 64 - $ShadowGroupSuffix.Length) + $ShadowGroupSuffix 
        }
        else
        {
            $GroupAlias + $ShadowGroupSuffix
        }
	}
	else
	{
		$GroupAlias + $ShadowGroupSuffix
	}
}

function Get-ShadowGroupPrimarySmtpAddress ([string] $GroupPrimarySmtpAddress, [string] $ShadowGroupSuffix = $Global:NewGroupNameSuffix, [string] $ShadowGroupSmtpDomain = $Global:NewSmtpDomain)
{
	# TODO: handle use case when the shadow group domain is to be kept the same
	if ($GroupPrimarySmtpAddress.EndsWith($ShadowGroupSmtpDomain))
	{
		$GroupPrimarySmtpAddress
		return
	}

	if ($ShadowGroupSmtpDomain)
	{
		$GroupPrimarySmtpAddress.Split('@')[0] + $ShadowGroupSmtpDomain
	}
	else
	{
		$GroupPrimarySmtpAddress.Split('@')[0] + $ShadowGroupSuffix + "@" + $PrimarySmtpAddress.Split('@')[1]
	}
}

<#
.Synopsis
	Removes broken and closed sessions
.DESCRIPTION
	Removes broken and closed sessions
.EXAMPLE
	Remove-BrokenOrClosedPSSession
#>
function Remove-BrokenOrClosedPSSession
{
    Get-PSSession | Where { $_.State -like "*Broken*" } | % { if ($_) { Remove-PSSession -session $_ } }
    Get-PSSession | Where { $_.State -like "*Closed*" } | % { if ($_) { Remove-PSSession -session $_ } }
}

<#
.Synopsis
	Creates a new Onprem Exchange Session or returns the existing Onprem Exchange session
.DESCRIPTION
	Creates a new Onprem Exchange Session or returns the existing Onprem Exchange session
.EXAMPLE
	New-OnpremExchangeSession
	New-OnpremExchangeSession -Force
#>
function New-OnpremExchangeSession
{
	[CmdletBinding()]
	param
	(
		[switch]
		$Force
	)

	Remove-BrokenOrClosedPSSession

	$OnpremExchangeSession = $null
	$OnpremExchangeSession = Get-PSSession | Where { $Global:OnpremExchangeUri -like "*//$($_.ComputerName)/*" } 
	
	if ($Force -or $OnpremExchangeSession -eq $null)
	{
		$OnpremExchangeSession | Remove-PSSession

		Write-Debug "Creating Onprem Exchange Remote PowerShell session for: '$Global:OnpremExchangeUri'..."

		if ($Global:OnpremExchangeUserCredential)
		{
			$OnpremExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Global:OnpremExchangeUri -Authentication Kerberos -AllowRedirection -Credential $Global:OnpremExchangeUserCredential 
		}
		else
		{
			$OnpremExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Global:OnpremExchangeUri -Authentication Kerberos -AllowRedirection 
		}

		Import-Module (Import-PSSession $OnpremExchangeSession -AllowClobber -Prefix "Onprem" -DisableNameChecking) -Prefix "Onprem" -Global -DisableNameChecking
	}
}

#endregion Utility Functions

#region Onprem "Get" Cmdlets

<#
.Synopsis
	Returns the membership of the specified onprem distribtution groups. 
.DESCRIPTION
	Returns the membership of the specified onprem distribtution groups. 
.PARAMETER GroupsJsonFilePath
	The file path of the json file containing groups information previously saved. 
.EXAMPLE
	Get-OnpremDistributionGroupMembers -GroupsJsonFilePath $Global:OnpremGroupExportFileName
#>
function Get-OnpremDistributionGroupMembers
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath"

		New-OnpremExchangeSession
	}

	process
	{
		$groups = Get-Content $GroupsJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnpremGroupMemberExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

        $count = $groups.Count
        $index = 0
		foreach ($group in $groups)
		{
            ++$index

			Write-Log "($index/$count) Reading membership for group: '$($group.Name)' - '$($group.PrimarySmtpAddress)'"

			$distinguishedName = $group.DistinguishedName
			$domainController = Get-DomainController $distinguishedName
            try
            {
				$members = [array](Get-OnpremDistributionGroupMember -ResultSize Unlimited -IgnoreDefaultScope -Identity $distinguishedName `
					| Select -Property Name, Identity, RecipientType, RecipientTypeDetails, PrimarySmtpAddress)

				$props = @{
						"Name" = $group.Name;
						"Identity" = $group.Identity;
						"Guid" = $group.Guid;
						"DistinguishedName" = $group.DistinguishedName;
						"PrimarySmtpAddress" = $group.PrimarySmtpAddress;
						"Members" = $members;
					}

				$output = New-Object –TypeName PSObject –Prop $props

				$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

				$output # send the output on the return pipeline 
            }
            catch
            {
                Write-Error "Error: $_"
            }
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath"
	}
}

<#
.Synopsis
	Returns the delivery restrictions and SendOnBehalfTo grants of onprem recipient objects. 
.DESCRIPTION
	Returns the delivery restrictions and SendOnBehalfTo grants of onprem recipient objects.
	Since the cmdlet returns relatively quickly, does not use any filtering based on the list specified in $Global:SmtpDomains.
.EXAMPLE
	Get-OnpremRecipientDeliveryRestrictionsAndPublicDelegates
#>
function Get-OnpremRecipientDeliveryRestrictionsAndPublicDelegates
{
	[CmdletBinding()]
	param
	(
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand)."

		New-OnpremExchangeSession
	}
	
	process
	{
		Write-Log "Reading Mailboxes with AcceptMessagesOnlyFromDLMembers/RejectMessagesFromDLMembers/GrantSendOnBehalfTo configured..."
		$output = Get-OnpremMailbox -ResultSize Unlimited -IgnoreDefaultScope -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
			| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, DistinguishedName, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers `
		$output # send the output on the return pipeline 

		Write-Log "Reading DistributionGroups with AcceptMessagesOnlyFromDLMembers/RejectMessagesFromDLMembers/GrantSendOnBehalfTo configured..."
		$output = Get-OnpremDistributionGroup -ResultSize Unlimited -IgnoreDefaultScope -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
			| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, DistinguishedName, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers `
		$output # send the output on the return pipeline 

		Write-Log "Reading MailUsers with AcceptMessagesOnlyFromDLMembers/RejectMessagesFromDLMembers/GrantSendOnBehalfTo configured..."
		$output = Get-OnpremMailUser -ResultSize Unlimited -IgnoreDefaultScope -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
			| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, DistinguishedName, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers `
		$output # send the output on the return pipeline 

		Write-Log "Reading MailContacts with AcceptMessagesOnlyFromDLMembers/RejectMessagesFromDLMembers/GrantSendOnBehalfTo configured..."
		$output = Get-OnpremMailContact -ResultSize Unlimited -IgnoreDefaultScope -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
			| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, DistinguishedName, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers `
		$output # send the output on the return pipeline 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand)."
	}
}

<#
.Synopsis
	Returns the FullAccess permissions on the specified onprem mailboxes. 
.DESCRIPTION
	Returns the FullAccess permissions on the specified onprem mailboxes. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnpremMailboxFullAccessPermissions -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName
#>
function Get-OnpremMailboxFullAccessPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing mailbox information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"

		New-OnpremExchangeSession
	}

	process
	{
		$mailboxes = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnpremFullAccessPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

        $count = $mailboxes.Count
        $index = 0
		foreach ($mailbox in $mailboxes)
		{
            ++$index

			Write-Log "($index/$count) Reading FullAccess permissions for mailbox: '$($mailbox.Name)' - '$($mailbox.PrimarySmtpAddress)'"

			$distinguishedName = $mailbox.DistinguishedName
			$domainController = Get-DomainController $distinguishedName
            try
            {
			    $output = Get-OnpremMailboxPermission -Identity $distinguishedName -DomainController $domainController | Where { $_.AccessRights -like "*FullAccess*" -and $_.IsInherited -eq $false -and $_.User -ne "NT AUTHORITY\SELF" -and $_.User -notlike "S-1-5-21*" } `
					    | Select Identity, User, AccessRights, Deny

				$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

				$output # send the output on the return pipeline 
            }
            catch
            {
                Write-Error "Error: $_"
            }
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

<#
.Synopsis
	Returns the SendAs permissions that are configured on the specified onprem groups and mailboxes. 
.DESCRIPTION
	Returns the SendAs permissions that are configured on the specified onprem groups and mailboxes. 
.PARAMETER GroupsJsonFilePath
	The file path of the json file containing groups information previously saved. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnpremRecipientSendAsPermissions -GroupsJsonFilePath $Global:OnpremGroupExportFileName -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName
#>
function Get-OnpremRecipientSendAsPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsJsonFilePath,

		# The file path of the json file containing mailboxes information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath MailboxesJsonFilePath: $MailboxesJsonFilePath"

		New-OnpremExchangeSession
	}

	process
	{
		$groups = Get-Content $GroupsJsonFilePath -Raw | ConvertFrom-Json
		$mailboxess = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnpremSendAsPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

		$recipients = @()
		$recipients += $groups | Select Name, DistinguishedName, PrimarySmtpAddress
		$recipients += $mailboxess | Select Name, DistinguishedName, PrimarySmtpAddress

        $count = $recipients.Count
        $index = 0
		foreach ($recipient in $recipients)
		{
            ++$index

			Write-Log "($index/$count) Reading SendAS permissions for recipient: '$($recipient.Name)' - '$($recipient.PrimarySmtpAddress)'"

			$distinguishedName = $recipient.DistinguishedName
			$domainController = Get-DomainController $distinguishedName
            try
            {
				$output = Get-OnpremADPermission -Identity $distinguishedName -DomainController $domainController | Where { $_.ExtendedRights -like "*Send-As*" -and $_.User -ne "NT AUTHORITY\SELF" -and $_.User -notlike "S-1-5-21*" } `
						| Select Identity, User, ExtendedRights

				$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

				$output # send the output on the return pipeline 
            }
            catch
            {
                Write-Error "Error: $_"
            }
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

<#
.Synopsis
	Returns the calendar processing configuration on the specified onprem (room) mailboxes. 
.DESCRIPTION
	Returns the calendar processing configuration on the specified onprem (room) mailboxes. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnpremCalendarProcessingConfiguration -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName
#>
function Get-OnpremCalendarProcessingConfiguration
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing mailbox information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"

		New-OnpremExchangeSession
	}

	process
	{
		$mailboxes = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		[array]$roomMailboxes = $mailboxes | Where { $_.RecipientTypeDetails -eq "RoomMailbox" }

		$dropFile = $Global:OnpremCalendarSchedulingPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

        $count = $roomMailboxes.Count
        $index = 0
		foreach ($roomMailbox in $roomMailboxes)
		{
            ++$index

			Write-Log "($index/$count) Reading CalendarProcessing configuration for room mailbox: '$($roomMailbox.Name)' - '$($roomMailbox.PrimarySmtpAddress)'"

			$distinguishedName = $roomMailbox.DistinguishedName
			$domainController = Get-DomainController $distinguishedName
            try
            {
			    $output = Get-OnpremCalendarProcessing -Identity $distinguishedName -DomainController $domainController

				if ($output.BookInPolicy -ne $null -or $output.RequestInPolicy -ne $null -or $output.RequestOutOfPolicy -ne $null)
				{
					$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

					$output # send the output on the return pipeline 
				}
            }
            catch
            {
                Write-Error "Error: $_"
            }
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

#endregion Onprem "Get" Cmdlets

#region Online "Get" Cmdlets

<#
.Synopsis
	Returns the EXO distribtution group matching specified SMTP domains in the global variable $Global:SmtpDomains. 
.DESCRIPTION
	Returns the EXO distribtution group matching specified SMTP domains in the global variable $Global:SmtpDomains.
	$Global:SmtpDomains is defined only when you are absolutely sure that cross-dmain objects are not used for configuring permissions and such
	on groups, mailboxes and other recipients.
.EXAMPLE
	Get-OnlineDistributionGroups
#>
function Get-OnlineDistributionGroups()
{
	[CmdletBinding()]
	param
	(
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). SmtpDomains: $Global:SmtpDomains"
	}

	process
	{
		$recipents = $Global:SmtpDomains | foreach { @{ "Name" = $_ } }

		$scriptBlock = {
			$smtpDomain = $Input.Name # InputObject from Start-RobustCloudCommand
			$smtpDomain = if($smtpDomain -notmatch "@") {"@$smtpDomain"} else { $smtpDomain }

			# PrimarySmtpAddress -like filter does not seems to work when started with a * inspite of documentation saying it should. Possible regression??
			# So using EmailAddresses instead for filtering based on domain
			$filter = "EmailAddresses -like '*$smtpDomain'"

			Write-Log "Reading online distribution groups. Filter: '$filter'"
			
			$output = Get-DistributionGroup -ResultSize Unlimited -Filter $filter

			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $recipents -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). SmtpDomains: $Global:SmtpDomains"
	}
}

<#
.Synopsis
	Returns the membership of the specified EXO distribtution groups. 
.DESCRIPTION
	Returns the membership of the specified EXO distribtution groups. 
.PARAMETER GroupsJsonFilePath
	The file path of the json file containing groups information previously saved. 
.EXAMPLE
	Get-OnlineDistributionGroupMembers -GroupsJsonFilePath $Global:OnlineGroupExportFileName
#>
function Get-OnlineDistributionGroupMembers
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath"
	}

	process
	{
		$groups = Get-Content $GroupsJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnlineGroupMemberExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

		$scriptBlock = {
			$group = $input # InputObject from Start-RobustCloudCommand
			$members = [array](Get-DistributionGroupMember -ResultSize Unlimited -Identity $group.Guid `
				| Select -Property Name, Identity, RecipientType, RecipientTypeDetails, PrimarySmtpAddress)

			$props = @{
					"Name" = $group.Name;
					"Identity" = $group.Identity;
					"Guid" = $group.Guid;
					"PrimarySmtpAddress" = $group.PrimarySmtpAddress;
					"Members" = $members;
				}

			$output = New-Object –TypeName PSObject –Prop $props

			$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $groups -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath"
	}
}

<#
.Synopsis
	Returns the delivery restrictions and SendOnBehalfTo grants of EXO recipient objects. 
.DESCRIPTION
	Returns the delivery restrictions and SendOnBehalfTo grants of EXO recipient objects.
	Since the cmdlet returns relatively quickly, does not use any filtering based on the list specified in $Global:SmtpDomains. 
.EXAMPLE
	Get-OnlineRecipientDeliveryRestrictionsAndPublicDelegates
#>
function Get-OnlineRecipientDeliveryRestrictionsAndPublicDelegates
{
	[CmdletBinding()]
	param
	(
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand)."
	}
	
	process
	{
		$filterAttributes = @("Mailbox", "DistributionGroup", "MailUser", "MailContact") | foreach { @{ "Name" = $_ } }

		$scriptBlock = {
			$recipient = $Input.Name # InputObject from Start-RobustCloudCommand

			switch ($recipient)
			{
				"Mailbox" {
					Write-Log "Reading Mailboxes with AcceptMessagesOnlyFromDLMembers/RejectMessagesFromDLMembers/GrantSendOnBehalfTo configured..."
					$output = Get-Mailbox -ResultSize Unlimited -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
						| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
					$output # send the output on the return pipeline 
				}

				"DistributionGroup" {
					Write-Log "Reading DistributionGroups with $filterAttribute configured..."
					$output = Get-DistributionGroup -ResultSize Unlimited -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
						| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
					$output # send the output on the return pipeline 
				}

				"MailUser" {
					# Any on-prem ones should have already been remediated
					Write-Log "Reading MailUsers with $filterAttribute configured..."
					$output = Get-MailUser -ResultSize Unlimited -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
						| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
					$output # send the output on the return pipeline 
				}

				"MailContact" {
					# Any on-prem ones should have already been remediated
					Write-Log "Reading MailContacts with $filterAttribute configured..."
					$output = Get-MailContact -ResultSize Unlimited -Filter { AcceptMessagesOnlyFromDLMembers -ne $null -or RejectMessagesFromDLMembers -ne $null -or GrantSendOnBehalfTo -ne $null } `
						| Select Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
					$output # send the output on the return pipeline 
				}
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $filterAttributes -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand)."
	}
}

<#
.Synopsis
	Returns all mailboxes matching specified SMTP domains in the global variable $Global:SmtpDomains (so that FullAccess and SendAs delegations can be subsequently queried). 
.DESCRIPTION
	Returns all mailboxes matching specified SMTP domains in the global variable $Global:SmtpDomains (so that FullAccess and SendAs delegations can be subsequently queried). 
	$Global:SmtpDomains is defined only when you are absolutely sure that cross-dmain objects are not used for configuring permissions and such
	on groups, mailboxes and other recipients.
	Attributes returned are: Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
.EXAMPLE
	Get-OnlineMailboxes
#>
function Get-OnlineMailboxes
{
	[CmdletBinding()]
	param
	(
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). SmtpDomains: $Global:SmtpDomains"
	}
	
	process
	{
		# If we ever need paging we'll see to it then. RIght not just make the single call

		$recipents = $Global:SmtpDomains | foreach { @{ "Name" = $_ } }
		$scriptBlock = {
			$smtpDomain = $Input.Name # InputObject from Start-RobustCloudCommand
			$smtpDomain = if($smtpDomain -notmatch "@") {"@$smtpDomain"} else { $smtpDomain }

			$filter = "EmailAddresses -like '*$smtpDomain'"

			$output = Get-Mailbox -ResultSize Unlimited  -Filter $filter | `
				 Select -Property Identity, Name, Guid, RecipientType, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, RejectMessagesFrom, RejectMessagesFromDLMembers
			
			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $recipents -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). SmtpDomains: $Global:SmtpDomains"
	}
}

<#
.Synopsis
	Returns the FullAccess permissions on the specified EXO mailboxes. 
.DESCRIPTION
	Returns the FullAccess permissions on the specified EXO mailboxes. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnlineMailboxFullAccessPermissions -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName
#>
function Get-OnlineMailboxFullAccessPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing mailboxes information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}

	process
	{
		$mailboxes = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnlineFullAccessPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

		$scriptBlock = {
			$mailbox = $input # InputObject from Start-RobustCloudCommand

			$output = Get-MailboxPermission -Identity $mailbox.Guid | Where { $_.AccessRights -like "*FullAccess*" -and $_.IsInherited -eq $false -and $_.User -ne "NT AUTHORITY\SELF" -and $_.User -notlike "S-1-5-21*" } `
				| Select Identity, User, AccessRights

			$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $mailboxes -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

<#
.Synopsis
	Returns the SendAs permissions that are configured on the specified EXO groups and mailboxes. 
.DESCRIPTION
	Returns the SendAs permissions that are configured on the specified EXO groups and mailboxes. 
.PARAMETER GroupsJsonFilePath
	The file path of the json file containing groups information previously saved. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnlineRecipientSendAsPermissions -GroupsJsonFilePath $Global:OnlineGroupExportFileName -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName
#>
function Get-OnlineRecipientSendAsPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsJsonFilePath,

		# The file path of the json file containing mailboxes information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath. MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
	
	process
	{
		$groups = Get-Content $GroupsJsonFilePath -Raw | ConvertFrom-Json
		$mailboxess = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		$dropFile = $Global:OnlineSendAsPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

		$recipients = @()
		$recipients += $groups | Select Name, PrimarySmtpAddress, Identity, Guid
		$recipients += $mailboxess | Select Name, PrimarySmtpAddress, Identity, Guid

		$scriptBlock = {
			$recipient = $input # InputObject from Start-RobustCloudCommand

			$output = Get-RecipientPermission -Identity $recipient.Guid | Where { $_.AccessRights -like "*SendAs*" -and $_.IsInherited -eq $false -and $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "S-1-5-21*" } `
				| Select Identity, Trustee, AccessRights

			$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $recipients -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $GroupsJsonFilePath. MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

<#
.Synopsis
	Returns the calendar processing configuration on the specified EXO (room) mailboxes. 
.DESCRIPTION
	Returns the calendar processing configuration on the specified EXO (room) mailboxes. 
.PARAMETER MailboxesJsonFilePath
	The file path of the json file containing mailboxes information previously saved. 
.EXAMPLE
	Get-OnlineCalendarProcessingConfiguration -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName
#>
function Get-OnlineCalendarProcessingConfiguration
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing mailbox information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxesJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}

	process
	{
		$mailboxes = Get-Content $MailboxesJsonFilePath -Raw | ConvertFrom-Json

		[array]$roomMailboxes = $mailboxes | Where { $_.RecipientTypeDetails -eq "RoomMailbox" }

		$dropFile = $Global:OnlineCalendarSchedulingPermissionsExportFileName + ".tmp"
		if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

		$scriptBlock = {
			$mailbox = $input # InputObject from Start-RobustCloudCommand

			$output = Get-CalendarProcessing -Identity $mailbox.Guid | Where { $_.BookInPolicy -ne $null -or $_.RequestInPolicy -ne $null -or $_.RequestOutOfPolicy -ne $null }

			$output |% { $_ | ConvertTo-Json | Out-File $dropFile -Append } # write individual object json so that it can be edited for array easily

			$output # send the output on the return pipeline 
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $roomMailboxes -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). MailboxesJsonFilePath: $MailboxesJsonFilePath"
	}
}

#endregion Online "Get" Cmdlets

#region Helper Functions for Validating that Groups are Good to Migrate

<#
.Synopsis
	Returns the groups that cannot be migrated because they are still used to configure delivery restrictions and public delegates on onprem mailboxes and not-to-be migrate groups. 
.DESCRIPTION
	Returns the groups that cannot be migrated because they are still used to configure delivery restrictions and public delegates on onprem mailboxes and not-to-be migrate groups. 
.PARAMETER DeliveryRestrictionsAndDelegationsJsonFilePath
	The file path of the json file containing delivery restrictions and public delegations information previously saved. 
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing candidate to-be-migrated groups information previously saved. 
.EXAMPLE
	Get-DistributionGroupsFailingDeliveryRestrictionsAndDelegatesValidation -DeliveryRestrictionsAndDelegationsJsonFilePath $Global:OnpremDeliveryRestrictionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName
#>
function Get-DistributionGroupsFailingDeliveryRestrictionsAndDelegatesValidation
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing delivery restrictions and public delegations previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$DeliveryRestrictionsAndDelegationsJsonFilePath,

		# The file path of the json file containing candidate to-be-migrated groups information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). DeliveryRestrictionsAndDelegationsJsonFilePath: $DeliveryRestrictionsAndDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}

	process
	{
		$recipients = Get-Content $DeliveryRestrictionsAndDelegationsJsonFilePath -Raw | ConvertFrom-Json
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json

        $count = $groupsToMigrate.Count
        $index = 0
		foreach ($groupToMigrate in $groupsToMigrate)
		{
            ++$index

			Write-Log "($index/$count) Processing delivery restrictions and public delegations validation for group: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"

			$allRecipientsNeedingRemediation = $recipients | Where { $_.GrantSendOnBehalfTo -eq  $groupToMigrate.Identity -or  `
												 $_.AcceptMessagesOnlyFromDLMembers -eq  $groupToMigrate.Identity -or `
												 $_.RejectMessagesFromDLMembers -eq  $groupToMigrate.Identity }

			# Filter out all groups that we are migrating from the result set
			$nonMigratingRecipientsNeedingRemediation = $allRecipientsNeedingRemediation |  where { $_.Identity -notin  $groupsToMigrate.Identity }

			if ($nonMigratingRecipientsNeedingRemediation)
			{
				$msg = "[ERROR] - Failed Delivery Restrictions and Public Delegates Validation: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"
				Write-Warning $msg
				Write-Log $msg

				$props = @{
						"Name" = $groupToMigrate.Name;
						"Identity" = $groupToMigrate.Identity;
						"PrimarySmtpAddress" = $groupToMigrate.PrimarySmtpAddress;
						"Recipient" = @($nonMigratingRecipientsNeedingRemediation);
					}

				$failedGroup = New-Object –TypeName PSObject –Prop $props
				$failedGroup # send the output on the return pipeline 
			}
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). DeliveryRestrictionsAndDelegationsJsonFilePath: $DeliveryRestrictionsAndDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}
}

<#
.Synopsis
	Returns the groups that cannot be migrated because they are still used to configure FullAccess delegations on onprem mailboxes. 
.DESCRIPTION
	Returns the groups that cannot be migrated because they are still used to configure FullAccess delegations on onprem mailboxes. 
.PARAMETER FullAccessDelegationsJsonFilePath
	The file path of the json file containing FullAccess delegations information previously saved. 
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing candidate to-be-migrated groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingFullAccessDelegationsValidation -FullAccessDelegationsJsonFilePath $Global:OnpremFullAccessPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName
#>
function Get-DistributionGroupsFailingFullAccessDelegationsValidation
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing FullAccess delegations previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$FullAccessDelegationsJsonFilePath,

		# The file path of the json file containing candidate to-be-migrated groups information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). FullAccessDelegationsJsonFilePath: $FullAccessDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}

	process
	{
		$mailboxes = Get-Content $FullAccessDelegationsJsonFilePath -Raw | ConvertFrom-Json
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json

        $count = $groupsToMigrate.Count
        $index = 0
		foreach ($groupToMigrate in $groupsToMigrate)
		{
            ++$index

			Write-Log "($index/$count) Processing Mailbox FullAccess validation for group: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"

			$mailboxesNeedingRemediation = $mailboxes | Where { $_.User -match  "$($groupToMigrate.SamAccountName)`$" }

			if ($mailboxesNeedingRemediation)
			{
				$msg = "[ERROR] - Failed Mailbox Remediation Validation: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"
				Write-Warning $msg
				Write-Log $msg

				$props = @{
						"Name" = $groupToMigrate.Name;
						"Identity" = $groupToMigrate.Identity;
						"PrimarySmtpAddress" = $groupToMigrate.PrimarySmtpAddress;
						"Recipient" = @($mailboxesNeedingRemediation);
					}

				$failedGroup = New-Object –TypeName PSObject –Prop $props
				$failedGroup # send the output on the return pipeline 
			}
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). FullAccessDelegationsJsonFilePath: $FullAccessDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}
}

<#
.Synopsis
	Returns the groups that cannot be migrated because they are still used to configure SendAs delegations on onprem recipients. 
.DESCRIPTION
	Returns the groups that cannot be migrated because they are still used to configure SendAs delegations on onprem recipients. 
.PARAMETER FullAccessDelegationsJsonFilePath
	The file path of the json file containing SendAs delegations information previously saved. 
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing candidate to-be-migrated groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingSendAsDelegationsValidation -SendAsDelegationsJsonFilePath $Global:OnpremSendAsPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName
#>
function Get-DistributionGroupsFailingSendAsDelegationsValidation
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing SendAs delegations previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$SendAsDelegationsJsonFilePath,

		# The file path of the json file containing groups to migrate information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsJsonFilePath: $SendAsDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}

	process
	{
		$recipients = Get-Content $SendAsDelegationsJsonFilePath -Raw | ConvertFrom-Json
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json

        $count = $groupsToMigrate.Count
        $index = 0
		foreach ($groupToMigrate in $groupsToMigrate)
		{
            ++$index

			Write-Log "($index/$count) Processing Mailbox SendAs validation for group: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"

			$recipientsNeedingRemediation = $recipients | Where { $_.User -match  "$($groupToMigrate.SamAccountName)`$" }

			if ($recipientsNeedingRemediation)
			{
				$msg = "[ERROR] - Failed Mailbox Remediation Validation: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"
				Write-Warning $msg
				Write-Log $msg

				$props = @{
						"Name" = $groupToMigrate.Name;
						"Identity" = $groupToMigrate.Identity;
						"PrimarySmtpAddress" = $groupToMigrate.PrimarySmtpAddress;
						"Recipient" = @($recipientsNeedingRemediation);
					}

				$failedGroup = New-Object –TypeName PSObject –Prop $props
				$failedGroup # send the output on the return pipeline 
			}
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). SendAsDelegationsJsonFilePath: $SendAsDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}
}

<#
.Synopsis
	Returns the groups that cannot be migrated because they are still used to configure calender processing policies on onprem mailboxes. 
.DESCRIPTION
	Returns the groups that cannot be migrated because they are still used to configure calender processing policies on onprem mailboxes. 
.PARAMETER CalendarSchedulingDelegationsJsonFilePath
	The file path of the json file containing calender processing configuration information previously saved. 
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing candidate to-be-migrated groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingCalendarSchedulingDelegationsValidation -CalendarSchedulingDelegationsJsonFilePath $Global:OnpremCalendarSchedulingPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName
#>
function Get-DistributionGroupsFailingCalendarSchedulingDelegationsValidation
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing calender processing configuration information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$CalendarSchedulingDelegationsJsonFilePath,

		# The file path of the json file containing mailbox information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). CalendarSchedulingDelegationsJsonFilePath: $CalendarSchedulingDelegationsJsonFilePath. GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}

	process
	{
		$CalendarProcessing = Get-Content $CalendarSchedulingDelegationsJsonFilePath -Raw | ConvertFrom-Json
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json

        $count = $groupsToMigrate.Count
        $index = 0
		foreach ($groupToMigrate in $groupsToMigrate)
		{
            ++$index

			Write-Log "($index/$count) Processing Calendar Scheduling Permissions validation for group: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"

			[array]$emailAddresses = $groupToMigrate.EmailAddresses
			$emailAddresses += "X500:" + $groupToMigrate.LegacyExchangeDN

			$roomMailboxesNeedingRemediation = $CalendarProcessing | % {
				if ("X500:" + $_.BookInPolicy -in $emailAddresses `
					-or "X500:" + $_.RequestInPolicy -in $emailAddresses `
					-or "X500:" + $_.RequestOutOfPolicy -in $emailAddresses) { $_ } 
			} 

			if ($roomMailboxesNeedingRemediation)
			{
				$msg = "[ERROR] - Failed Mailbox Remediation Validation: '$($groupToMigrate.Name)' - '$($groupToMigrate.PrimarySmtpAddress)'"
				Write-Warning $msg
				Write-Log $msg

				$props = @{
						"Name" = $groupToMigrate.Name;
						"Identity" = $groupToMigrate.Identity;
						"PrimarySmtpAddress" = $groupToMigrate.PrimarySmtpAddress;
						"CalendarProcessing" = @($roomMailboxesNeedingRemediation);
					}

				$failedGroup = New-Object –TypeName PSObject –Prop $props
				$failedGroup # send the output on the return pipeline 
			}
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). CalendarSchedulingDelegationsJsonFilePath: $CalendarSchedulingDelegationsJsonFilePath.  GroupToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath"
	}
}

<#
.Synopsis
	Returns the groups that cannot be migrated because they are child groups of a group that in not on the migration list.
.DESCRIPTION
	Returns the groups that cannot be migrated because they are child groups of a group that in not on the migration list.
	If a group is not being migrated then exclude all the child groups as well as we cannot update the membership of the parent dirsycned group with the new cloud group.
.PARAMETER GroupsMembersJsonFilePath
	The file path of the json file containing all group membership information previously saved. 
.PARAMETER GroupsToExcludeJsonFilePath
	The file path of the json file containing to-be-excluded groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingNestingValidation -GroupsMembersJsonFilePath $Global:OnpremGroupMemberExportFileName -GroupsToExcludeJsonFilePath  $Global:OnpremGroupsExcludedFileName
#>
function Get-DistributionGroupsFailingNestingValidation
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing all group membership information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsMembersJsonFilePath,

		# The file path of the json file containing to-be-excluded groups information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToExcludeJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsMembersJsonFilePath: $GroupsMembersJsonFilePath. GroupsToExcludeJsonFilePath: $GroupsToExcludeJsonFilePath"
	}

	process
	{
		# If a group is not being migrated then exclude all the child groups as well as we cannot update the membership of the parent dirsycned group with the new cloud group
		$allGroupMembers = Get-Content $GroupsMembersJsonFilePath -Raw | ConvertFrom-Json
		$groupsToExclude = Get-Content $GroupsToExcludeJsonFilePath -Raw | ConvertFrom-Json

		$maxNestingCount = 10
		$currentGroupsToExclude = $groupsToExclude.Identity
		$nestedFailedGroups = @()
		for ($i = 1; $i -le $maxNestingCount; ++$i)
		{
			Write-Log "Processing to exclude child groups of any groups that have failed previous validations or excluded. Nesting level ($i / $maxNestingCount)..."
			$allChildGroupsOfFailedGroups = $allGroupMembers | Where { $_.Identity -in $currentGroupsToExclude } | Select -ExpandProperty Members | Where { $_.RecipientType -match "Group" }
			$currentGroupsToExclude = $allChildGroupsOfFailedGroups.Identity
			Write-Log "Child groups failing nesting validations $($currentGroupsToExclude.Count). Nesting level ($i / $maxNestingCount)..."
			$nestedFailedGroups += $allChildGroupsOfFailedGroups
		}

		$nestedFailedGroups | Sort Identity -Unique
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsMembersJsonFilePath: $GroupsMembersJsonFilePath. GroupsToExcludeJsonFilePath: $GroupsToExcludeJsonFilePath"
	}
}

#endregion Helper Functions for Validating that Groups are Good to Migrate

#region Shadow Group functions

<#
.Synopsis
	Returns the guid of the EXO recipient matching the specified Name.
.DESCRIPTION
	Returns the guid of the EXO recipient matching the specified Name.
	If the recipient name is in the list of $OriginalGroupNames, then the guid of the shadow group is returned.
	For performance optimisations, it caches the guids of the recipients already fetched in the global variable $Global:RecipientGuidCache
.PARAMETER RecipientName
	The Name of the recipient. 
.PARAMETER OriginalGroupNames
	The optional list of names of original groups to check if the guid of the shadow group is to be returned. 
.EXAMPLE
	Get-OnlineRecipientGuidFromName -RecipientName $RecipientName
	Get-OnlineRecipientGuidFromName -RecipientName $RecipientName -OriginalGroupNames $OriginalGroupNames
#>
function Get-OnlineRecipientGuidFromName()
{
	[CmdletBinding()]
	param
	(
		# The recipient name
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateNotNullOrEmpty()]
		[string]
		$RecipientName,

		# The recipient name
		[Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $false)]
		[string[]]
		$OriginalGroupNames
	)

	begin
	{
		# none
	}

	process
	{
		Write-Log "Getting Guid for recipient '$RecipientName'"

		# Escape any quotes
		$RecipientName = $RecipientName -replace "'", "''" 

		if ($RecipientName -in $OriginalGroupNames)
		{
			$RecipientName = Get-ShadowGroupName $RecipientName
		}

		if ($Global:RecipientGuidCache.ContainsKey($RecipientName))
		{
			$Global:RecipientGuidCache[$RecipientName]
			return
		}
		else
		{
			$recipientFilter = "{Name -eq '$RecipientName'}"
			$recipient = Get-Recipient -Filter $recipientFilter

			if ($recipient)
			{
				$guid = $recipient.Guid.ToString()
				$null = $Global:RecipientGuidCache.Add($RecipientName,$guid)
				$guid
			}
			else
			{
				Write-Log "[WARNING] - recipient not found. Recipient Name: '$RecipientName'"
			}
		}
	}

	end
	{
		# none
	}
}

<#
.Synopsis
	Creates the shadow group shell objects.
.DESCRIPTION
	Creates the shadow group shell objects. By default all groups will be created as pure distribution groups unless the name is in the list of groups to be migrated as security groups
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online groups information previously saved. 
.PARAMETER GroupsToMigrateAsSecurityGroupsJsonFilePath
	The file path of the json file containing online groups information to be migrated as security groups previously saved. 
.EXAMPLE
	Create-ShadowGroupShellObjects -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -GroupsToMigrateAsSecurityGroupsJsonFilePath $Global:OnlineGroupsGoodToMigrateAsSecurityGroupsFileName
#>
function Create-ShadowGroupShellObjects
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath,

		# The file path of the json file containing online groups information to be migrated as security groups previously saved
		[Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		$GroupsToMigrateAsSecurityGroupsJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. GroupsToMigrateAsSecurityGroupsJsonFilePath: $GroupsToMigrateAsSecurityGroupsJsonFilePath"
	}

	process
	{
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json
		$groupsToMigrateAsSecurityGroups = Get-Content $GroupsToMigrateAsSecurityGroupsJsonFilePath -Raw | ConvertFrom-Json

		[array]$AllSecurityGroupNames = $groupsToMigrateAsSecurityGroups | % { Get-ShadowGroupName $_.Name }
	
		$scriptBlock = {
			$group = $Input # InputObject from Start-RobustCloudCommand

			$Name =  Get-ShadowGroupName $group.Name
			$AllGroupNames = $groupsToMigrate.Name
			$Alias = Get-ShadowGroupAlias $group.Alias
			$ArbitrationMailbox = $group.ArbitrationMailbox
			if ($ArbitrationMailbox.StartsWith("SystemMailbox"))
			{
				$ArbitrationMailbox = $null
			}

			$BypassNestedModerationEnabled = $group.BypassNestedModerationEnabled
			$CopyOwnerToMember = $false
			$DisplayName = Get-ShadowGroupName $group.DisplayName
			$IgnoreNamingPolicy = $true
			# TODO: Updated this in the second pass in case as it may be a group
			$ManagedBy = $group.ManagedBy | % { if ($_ -ne $null -and $_ -ne "Organization Management") { $guid = Get-OnlineRecipientGuidFromName $_ $AllGroupNames; if ($guid -ne $null) { $guid } else { $DefaultGroupOwner } } else { $DefaultGroupOwner } }
			$MemberDepartRestriction = $group.MemberDepartRestriction
			$MemberJoinRestriction = $group.MemberJoinRestriction
			# TODO: Updated this in the second pass in case 
			$ModeratedBy = $group.ModeratedBy | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $AllGroupNames} }
			$ModerationEnabled  = $group.ModerationEnabled 
			$Notes  = $group.Notes 
			$PrimarySmtpAddress = Get-ShadowGroupPrimarySmtpAddress $group.PrimarySmtpAddress
			$RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
			$RoomList = $group.RecipientTypeDetails -eq "RoomList"
			#$SamAccountName
			$SendModerationNotifications = $group.SendModerationNotifications
			$Type = "Distribution" # $group.RecipientType - We'll always create pure Distribution groups as per the design decision

			if ($Name -in $AllSecurityGroupNames)
			{
				Write-Log "DL: '$($Name)' - '$($PrimarySmtpAddress)' will be created as a MESG."
				$Type = "Security"
			}

			#TODO - Populate any addtional attributes such as extension attributes

			#$newDL = Get-DistributionGroup -Identity $PrimarySmtpAddress -ErrorAction Ignore
			$newDL = Get-Recipient -Identity $PrimarySmtpAddress -ErrorAction Ignore
			if ($newDL)
			{
				if ($newDL.Name -eq $Name -and $newDL.RecipientType -match "Group")
				{
					$msg = "[WARNING] - DL: '$($Name)' - '$($PrimarySmtpAddress)' already exists."
					Write-Warning $msg
					Write-Log $msg
				}
				else
				{
					$dlName = $newDL.Name
					$msg = "[WARNING] - PrimarySmtpAddress Conflict: '$($dlName)' and '$($Name)' - '$($PrimarySmtpAddress)'."
					Write-Warning $msg
					Write-Log $msg

					$PrimarySmtpAddress =  $PrimarySmtpAddress.Split('@')[0] + $Global:NewGroupNameSuffix + "@" + $PrimarySmtpAddress.Split('@')[1]

					$msg = "[WARNING] - Trying New PrimarySmtpAddress: '$($Name)' - '$($PrimarySmtpAddress)'."
					Write-Warning $msg
					Write-Log $msg

					$newDL = Get-DistributionGroup -Identity $PrimarySmtpAddress -ErrorAction Ignore
				}
			}

			if ($newDL -eq $null)
			{
				Write-Log "Creating DL: '$($Name)' - '$($PrimarySmtpAddress)'..."

				$newDL = New-DistributionGroup -Name $Name `
									  -Alias $Alias `
									  -BypassNestedModerationEnabled:$BypassNestedModerationEnabled `
									  -DisplayName $DisplayName `
									  -IgnoreNamingPolicy:$true `
									  -ManagedBy $ManagedBy `
									  -MemberDepartRestriction $MemberDepartRestriction `
									  -MemberJoinRestriction $MemberJoinRestriction `
									  -ModeratedBy $ModeratedBy `
									  -ModerationEnabled:$ModerationEnabled `
									  -Notes $Notes `
									  -PrimarySmtpAddress $PrimarySmtpAddress `
									  -RequireSenderAuthenticationEnabled:$RequireSenderAuthenticationEnabled `
									  -RoomList:$RoomList `
									  -SendModerationNotifications $SendModerationNotifications `
									  -Type $Type
			}

			Set-DistributionGroup -Identity $newDL.Guid.ToString() -HiddenFromAddressListsEnabled:$true
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $groupsToMigrate -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. GroupsToMigrateAsSecurityGroupsJsonFilePath: $GroupsToMigrateAsSecurityGroupsJsonFilePath"
	}
}

<#
.Synopsis
	Sets the delivery restrictions and public delegates on the shadow groups as well as any other recipient objects.
.DESCRIPTION
	Sets the delivery restrictions and public delegates on the shadow groups as well as any other recipient objects.
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online groups information previously saved. 
.PARAMETER RecipientDeliveryRestrictionsJsonFilePath
	The file path of the json file containing recipient delivery restrictions and public delegates information previously saved. 
.EXAMPLE
	Set-RecipientDeliveryRestrictionsAndPublicDelegates -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -RecipientDeliveryRestrictionsJsonFilePath $Global:OnlineDeliveryRestrictionsExportFileName
#>
function Set-RecipientDeliveryRestrictionsAndPublicDelegates
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath,

		# The file path of the json file containing recipient delivery restrictions and public delegates information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$RecipientDeliveryRestrictionsJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. RecipientDeliveryRestrictionsJsonFilePath: $RecipientDeliveryRestrictionsJsonFilePath"
	}

	process
	{
		[array]$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json
		[array]$recipientDeliveryRestrictions = Get-Content $RecipientDeliveryRestrictionsJsonFilePath -Raw | ConvertFrom-Json
		[array]$recipientDeliveryRestrictions = $recipientDeliveryRestrictions | Sort Name -Unique
		[array]$OriginalGroupNames = $groupsToMigrate.Name
		[array]$nonShadowGroupsrecipientDeliveryRestrictions = $recipientDeliveryRestrictions | Where { $_.Name -notin $OriginalGroupNames }

		$scriptBlock = {
			$group = $Input # InputObject from Start-RobustCloudCommand

			$AcceptMessagesOnlyFrom = $group.AcceptMessagesOnlyFrom | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ } }
			$AcceptMessagesOnlyFromDLMembers =  $group.AcceptMessagesOnlyFromDLMembers | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }
			$RejectMessagesFrom = $group.RejectMessagesFrom | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ } }
			$RejectMessagesFromDLMembers = $group.RejectMessagesFromDLMembers | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }
			$GrantSendOnBehalfTo = $group.GrantSendOnBehalfTo | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }

			if ($AcceptMessagesOnlyFrom -or $AcceptMessagesOnlyFromDLMembers -or $RejectMessagesFrom -or $RejectMessagesFromDLMembers -or $GrantSendOnBehalfTo)
			{
				$Name =  Get-ShadowGroupName $group.Name
				$PrimarySmtpAddress = Get-ShadowGroupPrimarySmtpAddress $group.PrimarySmtpAddress
				$RecipientType = $group.RecipientType
				$guid = Get-OnlineRecipientGuidFromName $Name $OriginalGroupNames

				if ($AcceptMessagesOnlyFrom)
				{
					Write-Log "Setting AcceptMessagesOnlyFrom for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					Set-DistributionGroup -Identity $guid -AcceptMessagesOnlyFrom $AcceptMessagesOnlyFrom
				}

				if ($AcceptMessagesOnlyFromDLMembers)
				{
					Write-Log "Setting AcceptMessagesOnlyFromDLMembers for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					Set-DistributionGroup -Identity $guid -AcceptMessagesOnlyFromDLMembers $AcceptMessagesOnlyFromDLMembers
				}

				if ($RejectMessagesFrom)
				{
					Write-Log "Setting RejectMessagesFrom for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					Set-DistributionGroup -Identity $guid -RejectMessagesFrom $RejectMessagesFrom
				}

				if ($RejectMessagesFromDLMembers)
				{
					Write-Log "Setting RejectMessagesFromDLMembers for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					Set-DistributionGroup -Identity $guid -RejectMessagesFromDLMembers $RejectMessagesFromDLMembers
				}

				if ($GrantSendOnBehalfTo)
				{
					Write-Log "Setting GrantSendOnBehalfTo for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					Set-DistributionGroup -Identity $guid -GrantSendOnBehalfTo $GrantSendOnBehalfTo
				}
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $groupsToMigrate -IdentifyingProperty "Name" -ScriptBlock $scriptBlock

		# Update delivery restrictions on any cloud managed objects.
		# Let it report error if we try to do this on a dirsynced DL

		if ($nonShadowGroupsrecipientDeliveryRestrictions -eq $null)
		{
			return
		}

		$scriptBlock = {
			param ([hashtable]$groupsToMigrate)

			$recipient = $Input # InputObject from Start-RobustCloudCommand
			[array]$groupsToMigrateNames = $groupsToMigrate.Name

			if ($recipient.Name -in $groupsToMigrateNames)
			{
				Write-Log "Skipping already processed shadow group '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
				return
			}

			$AcceptMessagesOnlyFromDLMembers = $recipient.AcceptMessagesOnlyFromDLMembers | Where { $_ -in $groupsToMigrateNames } | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }
			$RejectMessagesFromDLMembers = $recipient.RejectMessagesFromDLMembers | Where { $_ -in $groupsToMigrateNames } | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }
			$GrantSendOnBehalfTo = $recipient.GrantSendOnBehalfTo | Where { $_ -in $groupsToMigrateNames } | % { if ($_ -ne $null) { Get-OnlineRecipientGuidFromName $_ $OriginalGroupNames} }

			if ($AcceptMessagesOnlyFromDLMembers -or $RejectMessagesFromDLMembers -or $GrantSendOnBehalfTo)
			{
				$Name =  $recipient.Name
				$PrimarySmtpAddress = $recipient.PrimarySmtpAddress
				$RecipientType = $recipient.RecipientType
				$guid = Get-OnlineRecipientGuidFromName $Name

				if ($AcceptMessagesOnlyFromDLMembers)
				{
					Write-Log "Setting AcceptMessagesOnlyFromDLMembers for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					switch ($RecipientType)
					{
						"UserMailbox" { Set-Mailbox -Identity $guid -AcceptMessagesOnlyFromDLMembers @{ Add = $AcceptMessagesOnlyFromDLMembers }; break }
						"MailUser" { Set-MailUser -Identity $guid -AcceptMessagesOnlyFromDLMembers @{ Add = $AcceptMessagesOnlyFromDLMembers }; break }
						"MailContact" { Set-MailContact -Identity $guid -AcceptMessagesOnlyFromDLMembers @{ Add = $AcceptMessagesOnlyFromDLMembers }; break }
						default {Set-DistributionGroup -Identity $guid -AcceptMessagesOnlyFromDLMembers @{ Add = $AcceptMessagesOnlyFromDLMembers }; break }
					}
				}

				if ($RejectMessagesFromDLMembers)
				{
					Write-Log "Setting RejectMessagesFromDLMembers for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					switch ($RecipientType)
					{
						"UserMailbox" { Set-Mailbox -Identity $guid -RejectMessagesFromDLMembers @{ Add = $RejectMessagesFromDLMembers }; break }
						"MailUser" { Set-MailUser -Identity $guid -RejectMessagesFromDLMembers @{ Add = $RejectMessagesFromDLMembers }; break }
						"MailContact" { Set-MailContact -Identity $guid -RejectMessagesFromDLMembers @{ Add = $RejectMessagesFromDLMembers }; break }
						default {Set-DistributionGroup -Identity $guid -RejectMessagesFromDLMembers @{ Add = $RejectMessagesFromDLMembers }; break }
					}
				}

				if ($GrantSendOnBehalfTo)
				{
					Write-Log "Setting GrantSendOnBehalfTo for '$RecipientType': '$($Name)' - '$($PrimarySmtpAddress)'..."
					switch ($RecipientType)
					{
						"UserMailbox" { Set-Mailbox -Identity $guid -GrantSendOnBehalfTo @{ Add = $GrantSendOnBehalfTo }; break }
						"MailUser" { Set-MailUser -Identity $guid -GrantSendOnBehalfTo @{ Add = $GrantSendOnBehalfTo }; break }
						"MailContact" { Set-MailContact -Identity $guid -GrantSendOnBehalfTo @{ Add = $GrantSendOnBehalfTo }; break }
						default {Set-DistributionGroup -Identity $guid -GrantSendOnBehalfTo @{ Add = $GrantSendOnBehalfTo }; break }
					}
				}
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $nonShadowGroupsrecipientDeliveryRestrictions -IdentifyingProperty "Name" -ScriptBlock $scriptBlock -ArgumentList @{Name=$groupsToMigrate.Name}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. RecipientDeliveryRestrictionsJsonFilePath: $RecipientDeliveryRestrictionsJsonFilePath"
	}
}

<#
.Synopsis
	Adds mailbox FullAccess permissions based on the shadow groups.
.DESCRIPTION
	Adds mailbox FullAccess permissions based on the shadow groups.
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online groups information previously saved. 
.PARAMETER MailboxFullAccessPermissionsJsonFilePath
	The file path of the json file containing mailbox FullAccess permissions information previously saved. 
.EXAMPLE
	Add-MailboxFullAccessPermissions -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -MailboxFullAccessPermissionsJsonFilePath $Global:OnlineFullAccessPermissionsExportFileName
#>
function Add-MailboxFullAccessPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath,

		# The file path of the json file containing mailbox FullAccess permissions information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$MailboxFullAccessPermissionsJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. MailboxFullAccessPermissionsJsonFilePath: $MailboxFullAccessPermissionsJsonFilePath"
	}

	process
	{
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json
		$fullAccessPermissions = Get-Content $MailboxFullAccessPermissionsJsonFilePath -Raw | ConvertFrom-Json
		[array]$groupsToMigrateNames = $groupsToMigrate.Name
		[array]$fullAccessPermissions = $fullAccessPermissions | Where {$_.User -in $groupsToMigrateNames}
		[array]$OriginalGroupNames = $groupsToMigrate.Name

		if ($fullAccessPermissions -eq $null)
		{
			return
		}

		$scriptBlock = {
			param ([hashtable]$groupsToMigrate)

			$fullAccessPermission = $Input # InputObject from Start-RobustCloudCommand
			[array]$groupsToMigrateNames = $groupsToMigrate.Name

			if ($fullAccessPermission.User -in $groupsToMigrateNames)
			{
				Write-Log "Adding FullAccess permission on mailbox '$($fullAccessPermission.Identity)' for: '$($fullAccessPermission.User)'..."

				$identity =  $fullAccessPermission.Identity
				$guid =  Get-OnlineRecipientGuidFromName $identity $OriginalGroupNames 
				$user = Get-ShadowGroupName $fullAccessPermission.User
				$null = Add-MailboxPermission -Identity $guid -User $user -AccessRights "FullAccess" -InheritanceType "All"
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $fullAccessPermissions -IdentifyingProperty "Identity" -ScriptBlock $scriptBlock -ArgumentList @{Name=$groupsToMigrate.Name}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. MailboxFullAccessPermissionsJsonFilePath: $MailboxFullAccessPermissionsJsonFilePath"
	}
}

<#
.Synopsis
	Adds recipient SendAs permissions based on the shadow groups.
.DESCRIPTION
	Adds recipient SendAs permissions based on the shadow groups.
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online groups information previously saved. 
.PARAMETER RecipientSendAsPermissionsJsonFilePath
	The file path of the json file containing recipient SendAs permissions information previously saved. 
.EXAMPLE
	Add-RecipientSendAsPermissions -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -RecipientSendAsPermissionsJsonFilePath $Global:OnlineSendAsPermissionsExportFileName
#>
function Add-RecipientSendAsPermissions
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath,

		# The file path of the json file containing recipient SendAs information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$RecipientSendAsPermissionsJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. RecipientSendAsPermissionsJsonFilePath: $RecipientSendAsPermissionsJsonFilePath"
	}

	process
	{
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json
		$sendAsPermissions = Get-Content $RecipientSendAsPermissionsJsonFilePath -Raw | ConvertFrom-Json
		[array]$groupsToMigrateNames = $groupsToMigrate.Name
		[array]$sendAsPermissions = $sendAsPermissions | Where {$_.Trustee -in $groupsToMigrateNames}
		[array]$OriginalGroupNames = $groupsToMigrate.Name

		if ($sendAsPermissions -eq $null)
		{
			return
		}

		$scriptBlock = {
			param ([hashtable]$groupsToMigrate)

			$sendAsPermission = $Input # InputObject from Start-RobustCloudCommand
			[array]$groupsToMigrateNames = $groupsToMigrate.Name

			if ($sendAsPermission.Trustee -in $groupsToMigrateNames)
			{
				Write-Log "Adding SendAs permission on recipient '$($sendAsPermission.Identity)' to: '$($sendAsPermission.Trustee)'..."

				$identity =  $sendAsPermission.Identity
				$guid =  Get-OnlineRecipientGuidFromName $identity $OriginalGroupNames 
				$trustee = Get-ShadowGroupName $sendAsPermission.Trustee
				$null = Add-RecipientPermission -Identity $guid -Trustee $trustee -AccessRights "SendAs" -Confirm:$false
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $sendAsPermissions -IdentifyingProperty "Identity" -ScriptBlock $scriptBlock -ArgumentList @{Name=$groupsToMigrate.Name}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. RecipientSendAsPermissionsJsonFilePath: $RecipientSendAsPermissionsJsonFilePath"
	}
}

<#
.Synopsis
	Populates shadow group membership.
.DESCRIPTION
	Populates shadow group membership.
.PARAMETER GroupsToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online groups information previously saved. 
.PARAMETER RecipientSendAsPermissionsJsonFilePath
	The file path of the json file containing group membership information previously saved. 
.EXAMPLE
	Add-ShadowGroupMembers -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -GroupMembersJsonFilePath $Global:OnlineGroupMemberExportFileName
#>
function Add-ShadowGroupMembers
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online groups information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToMigrateJsonFilePath,

		# The file path of the json file containing group membership information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupMembersJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. GroupMembersJsonFilePath: $GroupMembersJsonFilePath"
	}

	process
	{
		$groupsToMigrate = Get-Content $GroupsToMigrateJsonFilePath -Raw | ConvertFrom-Json
		$groupsMembership = Get-Content $GroupMembersJsonFilePath -Raw | ConvertFrom-Json
		[array]$groupsToMigrateNames = $groupsToMigrate.Name
		$shadowGroupsMembership = $groupsMembership | Where { $_.Name -in $groupsToMigrateNames }
		$nonShadowGroupsMembership = $groupsMembership | Where { $_.Name -notin $groupsToMigrateNames } | % {
			$nonShadowGroupMembership = $_
			$_.Members.Name | % { if ($_ -in $groupsToMigrateNames) { $nonShadowGroupMembership } }
		} | Sort Name -Unique

		$scriptBlock = {
			$group = $Input # InputObject from Start-RobustCloudCommand
			$Name =  Get-ShadowGroupName $group.Name
			$PrimarySmtpAddress = Get-ShadowGroupPrimarySmtpAddress $group.PrimarySmtpAddress
			$guid = Get-OnlineRecipientGuidFromName $Name

			Write-Log "Processing membership for DL: '$($Name)' - '$($PrimarySmtpAddress)'"

			$members = $group.Members | % {
				if ($_ -ne $null)
				{
					$memberName = $_.Name
					Get-OnlineRecipientGuidFromName $memberName $groupsToMigrateNames
				}
			}

			Update-DistributionGroupMember -Identity $guid -Members $members -Confirm:$false -BypassSecurityGroupManagerCheck:$true
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $shadowGroupsMembership -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 

		if ($nonShadowGroupsMembership -eq $null)
		{
			return
		}
	
		$scriptBlock = {
			param ([hashtable]$groupsToMigrate)

			$group = $Input # InputObject from Start-RobustCloudCommand
			[array]$groupsToMigrateNames = $groupsToMigrate.Name
			$Name =  $group.Name
			$PrimarySmtpAddress = $group.PrimarySmtpAddress
			$guid = Get-OnlineRecipientGuidFromName $Name

			Write-Log "Processing membership for DL: '$($Name)' - '$($PrimarySmtpAddress)'"

			$group.Members | Where {$_.Name -in $groupsToMigrateNames} | % {
				if ($_ -ne $null)
				{
					$memberName = Get-ShadowGroupName $_.Name
					Write-Log "Adding member '$memberName' to DL: '$Name' - '$PrimarySmtpAddress'"
					$member =  Get-OnlineRecipientGuidFromName $memberName
				
					try
					{
						Add-DistributionGroupMember -Identity $guid -Member $member -Confirm:$false -BypassSecurityGroupManagerCheck:$true
					}
					catch
					{
						if ($_.CategoryInfo.Reason -match "MemberAlreadyExistsException")
						{
							Write-Log "[INFO] - $_"
							$Error.Clear()
							$Global:Error.Clear()
						}
						elseif ($_.CategoryInfo.Reason -match "OperationRequiresGroupManagerException")
						{
							Write-Log "[ERROR] - $_"
							$Error.Clear()
							$Global:Error.Clear()
						}
						else
						{
							Write-Log "[ERROR] - $_"
						}
					}
				}
			}
		}

		Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $nonShadowGroupsMembership -IdentifyingProperty "Name" -ScriptBlock $scriptBlock -ArgumentList @{Name=$groupsToMigrate.Name}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToMigrateJsonFilePath: $GroupsToMigrateJsonFilePath. GroupMembersJsonFilePath: $GroupMembersJsonFilePath"
	}
}

#endregion Shadow Group functions

#region Shadow Group Animation functions

<#
.Synopsis
	Sets the onprem group object out of AAD Connect sync scope.
.DESCRIPTION
	Sets the onprem group object out of AAD Connect sync scope by setting extensionAttribute1 = DoNotSync .
.PARAMETER Group
	The json of the onprem group object. 
.EXAMPLE
	Set-OnpremGroupOutOfSyncScope -Group $OnpremGroup
#>
function Set-OnpremGroupOutOfSyncScope
{
	[CmdletBinding()]
	param
	(
		# The json of the onprem group object
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		$Group
	)

	begin
	{
		# none
	}

	process
	{
		Write-Log "Setting the on-prem DL to DoNotSync: '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)'"

		$PrimarySmtpAddress = $Group.PrimarySmtpAddress -replace "'", "''" 
		$filter = "{ PrimarySmtpAddress -eq '$PrimarySmtpAddress' }"
		$onpremGroup = Get-OnpremDistributionGroup -Filter $filter -IgnoreDefaultScope
		if ($onpremGroup)
		{
			$distinguishedName = $onpremGroup.DistinguishedName
			$domainController = Get-DomainController $distinguishedName
			Set-OnpremDistributionGroup -DomainController $domainController -Identity $distinguishedName -CustomAttribute1 "DoNotSync" -Confirm:$false -ForceUpgrade
		}
		else
		{
			$msg = "[ERROR] - On-prem DL '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)' not found!!"
			Write-Warning $msg
			Write-Log $msg
		}
	}

	end
	{
		# none
	}
}

<#
.Synopsis
	Deletes the online dirsynced group object from Azure AD.
.DESCRIPTION
	Deletes the online dirsynced group object from Azure AD based on its ExternalDirectoryObjectId.
.PARAMETER Group
	The json of the online group object. 
.EXAMPLE
	Remove-SyncedGroupFromAzureAD -Group $OnlineGroup
#>
function Remove-SyncedGroupFromAzureAD
{
	[CmdletBinding()]
	param
	(
		# The json of the online group object
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateNotNull()]
		$Group
	)

	begin
	{
		# none
	}

	process
	{
		if ($Group.IsDirSynced)
		{
			Write-Log "Removing group '$($Group.ExternalDirectoryObjectId)' from Azure: '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)'"

			Remove-MsolGroup -ObjectId $Group.ExternalDirectoryObjectId -Force -ErrorAction Ignore
		}
		else
		{
			Write-Log "[ERROR] - Unexpected attempt to remove non-dirsynced group '$($Group.ExternalDirectoryObjectId)' from Azure: '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)'"
		}
	}

	end
	{
		# none
	}
}

<#
.Synopsis
	Effects the cut-over from original dirsynced group to shadow groups.
.DESCRIPTION
	Effects the cut-over from original dirsynced group to shadow groups.
	It first descopes a batch of onprem group from AAD Connect Sync and then deletes that batch of groups the Azure AD.
	Finally it removed the shadow group prefixes for attributes and copies over all proxyAddresses and legacyExchangeDN of the original group.
.PARAMETER OnlineGroupsGoodToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate online group information previously saved. 
.PARAMETER OnpremGroupsGoodToMigrateJsonFilePath
	The file path of the json file containing good-to-migrate onprem group information previously saved. 
.EXAMPLE
	Switch-ShadowGroups -OnlineGroupsGoodToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -OnpremGroupsGoodToMigrateJsonFilePath $Global:OnpremGroupsGoodToMigrateFileName
#>
function Switch-ShadowGroups
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-migrate online group information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$OnlineGroupsGoodToMigrateJsonFilePath,

		# The file path of the json file containing good-to-migrate onprem group information previously saved
		[Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$OnpremGroupsGoodToMigrateJsonFilePath
	)

	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). OnlineGroupsGoodToMigrateJsonFilePath: $OnlineGroupsGoodToMigrateJsonFilePath. OnpremGroupsGoodToMigrateJsonFilePath: $OnpremGroupsGoodToMigrateJsonFilePath"
	}
		
	process
	{
		$onlineGroups = Get-Content $OnlineGroupsGoodToMigrateJsonFilePath -Raw | ConvertFrom-Json
		$onpremGroups = Get-Content $OnpremGroupsGoodToMigrateJsonFilePath -Raw | ConvertFrom-Json

		$index = 0
		$batchSize = 50
		$count = $onlineGroups.Count
		for ($loopIndex = 0; $loopIndex -le $count/$batchSize; ++$loopIndex)
		{
			$skip = $loopIndex*$batchSize

			Write-Log "Processing loop: $loopIndex skip: $skip"

			[array]$onlineGroupsBatch = $onlineGroups | Select -Skip $skip -First $batchSize
			[array]$onlineGroupsBatchNames = $onlineGroupsBatch.Name
			[array]$onlineGroupsBatchPrimarySmtpAddresses = $onlineGroupsBatch.PrimarySmtpAddress
	
			New-OnpremExchangeSession
			$onpremGroups | Where { $_.Name -in $onlineGroupsBatchNames -or $_.PrimarySmtpAddress -in $onlineGroupsBatchPrimarySmtpAddresses} | % { if ($_ -ne $null) { Set-OnpremGroupOutOfSyncScope $_ } }

			$onlineGroupsBatch = $onlineGroupsBatch | Sort Name
			$onlineGroupsBatch | % { Remove-SyncedGroupFromAzureAD -Group $_ }

			$scriptBlock = {
				param($OnpremData)
				$onlineGroup = $Input # InputObject from Start-RobustCloudCommand
				$onpremGroups = $OnpremData.Groups

				$Name = $onlineGroup.Name

				$onpremGroup = $onpremGroups | where { $_.Name -eq $Name }

				$Alias = $onlineGroup.Alias
				$DisplayName = $onlineGroup.DisplayName
				$emailAddresses = $onlineGroup.EmailAddresses
				$emailAddresses += "x500:" + $onlineGroup.LegacyExchangeDN
				if ($onpremGroup.LegacyExchangeDN) { $emailAddresses += "x500:" + $onpremGroup.LegacyExchangeDN }

				$emailAddresses = $emailAddresses | Sort -Unique #case insensitive unique

				$shadowName = Get-ShadowGroupName $onlineGroup.Name
				$shadowPrimarySmtpAddress = Get-ShadowGroupPrimarySmtpAddress $onlineGroup.PrimarySmtpAddress

				Write-Log "Starting animating shadow DL: '$shadowName' - '$shadowPrimarySmtpAddress'"
		
				Remove-SyncedGroupFromAzureAD -Group $onlineGroup

				Write-Log "Checking in Exchange Online until the master object '$($onlineGroup.Name)' - '$($onlineGroup.PrimarySmtpAddress)' is gone..."
				$secondRemaining = 600
				$sleep = 2
				do
				{
					$ExternalDirectoryObjectId = $onlineGroup.ExternalDirectoryObjectId
					$filter = "{ ExternalDirectoryObjectId -eq '$ExternalDirectoryObjectId' }"
					$test = Get-DistributionGroup -Filter $filter -ResultSize 10 -ErrorAction Ignore
					if ($test)
					{
						Start-Sleep -Seconds $sleep
						$secondRemaining -= $sleep
						Write-Log "Waiting for master object '$($onlineGroup.Name)' - '$($onlineGroup.PrimarySmtpAddress)' to disappear from EXO... $secondRemaining"
					}
				}
				while ($test -and $secondRemaining -gt 0)

				Write-Log "The master object '$($onlineGroup.Name)' - '$($onlineGroup.PrimarySmtpAddress)' is gone..."

				$shadowGuid = Get-OnlineRecipientGuidFromName $shadowName

				if ($shadowGuid -ne $null)
				{
					Write-Log "Animating shadow DL '$shadowGuid' - '$shadowName' - '$shadowPrimarySmtpAddress'..."

					Set-DistributionGroup -Identity $shadowGuid `
						-Alias $Alias `
						-DisplayName $DisplayName `
						-Name $Name `
						-EmailAddresses  $emailAddresses `
						-HiddenFromAddressListsEnabled $onlineGroup.HiddenFromAddressListsEnabled `
						-IgnoreNamingPolicy:$true

					Write-Log "Restoring the PrimarySmtpAddress as per the new standard for animated DL '$shadowGuid' - '$Name' - '$shadowPrimarySmtpAddress'..."
					Set-DistributionGroup -Identity $shadowGuid -PrimarySmtpAddress $shadowPrimarySmtpAddress
				}
				else
				{
					Write-Log "[INFO] Seems already Animated shadow DL '$shadowGuid' - '$shadowName' - '$shadowPrimarySmtpAddress'..."
				}
			}

			Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $onlineGroupsBatch -IdentifyingProperty "Name" -ScriptBlock $scriptBlock -ArgumentList @{Groups=$onpremGroups}
		}
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). OnlineGroupsGoodToMigrateJsonFilePath: $OnlineGroupsGoodToMigrateJsonFilePath. OnpremGroupsGoodToMigrateJsonFilePath: $OnpremGroupsGoodToMigrateJsonFilePath"
	}
}

#endregion Shadow Group Animation functions

#region Migrated Onprem Group Deprovisioning functions

<#
.Synopsis
	Returns the groups that cannot be disabled because they are child groups of a group that is not on the list of groups to be disabled.
.DESCRIPTION
	Returns the groups that cannot be disabled because they are child groups of a group that is not on the list of groups to be disabled.
	If a group is to be disabled (and converted into a contact), it should not be part of any other group that we are not disabling.
	Otherwise we'll need to make the contact for disabled group a member of those groups that we are not disabling.
.PARAMETER GroupsMembersJsonFilePath
	The file path of the json file containing all group membership information previously saved. 
.PARAMETER GroupsToDisableJsonFilePath
	The file path of the json file containing to-be-disabled groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingNestingValidationForDisablement -GroupsMembersJsonFilePath $Global:OnpremGroupMemberExportFileName -GroupsToDisableJsonFilePath $Global:OnpremGroupsGoodToDisableFileName
#>
function Get-DistributionGroupsFailingNestingValidationForDisablement
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing all group membership information previously saved
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsMembersJsonFilePath,

		# The file path of the json file containing to-be-disable groups information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToDisableJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsMembersJsonFilePath: $GroupsMembersJsonFilePath.  GroupsToDisableJsonFilePath: $GroupsToDisableJsonFilePath"
	}

	process
	{
		# If a group is to be disabled (and converted into a contact), it should not be part of any other group that we are not disabling
		# Otherwise we'll need to make the contact for disabled group member of those groups that we are not disabling.
		$allGroupMembers = Get-Content $GroupsMembersJsonFilePath -Raw | ConvertFrom-Json
		$groupsToDisable = Get-Content $GroupsToDisableJsonFilePath -Raw | ConvertFrom-Json

		$maxNestingCount = 10
		$currentGroupsToDisable = $groupsToDisable
		$nestedFailedGroups = @()
		for ($i = 1; $i -le $maxNestingCount; ++$i)
		{
			Write-Log "Processing to exclude child groups of any groups we are not disabling. Nesting level ($i / $maxNestingCount)..."

			$dropFile = ($Global:OnpremGroupsGoodToDisableFileName + ".tmp.$i")
			if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

			$index = 0
			$count = $currentGroupsToDisable.Count
			$nestedFailedGroups += $currentGroupsToDisable | % {
				$groupToDisable = $_
				++$index

				Write-Log "($index/$count/$i) Checking the group for membership of any other group: '$($groupToDisable.Name)' - '$($groupToDisable.PrimarySmtpAddress)'"

				# All groups where the group to be migrated is a member
				$allNestingGroups = $AllGroupMembers | Where { $_.Members.PrimarySmtpAddress -eq  $groupToDisable.PrimarySmtpAddress }

				# All nesting groups that are not in the migrations list
				$nonMigratingNestingGroups = $allNestingGroups |  Where { $_.PrimarySmtpAddress -notin  $currentGroupsToDisable.PrimarySmtpAddress }

				if ($nonMigratingNestingGroups)
				{
					Write-Warning "Failed Nesting Validation: '$($groupToDisable.Name)' - '$($groupToDisable.PrimarySmtpAddress)'"

					$props = @{
							"Name" = $groupToDisable.Name;
							"Identity" = $groupToDisable.Identity;
							"PrimarySmtpAddress" = $groupToDisable.PrimarySmtpAddress;
							"MemberOf" = @($nonMigratingNestingGroups);
						}

					$failedGroup = New-Object –TypeName PSObject –Prop $props
					$failedGroup # send the output on the return pipeline 
				}
			}

			# Update the current list by removing groups that are already failed
            $currentGroupsToDisableOld = $currentGroupsToDisable
			$currentGroupsToDisable = $currentGroupsToDisable | Where { $_.PrimarySmtpAddress -notin $nestedFailedGroups.PrimarySmtpAddress }
            $currentGroupsToDisable | ConvertTo-Json | Out-File $dropFile
            if ($currentGroupsToDisable.Count -eq $currentGroupsToDisableOld.Count)
            {
                # all done - list not changing
                break
            }
		}

		$nestedFailedGroups | Sort Identity -Unique
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsMembersJsonFilePath: $GroupsMembersJsonFilePath.  GroupsToDisableJsonFilePath: $GroupsToDisableJsonFilePath"
	}
}

<#
.Synopsis
	Returns the groups that cannot be deleted because they are child groups of a group that is not on the list of groups to be disabled.
.DESCRIPTION
	Returns the groups that cannot be deleted because they are child groups of a group that is not on the list of groups to be deleted.
	If a group is to be deleted, it should not be part of any other group that we are not deleting.
.PARAMETER GroupsMembersJsonFilePath
	The file path of the json file containing all group membership information previously saved. 
.PARAMETER GroupsToDisableJsonFilePath
	The file path of the json file containing to-be-deleted groups information previously saved.
.EXAMPLE
	Get-DistributionGroupsFailingNestingValidationForDeletion -GroupsToDeleteJsonFilePath $Global:OnpremGroupsGoodToDisableFileName
#>
function Get-DistributionGroupsFailingNestingValidationForDeletion
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing to-be-disable groups information previously saved
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$GroupsToDeleteJsonFilePath
	)
	
	begin
	{
		Write-Log "Begin executing $($MyInvocation.MyCommand). GroupsToDeleteJsonFilePath: $GroupsToDeleteJsonFilePath"
	}

	process
	{
		# If a group is to be deleted, it should not be part of any other group that we are not deleting
		$groupsToDelete = Get-Content $GroupsToDeleteJsonFilePath -Raw | ConvertFrom-Json

		$maxNestingCount = 10
		$currentGroupsToDelete = $groupsToDelete
		$nestedFailedGroups = @()
		for ($i = 1; $i -le $maxNestingCount; ++$i)
		{
			Write-Log "Processing to exclude child groups of any groups we are not disabling. Nesting level ($i / $maxNestingCount)..."

			$dropFile = ($Global:OnpremGroupsGoodToDeleteFileName + ".tmp.$i")
			if (Test-Path $dropFile) { Remove-Item $dropFile -Force -Confirm:$false }

			$index = 0
			$count = $currentGroupsToDelete.Count
			$nestedFailedGroups += $currentGroupsToDelete | % {
				$groupToDelete = $_
				++$index

				Write-Log "($index/$count/$i) Checking the group for membership of any other group: '$($groupToDelete.DistinguishedName)'"

				# All groups where the group to be deleted is a member
				$allNestingGroups = $groupToDelete.MemberOf

				# All nesting groups that are not in the migrations list
				$nonMigratingNestingGroups = $allNestingGroups |  Where { $_.DistinguishedName -notin  $currentGroupsToDelete.DistinguishedName }

				if ($nonMigratingNestingGroups)
				{
					Write-Warning "Failed Nesting Validation: '$($groupToDelete.DistinguishedName)'"
					$groupToDelete
				}
			}

			# Update the current list by removing groups that are already failed
            $currentGroupsToDeleteOld = $currentGroupsToDelete
			$currentGroupsToDelete = $currentGroupsToDelete | Where { $_.DistinguishedName -notin $nestedFailedGroups.DistinguishedName }
            $currentGroupsToDelete | ConvertTo-Json | Out-File $dropFile
            if ($currentGroupsToDelete.Count -eq $currentGroupsToDeleteOld.Count)
            {
                # all done - list not changing
                break
            }
		}

		$nestedFailedGroups | Sort DistinguishedName -Unique
	}

	end
	{
		Write-Log "End executing $($MyInvocation.MyCommand). GroupsToDeleteJsonFilePath: $GroupsToDeleteJsonFilePath"
	}
}

<#
.Synopsis
	Disables the specified onprem distribtution groups and creates an onprem contacts for them. 
.DESCRIPTION
	Disables the specified onprem distribtution groups and creates onprem contacts for them. 
.PARAMETER OnpremGroupsGoodToDisableJsonFilePath
	The file path of the json file containing good-to-disable groups information previously saved. 
.EXAMPLE
	Disable-OnpremDistributionGroups -OnpremGroupsGoodToDisableJsonFilePath $Global:OnpremGroupsGoodToDisableFileName
#>
function Disable-OnpremDistributionGroups
{
	[CmdletBinding()]
	param
	(
		# The file path of the json file containing good-to-disable groups information previously saved
		[Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path $_ })]
		[string]
		$OnpremGroupsGoodToDisableJsonFilePath
	)
	
	process
	{
		$onpremGroups = Get-Content $OnpremGroupsGoodToDisableJsonFilePath -Raw | ConvertFrom-Json

		$index = 0
		$batchSize = 50
		$count = $onpremGroups.Count
		for ($loopIndex = 0; $loopIndex -le $count/$batchSize; ++$loopIndex)
		{
			$skip = $loopIndex*$batchSize

			Write-Log "Processing loop: $loopIndex skip: $skip"

			[array]$onpremGroupsBatch = $onpremGroups | Select -Skip $skip -First $batchSize

			$Global:batchIndex = 0
			$batchCount = $onpremGroupsBatch.Count
			$scriptBlock = {
				$Group = $input # InputObject from Start-RobustCloudCommand
				$batchIndex = ++$Global:batchIndex

				New-OnpremExchangeSession # in case oprem session is destroyed as well if EXO session needed to be rebuilt
	
				$onlineGroup = Get-DistributionGroup -Identity $Group.PrimarySmtpAddress -ErrorAction Ignore

				if ($onlineGroup)
				{
					Write-Log "($batchIndex / $batchCount / $loopIndex) Disabling the on-prem DL: '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)'"

					$PrimarySmtpAddress = $Group.PrimarySmtpAddress -replace "'", "''" 
					$filter = "{ PrimarySmtpAddress -eq '$PrimarySmtpAddress' }"
					$onpremGroup = Get-OnpremDistributionGroup -Filter $filter -IgnoreDefaultScope -ErrorAction Ignore
					$domainControllerGroup = $null
					if ($onpremGroup)
					{
						$distinguishedName = $onpremGroup.DistinguishedName
						$domainControllerGroup = Get-DomainController $distinguishedName
						Disable-OnpremDistributionGroup -DomainController $domainControllerGroup -Identity $distinguishedName -Confirm:$false
					}
					else
					{
						$msg = "[WARNING] - Onprem group '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)' not found!!"
						Write-Warning $msg
						Write-Log $msg
					}

					$ou = $Global:ContactSyncExclusionOU
					$domainControllerContact = Get-DomainController $ou
					$contactName = $onlineGroup.Name + " (Contact)"
					if ($contactName.Length -gt 64) { $contactName = $contactName.Substring(0, 64).Trim() }

					$ExternalEmailAddress = (($onlineGroup.EmailAddresses | Where { $_ -like "*$Global:HybridEmailRoutingDomain" } | Select -First 1) -Split ":")[1]
					$OnlinePrimarySmtpAddress = $onlineGroup.PrimarySmtpAddress
					if ($ExternalEmailAddress -eq $null)
					{
						$alias = $onlineGroup.alias
						$ExternalEmailAddress = $alias + $Global:HybridEmailRoutingDomain

						Write-Log "Adding Hybrid Email Routing Address to online DL: '$($onlineGroup.Identity)' - '$($onlineGroup.PrimarySmtpAddress)' - '$ExternalEmailAddress'"

						Set-DistributionGroup -Identity $onlineGroup.Guid.ToString() -EmailAddresses @{ Add = $ExternalEmailAddress }

						$onlineGroup = Get-DistributionGroup -Identity $Group.PrimarySmtpAddress
						$ExternalEmailAddress = (($onlineGroup.EmailAddresses | Where { $_ -like "*$Global:HybridEmailRoutingDomain" } | Select -First 1) -Split ":")[1]
						if ($ExternalEmailAddress -eq $null)
						{
							$ExternalEmailAddress = $OnlinePrimarySmtpAddress
							$msg = "[WARNING] - Using PrimarySmtpAddress instead of Hybrid Email Routing Address as targetAddress: '$($onlineGroup.Name)' - '$($ExternalEmailAddress)'"
							Write-Warning $msg
							Write-Log $msg
						}
					}
				}
				else
				{
					$msg = "[ERROR] - Online DL '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)' not found!!"
					Write-Warning $msg
					Write-Log $msg
				}
			}

			Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $onpremGroupsBatch -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
		
			$Global:batchIndex = 0
			$scriptBlock = {
				$Group = $input # InputObject from Start-RobustCloudCommand
				$batchIndex = ++$Global:batchIndex

				New-OnpremExchangeSession # in case oprem session is destroyed as well if EXO session needed to be rebuilt

				Write-Log "($batchIndex / $batchCount / $loopIndex) Post Processing for on-prem DL: '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)'"

				$onlineGroup = Get-DistributionGroup -Identity $Group.PrimarySmtpAddress -ErrorAction Ignore

				if ($onlineGroup)
				{
					$contactName = $onlineGroup.Name + " (Contact)"
					if ($contactName.Length -gt 64) { $contactName = $contactName.Substring(0, 64).Trim() }

					$ExternalEmailAddress = (($onlineGroup.EmailAddresses | Where { $_ -like "*$Global:HybridEmailRoutingDomain" } | Select -First 1) -Split ":")[1]
					$OnlinePrimarySmtpAddress = $onlineGroup.PrimarySmtpAddress
					if ($ExternalEmailAddress -eq $null)
					{
						$ExternalEmailAddress = $OnlinePrimarySmtpAddress
						$msg = "[WARNING] - Using PrimarySmtpAddress instead of Hybrid Email Routing Address as targetAddress: '$($onlineGroup.Name)' - '$($ExternalEmailAddress)'"
						Write-Warning $msg
						Write-Log $msg
					}

					[array]$emailAddresses = $onlineGroup.EmailAddresses
					$emailAddresses += "X500:" + $onlineGroup.LegacyExchangeDN

					$ou = $Global:ContactSyncExclusionOU
					$domainControllerContact = Get-DomainController $ou
					$mailContact = Get-OnpremMailContact -Identity $contactName -DomainController $domainControllerContact -ErrorAction Ignore
					if ($mailContact -eq $null)
					{
						if ($domainControllerGroup -ne $domainControllerContact)
						{
							$secondRemaining = 60
							$sleep = 5
							do
							{
								$PrimarySmtpAddress = $onlineGroup.PrimarySmtpAddress -replace "'", "''" 
								$filter = "{ PrimarySmtpAddress -eq '$PrimarySmtpAddress' }"
								$test = Get-OnpremDistributionGroup -Filter $filter -IgnoreDefaultScope -ErrorAction Ignore
								if ($test)
								{
									Start-Sleep -Seconds $sleep
									$secondRemaining -= $sleep
									Write-Log "Waiting for master object '$($onlineGroup.Name)' - '$($onlineGroup.PrimarySmtpAddress)' to disappear from onprem Exchange... $secondRemaining"
								}
							}
							while ($test -and $secondRemaining -gt 0)
						}

						Write-Log "Creating MailContact for the online DL: '$($onlineGroup.Identity)' - '$($onlineGroup.PrimarySmtpAddress)' - '$($ExternalEmailAddress)'"

						$mailContact = New-OnpremMailContact -Name $contactName -ExternalEmailAddress $ExternalEmailAddress -PrimarySmtpAddress $onlineGroup.PrimarySmtpAddress -DomainController $domainControllerContact `
							-Alias $onlineGroup.Alias -DisplayName $onlineGroup.DisplayName -OrganizationalUnit $ou
					}

					Write-Log "Updating MailContact for the online DL: '$($onlineGroup.Identity)' - '$($onlineGroup.PrimarySmtpAddress)'"

					# TODO - Set ManagedBy AcceptMessagesOnlyFrom, etc
					# TODO - Restore membership with contacts??
					Set-OnpremMailContact -Identity $mailContact.Guid.Tostring() -DomainController $domainControllerContact `
						-EmailAddresses $emailAddresses -EmailAddressPolicyEnabled:$false -ExternalEmailAddress $ExternalEmailAddress `
						-CustomAttribute1 "DoNotSync"
				}
				else
				{
					$msg = "[ERROR] - Online DL '$($Group.Identity)' - '$($Group.PrimarySmtpAddress)' not found!!"
					Write-Warning $msg
					Write-Log $msg
				}
			}

			Start-RobustCloudCommand -Agree -UserName $Global:__UPN__ -Recipients $onpremGroupsBatch -IdentifyingProperty "Name" -ScriptBlock $scriptBlock 
		}
	}
}

#endregion Migrated Onprem Group Deprovisioning functions

Export-ModuleMember -Function *
