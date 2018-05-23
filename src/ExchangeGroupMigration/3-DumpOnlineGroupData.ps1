#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script dumps the information on the usage of DLs in EXO.
#	The groups are recreated using this information on the already sycned EXO object.
#
#	All online DLs => .\Exports\Online-DL.json
#	All online DL membership => .\Exports\Online-DL-Member.json
#	All online mailboxes => .\Exports\Online-Mailbox.json
#	All online mailboxes => .\Exports\Online-Mailbox.json
#	All online recipients with delivery restrictions and public delegation based on groups => .\Exports\Online-Recipient-Restrictions-and-Delegates.json
#	All online mailboxes with Full Access Permissions => .\Exports\Online-FullAccess-Permissions.json
#	All online recipients with SendAs Access Permissions => .\Exports\Online-SendAs-Permissions.json
#	All online calender processing information => .\Exports\Online-CalendarScheduling-Permissions-Data.json
#########################################################################################################################################################################

#### Import the Microsoft Exchange Online Powershell Module if running from regular PowerShell prompt / ISE
$cmd = Get-Command Connect-EXOPSSession -ErrorAction Ignore
if ($cmd -eq $null)
{
    Write-Debug "Loading EXO Module"
	Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse | sort LastWriteTime).FullName | select -Last 1)
}
else
{
    Write-Debug "EXO Module is already loaded"
}

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Write-Log "Reading DLs to be migrated..."
Get-OnlineDistributionGroups | ConvertTo-Json | Out-File $Global:OnlineGroupExportFileName

Write-Log "Reading DL membership..."
Get-OnlineDistributionGroupMembers -GroupsJsonFilePath $Global:OnlineGroupExportFileName | ConvertTo-Json -Depth 3 | Out-File $Global:OnlineGroupMemberExportFileName

Write-Log "Reading Mailboxes and DLs with delivery restrictions configured..."
Get-OnlineRecipientDeliveryRestrictionsAndPublicDelegates | ConvertTo-Json | Out-File $Global:OnlineDeliveryRestrictionsExportFileName

Write-Log "Reading Mailboxes for checking FullAccess permissions configured..."
Get-OnlineMailboxes | ConvertTo-Json | Out-File $Global:OnlineMailboxExportFileName

Write-Log "Reading Mailbox FullAccess permissions configured..."
Get-OnlineMailboxFullAccessPermissions -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnlineFullAccessPermissionsExportFileName

Write-Log "Reading SendAs permissions on recipients..."
Get-OnlineRecipientSendAsPermissions -GroupsJsonFilePath $Global:OnlineGroupExportFileName -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnlineSendAsPermissionsExportFileName

Write-Log "Reading Calendar Processing configurations..."
Get-OnlineCalendarProcessingConfiguration -MailboxesJsonFilePath $Global:OnlineMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnlineCalendarSchedulingPermissionsExportFileName

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")