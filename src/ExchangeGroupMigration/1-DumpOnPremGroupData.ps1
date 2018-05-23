#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#   
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script dumps the information on the usage of onprem DLs.
#	This infromation will be used to validate which groups can be migrated or what remediation is need before they can be migrated.
#	This information is not used for recreating groups in the cloud. The groups are recreated using the information on the already sycned EXO object.
#	
#	All onprem DLs => .\Exports\Onprem-DL-Data.json
#	All onprem DL membership => .\Exports\Onprem-DL-Member-Data.json
#	All onprem mailboxes => .\Exports\Onprem-Mailbox-Data.json
#	All onprem recipients with delivery restrictions and public delegation based on groups => .\Exports\Onprem-Delilvery-Restrictions-And-Delegates-Data.json
#	All onprem mailboxes with Full Access Permissions => .\Exports\Onprem-FullAccess-Permissions-Data.json
#	All onprem recipients with SendAs Access Permissions => .\Exports\Onprem-SendAs-Permissions-Data.json
#	All onprem calender processing information => .\Exports\Onprem-CalendarScheduling-Permissions-Data.json
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

New-OnpremExchangeSession

Write-Log "Reading DLs to be migrated..."
Get-OnpremDistributionGroup -ResultSize Unlimited -IgnoreDefaultScope | ConvertTo-Json | Out-File $Global:OnpremGroupExportFileName

Write-Log "Reading DL membership..."
Get-OnpremDistributionGroupMembers -GroupsJsonFilePath $Global:OnpremGroupExportFileName | ConvertTo-Json -Depth 3 | Out-File $Global:OnpremGroupMemberExportFileName

Write-Log "Reading onprem Mailboxes and DLs with delivery restrictions configured..."
Get-OnpremRecipientDeliveryRestrictionsAndPublicDelegates | ConvertTo-Json | Out-File $Global:OnpremDeliveryRestrictionsExportFileName

Write-Log "Reading Mailboxes for checking FullAccess permissions configured..."
Get-OnpremMailbox -ResultSize Unlimited -IgnoreDefaultScope | ConvertTo-Json | Out-File $Global:OnpremMailboxExportFileName

Write-Log "Reading Onprem Mailbox FullAccess permissions configured..."
Get-OnpremMailboxFullAccessPermissions -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnpremFullAccessPermissionsExportFileName

Write-Log "Reading Onprem Recipient SendAs permissions configured..."
Get-OnpremRecipientSendAsPermissions -GroupsJsonFilePath $Global:OnpremGroupExportFileName -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnpremSendAsPermissionsExportFileName

Write-Log "Reading Onprem Calendar Processing configurations..."
Get-OnpremCalendarProcessingConfiguration -MailboxesJsonFilePath $Global:OnpremMailboxExportFileName | ConvertTo-Json | Out-File $Global:OnpremCalendarSchedulingPermissionsExportFileName

# TODO - Transport Rules using groups

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")