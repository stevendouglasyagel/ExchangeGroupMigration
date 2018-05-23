#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script creates the shadow groups for each DL synced from onprem.
#	The new shadow group as follows:
#		1. Created using the information in the ".\Exports\Online-Groups-Good-to-Migrate-Data.json"
#		2. HiddenFromAddressListsEnabled set to $true (in batches of 50 as it seems there needs to be some wait time after a DL is created)
#		3. The deleivery restriction and public delegates are mirrored after all shadow DLs are created
#		4. The mailbox FullAccess permissions are mirrored after all shadow DLs are created
#		5. The recipient SendAs permissions are mirrored after all shadow DLs are created
#		6. Membership is NOT restored by this script.
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

Write-Log "Creating the shadow DLs..."
Create-ShadowGroupShellObjects -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -GroupsToMigrateAsSecurityGroupsJsonFilePath $Global:OnlineGroupsGoodToMigrateAsSecurityGroupsFileName

Write-Log "Configuring delivery restrictions on the shadow DLs and update delivery restrictions on the any other (cloud-managed) recipients..."
Set-RecipientDeliveryRestrictionsAndPublicDelegates -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -RecipientDeliveryRestrictionsJsonFilePath $Global:OnlineDeliveryRestrictionsExportFileName

Write-Log "Configuring FullAccess permissions on mailboxes..."
Add-MailboxFullAccessPermissions -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -MailboxFullAccessPermissionsJsonFilePath $Global:OnlineFullAccessPermissionsExportFileName

Write-Log "Configuring SendAs permissions on recipients..."
Add-RecipientSendAsPermissions -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -RecipientSendAsPermissionsJsonFilePath $Global:OnlineSendAsPermissionsExportFileName

# TODO - Set CalenderProcessing Policies

# TODO - Set transport rules

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")