#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script validates the groups that are good to be migrated.
#	It also prepares the list of groups that must be migrated as security groups if they are used in FullAccess / SendAs delegations.
#	If a group is not used for these delegations, it will be recreated as a pure distribution group (even if the original groups was a security group).  
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

$AllOnlineGroups = Get-Content $Global:OnlineGroupExportFileName -Raw | ConvertFrom-Json
[array]$OnpremGroupsGoodToMigrate = Get-Content $Global:OnpremGroupsGoodToMigrateFileName -Raw | ConvertFrom-Json

$PrimarySmtpAddresses = $OnpremGroupsGoodToMigrate.PrimarySmtpAddress
$OnlineGroupsGoodToMigrate = $AllOnlineGroups | Where { $_.PrimarySmtpAddress -in $PrimarySmtpAddresses}

# Sort groups so that we migrate the parents first
# TODO - do multiple nesting level sorting
$onlineGroupsMembership = Get-Content $Global:OnlineGroupMemberExportFileName -Raw | ConvertFrom-Json
$onlineChildGroupMemberNames = ($onlineGroupsMembership.Members | Where { $_.RecipientType -match "Group" } | Sort Name -Unique).Name
$onlineParentGroups = $OnlineGroupsGoodToMigrate | Where { $_.Name -in $onlineChildGroupMemberNames }
$onlineNonParentGroups = $OnlineGroupsGoodToMigrate | Where { $_.Name -notin $onlineChildGroupMemberNames }
$OnlineGroupsGoodToMigrate = [array]$onlineParentGroups + [array]$onlineNonParentGroups

$OnlineGroupsGoodToMigrate | ConvertTo-Json -Depth 3 | Out-File $Global:OnlineGroupsGoodToMigrateFileName

$OnpremGroupsGoodToMigrateCount = $OnpremGroupsGoodToMigrate.Count
$OnlineGroupsGoodToMigrateCount = $OnlineGroupsGoodToMigrate.Count
Write-Log "Onprem Groups Good-To-Migrate: $OnpremGroupsGoodToMigrateCount"
Write-Log "Online Groups Good-To-Migrate: $OnlineGroupsGoodToMigrateCount"

if ($OnpremGroupsGoodToMigrateCount -ne $OnlineGroupsGoodToMigrateCount)
{
	Write-Log "[WARNING] - Onprem and Online Groups Good-To-Migrate does not match.."

    $OnpremPrimarySmtpAddresses = $OnpremGroupsGoodToMigrate.PrimarySmtpAddress
    $OnlinePrimarySmtpAddresses = $OnlineGroupsGoodToMigrate.PrimarySmtpAddress
    $diff = Compare-Object $OnpremPrimarySmtpAddresses $OnlinePrimarySmtpAddresses
    $diff | % { Write-Log "[WARNING] - Missing Group from Onprem or Online: $_" }
}

Write-Log "Checking Good-To-Migrate DLs that need to be created as MESG..."

[array]$OnlineGroupsGoodToMigrate = Get-Content $Global:OnlineGroupsGoodToMigrateFileName -Raw | ConvertFrom-Json
$OnlineGroupsGoodToMigrateNames = $OnlineGroupsGoodToMigrate.Name

$ShadowMailboxFullAccessPermissions = Get-Content $Global:OnlineFullAccessPermissionsExportFileName -Raw | ConvertFrom-Json
[array]$ShadowMailboxFullAccessPermissionGroups = ($ShadowMailboxFullAccessPermissions | Where { $_.User -in $OnlineGroupsGoodToMigrateNames } | Select User -Unique).User

$ShadowRecipientSendAsPermissions = Get-Content $Global:OnlineSendAsPermissionsExportFileName -Raw | ConvertFrom-Json
[array]$ShadowRecipientSendAsPermissionGroups = ($ShadowRecipientSendAsPermissions | Where { $_.Trustee -in $OnlineGroupsGoodToMigrateNames }  | Select Trustee -Unique).Trustee

[array]$PermissionGroups = $ShadowMailboxFullAccessPermissionGroups + $ShadowRecipientSendAsPermissionGroups | Select -Unique
$OnlineGroupsGoodToMigrateAsSecurityGroups = $OnlineGroupsGoodToMigrate | Where { $_.Name -in $PermissionGroups }
$OnlineGroupsGoodToMigrateAsSecurityGroups | ConvertTo-Json | Out-File $Global:OnlineGroupsGoodToMigrateAsSecurityGroupsFileName

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")