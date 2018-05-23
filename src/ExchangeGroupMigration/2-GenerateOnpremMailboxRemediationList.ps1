#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script validates the groups in the migration list "GroupsToMigrate.txt" to check if they can be migrated or what remediation is need before they can be migrated.
#
#	Groups failing validation because being used in delivery restrictions and public delegation on onprem mailboxes / groups not being migrated => .\Exports\Onprem-Groups-Failing-Delilvery-Restrictions-And-Delegates-Validation.json
#	Groups failing validation because being used in FullAccess Permissions on onprem mailboxes => .\Exports\Onprem-Groups-Failing-FullAccess-Delegations-Validation.json
#	Groups failing validation because being used in SendAs Permissions on onprem mailboxes => .\Exports\Onprem-Groups-Failing-SendAs-Delegations-Validation.json
#	Groups failing validation because being used in Calender Processing on onprem room mailboxes => .\Exports\Onprem-Groups-Failing-CalendarScheduling-Delegations-Validation.json
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Write-Log "Processing groups to check if they are good to migrate..."

$AllGroups = Get-Content $Global:OnpremGroupExportFileName -Raw | ConvertFrom-Json
$inputGroups = Get-Content $Global:GroupsToMigrateTxtFileName | % { $s = $_.Trim(); if (![string]::IsNullOrEmpty($s)) { $s } }
$inputGroups = $inputGroups | Where { $_ -notmatch "^;" } # remove commented out entries
$GroupsToMigrate = $AllGroups | where { $_.PrimarySmtpAddress -in $inputGroups -and $_.CustomAttribute1 -ne "DoNotSync" }
$GroupsNotOnMigrationList = $AllGroups | where { $_.PrimarySmtpAddress -notin $inputGroups -and $_.CustomAttribute1 -ne "DoNotSync" }
$PrimarySmtpAddresses = $GroupsToMigrate.PrimarySmtpAddress
[array]$orphanedGroups = $inputGroups | where { $_ -notin $PrimarySmtpAddresses}
if ($orphanedGroups)
{
    Write-Log "[WARNING] - $($orphanedGroups.Count) groups not found in the All Groups export file: '$Global:OnpremGroupExportFileName'."
    $orphanedGroups | % { Write-Log "Orphaned group found: $_ " }
}

$GroupsToMigrate | ConvertTo-Json -Depth 3 | Out-File $Global:OnpremGroupsToMigrateFileName

Write-Log "Processing groups to check if they are used in delivery restrictions or public delegates..."
Get-DistributionGroupsFailingDeliveryRestrictionsAndDelegatesValidation -DeliveryRestrictionsAndDelegationsJsonFilePath $Global:OnpremDeliveryRestrictionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName `
    | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingDeliveryRestrictionsValidationFileName

Write-Log "Processing groups to check if they are used to grant FullAccess mailbox permissions..."
Get-DistributionGroupsFailingFullAccessDelegationsValidation -FullAccessDelegationsJsonFilePath $Global:OnpremFullAccessPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName `
    | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingFullAccessDelegationsValidationFileName

Write-Log "Processing groups to check if they are used to grant SnedAs recipient permissions..."
Get-DistributionGroupsFailingSendAsDelegationsValidation -SendAsDelegationsJsonFilePath $Global:OnpremSendAsPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName `
    | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingSendAsDelegationsValidationFileName

Write-Log "Processing groups to check if they are used to grant calendar scheduling permissions..."
Get-DistributionGroupsFailingCalendarSchedulingDelegationsValidation -CalendarSchedulingDelegationsJsonFilePath $Global:OnpremCalendarSchedulingPermissionsExportFileName -GroupsToMigrateJsonFilePath $Global:OnpremGroupsToMigrateFileName `
    | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingCalendarSchedulingDelegationsValidationFileName

[array]$DeliveryRestrictionFailedGroups = Get-Content $Global:OnpremGroupsFailingDeliveryRestrictionsValidationFileName -Raw | ConvertFrom-Json
[array]$FullAccessFailedGroups = Get-Content $Global:OnpremGroupsFailingFullAccessDelegationsValidationFileName -Raw | ConvertFrom-Json
[array]$SendAsFailedGroups = Get-Content $Global:OnpremGroupsFailingSendAsDelegationsValidationFileName -Raw | ConvertFrom-Json
[array]$CalendarSchedulingFailedGroups = Get-Content $Global:OnpremGroupsFailingCalendarSchedulingDelegationsValidationFileName -Raw | ConvertFrom-Json

# TODO - Transport Rules using groups

$AllValidationFailedGroups = @()
$AllValidationFailedGroups += $DeliveryRestrictionFailedGroups
$AllValidationFailedGroups += $FullAccessFailedGroups
$AllValidationFailedGroups += $SendAsFailedGroups
$AllValidationFailedGroups += $CalendarSchedulingFailedGroups

$AllValidationFailedGroups + $GroupsNotOnMigrationList | Sort Identity -Unique | ConvertTo-Json -Depth 3 | Out-File ($Global:OnpremGroupsExcludedFileName + ".tmp")

# If a group is not being migrated then exclude all the child groups as well as we cannot update the membership of the parent dirsycned group with the new cloud group
Write-Log "Processing groups to exclude all the child groups as well if there are any parent groups that we are not migrating..."
Get-DistributionGroupsFailingNestingValidation -GroupsMembersJsonFilePath $Global:OnpremGroupMemberExportFileName -GroupsToExcludeJsonFilePath  ($Global:OnpremGroupsExcludedFileName + ".tmp") `
    | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingNestingValidationFileName

Remove-Item ($Global:OnpremGroupsExcludedFileName + ".tmp") -Force -Confirm:$false

[array]$NestedFailedGroups = Get-Content $Global:OnpremGroupsFailingNestingValidationFileName -Raw | ConvertFrom-Json

$AllValidationFailedGroups = $AllValidationFailedGroups +  $NestedFailedGroups | Sort Identity -Unique
$GroupsGoodToMigrate = $GroupsToMigrate | Where { $_.Name -notin $AllValidationFailedGroups.Name}
$OrphanedGroupsToMigrate = $orphanedGroups | % {
		$props = @{
			"Name" = $_;
			"Identity" = $_;
			"PrimarySmtpAddress" = $_;
			"LegacyExchangeDN" = $null;
		}

		New-Object –TypeName PSObject –Prop $props
	}

[array]$GroupsGoodToMigrate = $GroupsGoodToMigrate + $OrphanedGroupsToMigrate
$GroupsGoodToMigrate | select Name, Identity, PrimarySmtpAddress, LegacyExchangeDN, ManagedBy | ConvertTo-Json | Out-File $Global:OnpremGroupsGoodToMigrateFileName

[array]$GroupsWithOwnersPopulated = $GroupsGoodToMigrate | Where { $_.ManagedBy -ne $null}
$GroupsWithOwnersPopulated | select Name, Identity, PrimarySmtpAddress, LegacyExchangeDN, ManagedBy | ConvertTo-Json | Out-File $Global:OnpremGroupsGoodToMigrateWithOwnersFileName

Write-Log "Total Number of Groups to Migrate : $($GroupsToMigrate.Count)"
Write-Log "Total Number of Groups Failing Validation: $($AllValidationFailedGroups.Count)"
Write-Log "`t Groups Failing DeliveryRestrictions And Delegates Validation: $($DeliveryRestrictionFailedGroups.Count)"
Write-Log "`t Groups Failing FullAccess Validation: $($FullAccessFailedGroups.Count)"
Write-Log "`t Groups Failing SendAs Validation: $($SendAsFailedGroups.Count)"
Write-Log "`t Groups Failing Calendar Scheduling Permissions Validation: $($CalendarSchedulingFailedGroups.Count)"
Write-Log "`t Groups Failing Nesting Validation (due to above two categories): $($NestedFailedGroups.Count)"
Write-Log "Total Number of Groups Good To Migrate including any orphaned groups: $($GroupsGoodToMigrate.Count)"
Write-Log "Total Number of Groups Good To Migrate that have Owner Populated: $($GroupsWithOwnersPopulated.Count)"

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")