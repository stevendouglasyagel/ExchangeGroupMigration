#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script validates the groups in the deactivation list "GroupsToDisable.txt" to check if they can be migrated or what remediation is need before they can be migrated.
#
#	If the group to be disabled is a member of any other groups not in deactivation list, it will be excluded 
#	as for the sake of simplicity we don't want to be needing to update the membership of these other groups with the newly provisioned conact.
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

New-OnpremExchangeSession

Write-Log "Reading DLs afresh..."
Get-OnpremDistributionGroup -ResultSize Unlimited -IgnoreDefaultScope | ConvertTo-Json | Out-File $Global:OnpremGroupExportFileName

Write-Log "Validating groups to check if they are good to disable..."

$AllGroups = Get-Content $Global:OnpremGroupExportFileName -Raw | ConvertFrom-Json
$inputGroups = Get-Content $Global:GroupsToDisableTxtFileName | % { $s = $_.Trim(); if (![string]::IsNullOrEmpty($s)) { $s } }
$inputGroups = $inputGroups | Where { $_ -notmatch "^;" } # remove commented out entries
$GroupsToDisable = $AllGroups | where { $_.PrimarySmtpAddress -in $inputGroups -and $_.CustomAttribute1 -eq "DoNotSync" }
$PrimarySmtpAddresses = $GroupsToDisable.PrimarySmtpAddress
[array]$orphanedGroups = $inputGroups | where { $_ -notin $PrimarySmtpAddresses}
if ($orphanedGroups)
{
    Write-Log "[WARNING] - $($orphanedGroups.Count) groups not found in the All Groups export file: '$Global:OnpremGroupExportFileName'."
    $orphanedGroups | % { Write-Log "Orphaned group found: $_ " }
}

# If a group is nested in some other group that we are not disabling, we exclude the group for the sake of simplicity
# TODO: Remove this check as making contact as member of other groups is helpful for quick rollback purpose.
$GroupsToDisable | select Name, Identity, PrimarySmtpAddress, LegacyExchangeDN, ManagedBy | ConvertTo-Json | Out-File ($Global:OnpremGroupsGoodToDisableFileName + ".tmp")
$NestedFailedGroups = Get-DistributionGroupsFailingNestingValidationForDisablement -GroupsMembersJsonFilePath $Global:OnpremGroupMemberExportFileName -GroupsToDisableJsonFilePath ($Global:OnpremGroupsGoodToDisableFileName + ".tmp")
$NestedFailedGroups | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingNestingValidationForDisablementFileName
[array]$GroupsGoodToDisable = $GroupsToDisable | Where { $_.Name -notin $NestedFailedGroups.Name }
$GroupsGoodToDisable | select Name, Identity, PrimarySmtpAddress, LegacyExchangeDN, ManagedBy | ConvertTo-Json | Out-File $Global:OnpremGroupsGoodToDisableFileName

Write-Log "Total Number of Groups in the input file : $($inputGroups.Count)"
Write-Log "Total Number of Groups to Disable : $($GroupsToDisable.Count)"
Write-Log "Total Number of Groups Failing Validation: $($NestedFailedGroups.Count)"
Write-Log "Total Number of Groups Good To Disable excluding any orphaned groups: $($GroupsGoodToDisable.Count)"

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")