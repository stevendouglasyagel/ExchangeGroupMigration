#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script validates the groups in the deactivation list "GroupsToDelete.txt" to check if they can be deleted.
#
#	If the group to be deleted is a security group, it will be excluded as it needs to be manually investigated if it is really needed or not.
#	If the group to be deleted is a member of any other groups not in deletion list, it will be excluded for the sake of "atomicity" of operation.
#	If this is the last batch to be deleted, any validation failures can be ingnored.
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

Write-Log "Reading AD Groups afresh..."

$PropertyNames =  @("CanonicalName", "CN", "Description", "DisplayName", "DistinguishedName", "GroupCategory", "GroupScope", "groupType", "isDeleted", "LastKnownParent",
	"ManagedBy", "member", "MemberOf", "Members", "Name", "ObjectCategory", "ObjectClass", "objectGUID", "objectSid", "ProtectedFromAccidentalDeletion",
	"SamAccountName", "sAMAccountType", "SID", "SIDHistory", "whenChanged","whenCreated")

$adGroups = $Global:DomainControllerLookup.Values | Where { $_ -like '*.ukroi.*' -or $_ -like '*.in.*' -or $_ -like '*.tsl.*' } | % {
    $dc = $_
    Write-Log "Reading groups on domain controller '$dc'..."
    Get-ADGroup -Filter { legacyEXchangeDN -notlike '*' } -Server $dc -Properties $PropertyNames
}

$adGroups | Select $PropertyNames | ConvertTo-Json | Out-File $Global:OnpremAdGroupExportFileName

Write-Log "Validating groups to check if they are good to delete..."

$AllGroups = Get-Content $Global:OnpremAdGroupExportFileName -Raw | ConvertFrom-Json
$inputGroups = Get-Content $Global:GroupsToDeleteTxtFileName | % { $s = $_.Trim(); if (![string]::IsNullOrEmpty($s)) { $s } }
$inputGroups = $inputGroups | Where { $_ -notmatch "^;" } # remove commented out entries
$GroupsToDelete = $AllGroups | where { $_.CN -in $inputGroups }
$GroupsToDeleteNames = $GroupsToDelete.CN
[array]$orphanedGroups = $inputGroups | where { $_ -notin $GroupsToDeleteNames}
if ($orphanedGroups)
{
    Write-Log "[WARNING] - $($orphanedGroups.Count) groups not found in the All Groups export file: '$Global:OnpremAdGroupExportFileName'."
    $orphanedGroups | % { Write-Log "Orphaned group found: $_ " }
}

# If the group is a security group, we won't delete it

[array]$SecurityGroups = $GroupsToDelete | Where { $_.GroupCategory -eq 1 }
if ($SecurityGroups)
{
    Write-Log "[WARNING] - $($SecurityGroups.Count) groups are security groups and will not be processed."
    $SecurityGroups.DistinguishedName | % { Write-Log "Security group found: $_ " }
}

# If a group is nested in some other group that we are not disabling, we exclude the group for the sake of simplicity
$SecurityGroupDistinguishedNames = $SecurityGroups.DistinguishedName
$GroupsGoodToDelete = $GroupsToDelete | Where { $_.DistinguishedName -notin $SecurityGroupDistinguishedNames } 
$GroupsGoodToDelete | ConvertTo-Json | Out-File ($Global:OnpremGroupsGoodToDeleteFileName + ".tmp")
$NestedFailedGroups = Get-DistributionGroupsFailingNestingValidationForDeletion -GroupsToDeleteJsonFilePath ($Global:OnpremGroupsGoodToDeleteFileName + ".tmp")
$NestedFailedGroups | ConvertTo-Json -Depth 5 | Out-File $Global:OnpremGroupsFailingNestingValidationForDeletionFileName
[array]$GroupsGoodToDelete = $GroupsGoodToDelete | Where { $_.Name -notin $NestedFailedGroups.Name }
$GroupsGoodToDelete | select Name, ObjectGUID, DistinguishedName, MemberOf | ConvertTo-Json | Out-File $Global:OnpremGroupsGoodToDeleteFileName
$GroupsGoodToDelete | Select Name, DistinguishedName  | Export-Csv ($Global:GroupsToDeleteTxtFileName + ".csv") -NoTypeInformation

Write-Log "Total Number of Groups in the input file : $($inputGroups.Count)"
Write-Log "Total Number of Groups to Delete : $($GroupsToDelete.Count)"
Write-Log "Total Number of Groups Failing Validation: $($NestedFailedGroups.Count)"
Write-Log "Total Number of Groups Failing Security Validation: $($SecurityGroups.Count)"
Write-Log "Total Number of Groups Good To Delete : $($GroupsGoodToDelete.Count)"

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")