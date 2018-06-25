#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script deleted the onprem groups that have been previously validated to be good to be deleted.
#
#########################################################################################################################################################################

#### Import the ExchangeGroupMigration module
if (Get-Module -Name ExchangeGroupMigration) { Remove-Module -Name ExchangeGroupMigration }
Import-Module -Name (Join-Path -Path $PWD -ChildPath "ExchangeGroupMigration.psm1") -ErrorAction Stop

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

[array]$onpremGroupsGoodToDelete = Get-Content $Global:OnpremGroupsGoodToDeleteFileName -Raw | ConvertFrom-Json
[array]$groupDNs = $onpremGroupsGoodToDelete.DistinguishedName

#$groupDNs = Get-Content $Global:GroupsGoodToDeleteTxtFileName | % { $s = $_.Trim(); if (![string]::IsNullOrEmpty($s)) { $s } }
#[array]$groupDNs = $inputGroups | Where { $_ -notmatch "^;" } # remove commented out entries

$count = $groupDNs.Count
$index = 0
$groupDNs | % {
    $distinguishedName = $_
    ++$index
    $Error.Clear()
    Write-Log "($index / $count) - Deleting group '$distinguishedName'..."

    $domainController = Get-DomainController $distinguishedName

    Remove-ADGroup -Identity $distinguishedName -Server $domainController -Confirm:$false

    if ($Error)
    {
        Write-Log "($index / $count) - Error deleting group '$distinguishedName'. Error: $Error"
    }
}

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")