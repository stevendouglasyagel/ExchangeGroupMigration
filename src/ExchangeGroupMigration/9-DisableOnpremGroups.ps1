#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script disabled the onprem groups and creates an contact for the for online counter part of these onprem group.
#
#	The target address of the contact is set to tenant.mail.onmicrosoft.com to allow for proper hybrid coexistence.
#	If the online group is missing the tenant.mail.onmicrosoft.com proxyaddress, it will be updated to add this addtional proxy address.
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

Disable-OnpremDistributionGroups -OnpremGroupsGoodToDisableJsonFilePath $Global:OnpremGroupsGoodToDisableFileName

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")