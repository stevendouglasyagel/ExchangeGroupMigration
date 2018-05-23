#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script animates the shadow DL. It:
#		1. Stamps extensionAttribute1 = DoNotSync to move the onprem DLs out of scope of dirsync
#		2. Deletes the DL object from Azure AD
#		3. Updates the DL Name, DisplayName, Alias, EmailAddresses, HiddenFromAddressListsEnabled etc.
#		4. Also add the LegacyExchangeDN or the original DL as an addtional x500 proxyaddress
#
#	Please disable the AAD Connect sync cycle before runnning this script. Otherwise there is a chance that on-prem DL, if gets synced, will soft-match again.
#########################################################################################################################################################################

while(!($choice = (Read-Host "Have you disabled the AAD Connect sync cycle?") -match "y")){ "y?" }

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

#### Import the MSOnline module
if (!(Get-Module -Name MSOnline))
{
	Import-Module MSOnline
	if ($Global:AzureADUserCredential -eq $null)
	{
		$Global:AzureADUserCredential = Get-Credential
	}

	Connect-MsolService -Credential $Global:AzureADUserCredential # no use of specifying creds if this requires MFA
}

Switch-ShadowGroups -OnlineGroupsGoodToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -OnpremGroupsGoodToMigrateJsonFilePath $Global:OnpremGroupsGoodToMigrateFileName

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")
Write-Log "Do a manual Delta Import and Delta Sync on the affected forest and confirm that the DLs are infact descoped before enabling the sync cycle again."
