#########################################################################################################################################################################
#	This script is a part of the library of scripts to help migrate onprem Exchange DLs to cloud-only Exchange Online DLs.
#	The latest version of the script can be downloaded from https://github.com/Microsoft/ExchangeGroupMigration.
#
#	NO WARRANTY OF ANY KIND IS PROVIDED. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS SCRIPT REMAINS WITH THE USER.
#
#	This script populates the membership of the shadow DLs.
#	The shadow group membership is populated as follows:
#		1. Using the group information in the ".\Exports\Online-Groups-Good-to-Migrate-Data.json"
#		2. Using the group membership information in the ".\Exports\Online-DL-Member.json"
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

Add-ShadowGroupMembers -GroupsToMigrateJsonFilePath $Global:OnlineGroupsGoodToMigrateFileName -GroupMembersJsonFilePath $Global:OnlineGroupMemberExportFileName

$currentScriptName = Get-CurrentScriptName
Write-Log ("!!Script '$currentScriptName' execution complete!!")