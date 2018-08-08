# Exchange Group Migration

This project is a library of sample PowerShell scripts that demonstrate how onprem Exchange distribution groups can be converted to cloud-only Exchange Online distribution groups. 

The scripts are designed so that they facilitate a phased migration approach to minimise any adverse or unexpected impact and make roll-back of the change easier. If "T" is the "migration" date for a validated batch of DLs, the recommended schedule for the DL migration project is as follows:
1. **T-n** days - Implement AAD Connect change mentioned in the prerequisites below.
2. **T-5** days - Identify the DLs to migrate and perform any remediations required.
3. **T-2** days - Create "Shadow" DLs in Exchange Online.
4. **T-0** days - Animate "Shadow" DLs. This is when the onprem groups stop syncing to Azure AD and DLs created in Exchange Online take over.
5. **T+2** days - Disable Onprem DLs and provision mail-enabled contacts instead.
6. **T+5** days - Delete Onprem AD group objects.
7. **T+n** days - Convert to Office Groups.

Prerequisites:

1. PowerShell 5.1 or later.
2. [Exchange Online PowerShell Module with MFA support](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell).
3. The scripts assume that you have already configured a custom Inbound Sync Rule in AAD Connect that will filter out any group that the migration process stamps with extensionAttribute1 = DoNotSync.


How to use the tool:

1. Download the latest release ExchangeGroupMigration.zip from the [releases](https://github.com/Microsoft/ExchangeGroupMigration/releases) tab under the Code tab tab, UNBLOCK the downloaded zip file and extract the zip file to an empty local folder on a machine which has the prerequisites installed and connectivity to the AD/Exchange Remote PowerShell.
2. Launch PowerShell ISE and change the working directory to the folder where scripts are located. Then open the ExchangeGroupMigration.psm1 PowerShell module and edit the environment variables defined in the "Global Variables" region as appropriate to your environment.
3. Open the PowerShell scripts in the order of their names, review the purpose of the script and execute them in order one after another as per the migration schedule. 


# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
