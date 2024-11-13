# Deployment

Check [the official documentation](https://support.microsoft.com/en-us/topic/upgrade-signed-office-vba-macro-projects-to-v3-signature-kb5000676-2b8b3cae-ad64-4b4b-aa85-c4a98ca6da87) for more information on how to use offsign.bat to verify and re-sign office makros.

The offsign.bat script is a wrapper around the `Office File Validation Tool` and the `Office File Converter` to verify and re-sign office makros.

> Please make sure to use the latest supported version of the prerequisites.

## How to use of the runbook of the PSMacroSigning project

1. Create resource group (e.g. rg-PSMacroSigning)
2. Create VM for Hybrid Worker in Azure. If you already have a machine you want to use as a Hybrid Worker, skip this step.
3. Install Office on your virtual machine.
4. Install your self-signed certificate or any other certificate you want to use for code signing on your virtual machine at "Local Computer\Trusted Root Certification Authorities" and "Local Computer\Personal". If you already have a trusted certificate, skip this step.
5. Copy the "deployment" folder of this repository to any folder on your virtual machine
6. Replace the three .txt files in the deployment folder with the prerequisite mentioned in the txt files.
7. Run the New-PSMacroSigningDeployment.ps1 script. This will install the needed files for the offsign.bat script.
8. Create an Hybrid Worker Group in Azure Automation and add the previously configured virtual machine as a Hybrid Worker to it. [See here for more information](https://learn.microsoft.com/en-us/azure/automation/automation-windows-hrw-install).
    1. Create a Azure Automation Account (e.g. aa-PSMacroSigning).
    2. Create a Hybrid Worker Group (e.g. hwg-PSMacroSigning) in the Azure Automation Account.
    3. Add the Hybrid Worker to the Hybrid Worker Group.
9. Create the runbook `Set-OfficeMacroSignature` using the script [Set-AzureRunbook.ps1](../build/Set-AzureRunbook.ps1)

Now you're all set to use the runbook to sign your office makros. You can trigger the `Set-OfficeMacroSignature` runbook manually using the file to sign on the virtual machine directly or pass it's file stream to the runbook.
If you want to take a step further, you can create a Logic App to trigger the runbook on a file upload to a storage account.

Please check the comment-based help of the [Set-OfficeMacroSignature.ps1](../runbooks/SetOfficeMacroSigning.ps1) script for more information on how to use it.
