﻿#Requires -Modules @{ ModuleName="Az.Automation"; ModuleVersion="1.9.1" }

<#
	.SYNOPSIS
	Set-AzureRunbook.ps1

	.DESCRIPTION
	Import and publish changes to a specific runbook of a Azure Automation account.

	.PARAMETER AutomationAccountName
	Provide a String containing the name of your Azure Automation.

	.PARAMETER ResourceGroupName
	Provide a String containing the name of your Azure Resource Group.

	.PARAMETER ScriptPath
	Provide a String containing the path of your script files.

	.PARAMETER Type
	Provide a String containing the type of your script.
	This parameter defaults to 'PowerShell'.

	.PARAMETER Publish
	Provide a Boolean value to publish the runbook automatically after import.
	This parameter defaults to $true.

	.PARAMETER Force
	If this parameter is provided and the specified runbook exists already, the script of that runbook will be overwritten.

	.EXAMPLE
	Set-AzureRunbook.ps1 -AutomationAccountName auto-ExoTasks-Admin -ResourceGroupName rg-ExoTasks-Admin -ScriptPath 'C:\Repository\UnattendedTasks_ExchangeAdministrator\AzureAutomation\runbooks'

	Import and publish the script located at the specified scriptpath to the specified runbook.

	.NOTES
	Details at https://learn.microsoft.com/en-us/azure/automation/manage-runbooks
#>

[CmdletBinding(SupportsShouldProcess)]
param (
	[Parameter(Mandatory = $true)]
    [String]
    $AutomationAccountName,
    
	[Parameter(Mandatory = $true)]
    [String]
    $ResourceGroupName,
    
	[Parameter()]
    [String]
    $ScriptPath = "$(Split-Path -Path $PSScriptRoot)\runbook",

	[Parameter()]
	[ValidateSet('PowerShell','GraphicalPowerShell','PowerShellWorkflow','GraphicalPowerShellWorkflow','Python2')]
    [String]
    $Type = 'PowerShell',

	[Parameter()]
    [Bool]
    $Publish = $true,

	[Parameter()]
    [Switch]
    $Force
)

# Check, if PowerShell connection to Azure is established
Write-Host 'Checking, if connection to Azure has been established...' -NoNewline
if ($null -eq (Get-AzContext)) {
	$exception = [System.InvalidOperationException]::new('Not yet connected to the Azure Service. Use "Connect-AzAccount -TenantId <TenantId>" to establish a connection and select the correct subscription using "Set-AzContext"!')
	$errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, 'NotConnected', 'InvalidOperation', $null)
	
	$PSCmdlet.ThrowTerminatingError($errorRecord)
}
Write-Host 'OK' -ForegroundColor Green

foreach($runbook in (Get-ChildItem -Path "$ScriptPath" -Name "*.ps1")) {

	# Import runbook to Azure Automation
	$importAzAutomationRunbookParams = @{
		AutomationAccountName = $AutomationAccountName
		Name                  = "$($runbook.replace('-','').replace('.ps1',''))"
		ResourceGroupName     = $ResourceGroupName
		Type                  = $Type
		Path                  = "$(Join-Path -Path $ScriptPath -ChildPath $runbook)"
		Published			  = $Publish
		Force				  = $Force
		WhatIf				  = $WhatIfPreference
		Confirm				  = $ConfirmPreference
	}
	Import-AzAutomationRunbook @importAzAutomationRunbookParams

}