<#
	.SYNOPSIS
	Installs prerequisites for scripts.
	
	.DESCRIPTION
	Installs prerequisites for scripts.

	.INPUTS
	None.

	.OUTPUTS
	None.

	.EXAMPLE
	PS> .\Install-Scripts
#>
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Requires -Version 5.0
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param ([Parameter()] [switch] $UpdateHelp,
	   [Parameter(Mandatory = $true)] [string] $ModulesPath)

Begin
{
	$script = $MyInvocation.MyCommand.Name
	if(-Not (Test-Path ".\$script"))
	{
		Write-Host "Installation must be run from the same directory as the installer script."
		exit
	}

	if(-Not (Test-Path $ModulesPath))
	{
		Write-Host "'$ModulesPath' was not found."
		exit
	}

	$Env:PSModulePath += ";$ModulesPath"
	
	Import-LocalModule Varan.PowerShell.SelfElevate
	$boundParams = @{}
	$PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object { $boundParams[$_.Key] = $_.Value }
	Open-ElevatedConsole -CallerScriptPath $PSCommandPath -OriginalBoundParameters $boundParams
}

Process
{
	Import-LocalModule "Varan.PowerShell.Base"
	Import-LocalModule "Varan.PowerShell.Common"
	Import-LocalModule "Varan.PowerShell.SelfElevate"
		
	Add-PathToProfile -PathVariable 'Path' -Path (Get-Location).Path
	Add-PathToProfile -PathVariable 'PSModulePath' -Path $ModulesPath
	
	Add-ImportModuleToProfile "Varan.PowerShell.Base"
	Add-ImportModuleToProfile "Varan.PowerShell.Common"
	Add-ImportModuleToProfile "Varan.PowerShell.SelfElevate"
	
	Add-AliasToProfile -Script 'Get-CustomConfigurationHelp' -Alias 'cchelp'
	Add-AliasToProfile -Script 'Get-CustomConfigurationHelp' -Alias 'gsh'
	Add-AliasToProfile -Script 'Install-CustomConfiguration' -Alias 'icc'
	Add-AliasToProfile -Script 'Install-CustomConfiguration' -Alias 'cci'
	Add-AliasToProfile -Script 'Get-CustomConfigurationScriptVersion' -Alias 'ccver'
	Add-AliasToProfile -Script 'Get-CustomConfigurationScriptVersion' -Alias 'gccsv'
	
	Add-LineToProfile -Text '$ConfirmPreference = ''None'''
}

End
{
	Format-Profile
	Complete-Install
}