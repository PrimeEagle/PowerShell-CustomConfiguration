<#
	.SYNOPSIS
	Uninstalls prerequisites for scripts.
	
	.DESCRIPTION
	Uninstalls prerequisites for scripts.

	.INPUTS
	None.

	.OUTPUTS
	None.

	.EXAMPLE
	PS> .\Uninstall-Scripts
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
		Write-Host "Uninstallation must be run from the same directory as the uninstaller script."
		exit
	}

	if(-Not (Test-Path $ModulesPath))
	{
		Write-Host "'$ModulesPath' was not found."
		exit
	}

	$Env:PSModulePath += ";$ModulesPath"
	
	if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
		Start-Process -FilePath "pwsh.exe" -ArgumentList "-File `"$PSCommandPath`"", "-ModulesPath `"$ModulesPath`"" -Verb RunAs
		exit
	}
}

Process
{
	Import-LocalModule "Varan.PowerShell.Base"
	Import-LocalModule "Varan.PowerShell.Common"
	Import-LocalModule "Varan.PowerShell.SelfElevate"
		
	Remove-PathFromProfile -PathVariable 'Path' -Path (Get-Location).Path
	Remove-PathFromProfile -PathVariable 'PSModulePath' -Path $ModulesPath
	
	Remove-ImportModuleFromProfile "Varan.PowerShell.Base"
	Remove-ImportModuleFromProfile "Varan.PowerShell.Common"
	Remove-ImportModuleFromProfile "Varan.PowerShell.SelfElevate"
	
	Remove-AliasFromProfile -Script 'Get-SystemHelp' -Alias 'syshelp'
	Remove-AliasFromProfile -Script 'Get-SystemHelp' -Alias 'gsh'
	Remove-AliasFromProfile -Script 'Install-CustomConfiguration' -Alias 'sysic'
	Remove-AliasFromProfile -Script 'Install-CustomConfiguration' -Alias 'icc'
	Remove-AliasFromProfile -Script 'Get-CustomConfigurationScriptVersion' -Alias 'ccver'
	Remove-AliasFromProfile -Script 'Get-CustomConfigurationScriptVersion' -Alias 'gccsv'
	
	Remove-PathFromProfile -Text '$ConfirmPreference = ''None'''
}

End
{
	Format-Profile
	Complete-Install
}