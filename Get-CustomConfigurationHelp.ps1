<#
	.SYNOPSIS
	Shows help summary for all Custom Configuration scripts.
	
	.DESCRIPTION
	Shows help summary for all Custom Configuration scripts.

	.INPUTS
	None.

	.OUTPUTS
	None.

	.EXAMPLE
	PS> .\Get-CustomConfigurationHelp
#>

<#
	.BASEPARAMETERS Silent
	
	.TODO
#>
#Requires -Version 5.0
#Requires -Modules Varan.PowerShell.Base
#Requires -Modules Varan.PowerShell.Common
#Requires -Modules Varan.PowerShell.Validation
using module Varan.PowerShell.Validation
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'None')]
param (	
	  )
DynamicParam { Build-BaseParameters }

Begin
{	
	Write-LogTrace "Execute: $(Get-RootScriptName)"
	$minParams = Get-MinimumRequiredParameterCount -CommandInfo (Get-Command $MyInvocation.MyCommand.Name)
	$cmd = @{}

	if(Get-BaseParamHelpFull) { $cmd.HelpFull = $true }
	if((Get-BaseParamHelpDetail) -Or ($PSBoundParameters.Count -lt $minParams)) { $cmd.HelpDetail = $true }
	if(Get-BaseParamHelpSynopsis) { $cmd.HelpSynopsis = $true }
	
	if($cmd.Count -gt 1) { Write-DisplayHelp -Name "$(Get-RootScriptPath)" -HelpDetail }
	if($cmd.Count -eq 1) { Write-DisplayHelp -Name "$(Get-RootScriptPath)" @cmd }
}
Process
{
	try
	{
		$scriptCol = "$('Script'.PadRight(38, ' '))"
		$aliasCol = "$('Alias'.PadRight(16, ' '))"
		$descCol = 'Description'
		Write-DisplayHost "$scriptCol$aliasCol$descCol" -Style HelpItem
	
		Get-ChildItem $PSScriptRoot -Filter "*.ps1" |  
			Sort-Object { $_.BaseName } | 
			Foreach-Object -Process { 
				Write-DisplayHelp -Name $_.FullName -HelpSynopsis -DontExit
		} 

	}
	catch [System.Exception]
	{
		Write-DisplayError $PSItem.ToString() -Exit
	}
}
End
{
}