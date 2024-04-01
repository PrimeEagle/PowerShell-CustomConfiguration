<#
	.SYNOPSIS
	Installs a custom configuration.
	
	.DESCRIPTION
	Installs a custom configuration.

	.PARAMETER ListCategories
	List all available categories. No installation is performed.

	.PARAMETER Category
	Perform installation for the specified category.
	
	.PARAMETER All
	Perform installation for all categories.
	
	.PARAMETER ConfigurationDir
	The directory containing the configuration text files. If not provided, current directory is assumed.
	
	.PARAMETER WindowsDir
	The directory containing Windows. If not provided, "C:\Windows" is assumed.

	.PARAMETER LocalStorageDir
	The directory containing the local storage. If not provided, "downloads" is assumed.

	.PARAMETER LocalFontsDir
	The directory containing the local fonts. If not provided, "fonts" is assumed.

	.PARAMETER DefaultCategory
	The default category name that is always installed. If not provided, "All" is assumed.
	
	.INPUTS
	Category to install.

	.OUTPUTS
	None.

	.EXAMPLE
	PS> .\Install-CustomConfiguration.ps1 -All
#>
using module Varan.PowerShell.Validation
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Requires -Version 5.0
#Requires -Modules Varan.PowerShell.Validation
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (	
		[parameter(Mandatory=$true, ParameterSetName='ListCategories')]	[switch]$ListCategories,
		[parameter(Mandatory=$true, ParameterSetName='Category')]	    [string]$Category,
		[parameter(Mandatory=$true, ParameterSetName='All')]	        [switch]$All,
		[parameter()]											        [string]$ConfigurationDir,
		[parameter()]											        [string]$WindowsDir,
		[parameter()]											        [string]$LocalStorageDir,
		[parameter()]											        [string]$LocalFontsDir,
		[parameter()]											        [string]$DefaultCategory
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
		$isDebug = Assert-Debug

		# ----- start config --------------------------------------------------
		$defaultWindowsDir = "C:\Windows"
		$defaultConfigDir = (Get-Location).Path
		$defaultLocalStorageDir = "downloads"
		$defaultLocalFontsDir = "fonts"
		$defaultLocalCategoryName = "All"
		
		if($WindowsDir) {
			$windowsPath = $WindowsDir
		}
		else {
			$windowsPath = $defaultWindowsDir
		}
		
		$hostsPath = Join-Path $windowsPath "System32\drivers\etc\HOSTS"
		$windowsFontsPath = Join-Path $windowsPath "Fonts"
		
		if($ConfigurationDir) {
			$configDir = $ConfigurationDir
		}
		else {
			$configDir = $defaultConfigDir
		}
		
		if($LocalStorageDir) {
			$storageDir = $LocalStorageDir
		}
		else {
			$storageDir = $defaultLocalStorageDir
		}
		
		if($LocalFontsDir) {
			$fontsDir = $LocalFontsDir
		}
		else {
			$fontsDir = $defaultLocalFontsDir
		}
		
		if($DefaultCategory) {
			$defaultCategoryName = $DefaultCategory
		}
		else {
			$defaultCategoryName = $defaultLocalCategoryName
		}
		
		if($Category) {
			$categoryToUse = $Category
		}
		
		if($All) {
			$categoryToUse = $defaultCategoryName
		}
		
		$localStoragePath = Join-Path $configDir $storageDir
		$localFontsPath = Join-Path $configDir $fontsDir
		
		$fontKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
		
		Write-Host "Using Windows directory '$windowsPath'"
		Write-Host "Using configuration directory '$configDir'"
		Write-Host "Using local storage directory '$localStoragePath'"
		Write-Host "Using local fonts directory '$localFontsPath'"
		Write-Host "Using default category name '$defaultCategoryName'"
		
		if(-Not $ListCategories) {
			Write-Host "Processing for category '$categoryToUse'"
		}

		function Import-CategorizedData {
			param (
				[string]$FilePath
			)

			$categories = @{}
			$currentCategory = $null
			$lines = Get-Content -Path $FilePath

			# First pass to identify categories
			foreach ($line in $lines) {
				if ($line.Trim() -match '^\S+:$') {
					$trimmedLine = $line.Trim()
					$categoryName = $trimmedLine -replace ':$'
					if (-not $categories.ContainsKey($categoryName)) {
						$categories[$categoryName] = @()
					}
				}
			}

			# Second pass to add items
			foreach ($line in $lines) {
				#Write-Host "reading line '$line'"
				if ($line -match '^\t') {
					$trimmedLine = $line.Trim()
					if ($currentCategory) {
						$categories[$currentCategory] += $trimmedLine
						#Write-Host "   adding to '$currentCategory'"
						if ($currentCategory -ne $defaultCategoryName) {
							# Create a list of keys to iterate over to avoid modifying the collection directly
							$categoryKeys = $categories.Keys | Where-Object { $_ -ne $defaultCategoryName -and $_ -ne $currentCategory }
						}
						
						if ($currentCategory -eq $defaultCategoryName) {
							foreach ($key in $categoryKeys) {
								$categories[$key] += $trimmedLine
								#Write-Host "   adding to '$key'"
							}
						}
					}
					else {
						Write-Error "Invalid line, item not in a category: $line"
					}
				} elseif ($line.Trim() -match '^\S+:$') {
					$trimmedLine = $line.Trim()
					$currentCategory = $trimmedLine -replace ':$'
				} elseif ($line.Trim() -and !$line.StartsWith("--")) {
					Write-Error "Invalid line format: $line"
				}
			}

			return $categories
		}

		function Get-UniqueCategories {
			param (
				[string[]]$FileNames
			)

			$allCategories = @()

			foreach ($file in $FileNames) {
					$categories = Import-CategorizedData -FilePath $file
					
					# Add the keys from each categories hash table to the allCategories array
					$allCategories += $categories.Keys
			}

			# Filter out the default category, sort and select unique category names
			$sortedUniqueCategories = $allCategories | Where-Object { $_ -ne $defaultCategoryName } | Sort-Object -Unique -CaseSensitive:$false

			# Add the default category back at the start if it was originally present
			if ($allCategories -contains $defaultCategoryName) {
				$sortedUniqueCategories = @($defaultCategoryName) + $sortedUniqueCategories
			}

			return $sortedUniqueCategories
		}

		
		$categoryList = Get-UniqueCategories @("$configDir\wingetPackages.txt", 
											 "$configDir\downloadUrls.txt", 
											 "$configDir\localInstallers.txt", 
											 "$configDir\hostsEntries.txt", 
											 "$configDir\fonts.txt")
							
		$wingetPackages = Import-CategorizedData "$configDir\wingetPackages.txt"
		$downloadUrls = Import-CategorizedData "$configDir\downloadUrls.txt"
		$localInstallerPaths = Import-CategorizedData "$configDir\localInstallers.txt"
		$hostsEntries = Import-CategorizedData "$configDir\hostsEntries.txt"
		$fontPaths = Import-CategorizedData "$configDir\fonts.txt"

		# ----- end config --------------------------------------------------

		if($ListCategories) {
			Write-Host ""
			Write-Host "Categories:"
			$categoryList | ForEach-Object { Write-Output "  $_" }
			exit
		}
		
		if(-Not ($categoryList -Contains $categoryToUse)) {
			Write-Host "Category '$categoryToUse' not present in any of the config files. Nothing to do."
		}
		
		if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Host "This script needs to be run as Administrator." -ForegroundColor Red
			exit
		}
		
		if($wingetPackages -And $wingetPackages[$categoryToUse] -And $wingetPackages[$categoryToUse].Length -gt 0) { Write-Host "" }
		foreach ($pkg in $wingetPackages[$categoryToUse]) {
			Write-Host "Installing '$pkg' via winget..."
			winget install $pkg --accept-package-agreements --accept-source-agreements
		}

		$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"

		if($downloadUrls -And $downloadUrls[$categoryToUse] -And $downloadUrls[$categoryToUse].Length -gt 0) { Write-Host "" }
		foreach ($url in $downloadUrls[$categoryToUse]) {
			Write-Host "`nFetching filename for '$url'..."

			$filename = $null
			$response = $null

			# Try to get the filename from the Content-Disposition header using a HEAD request
			try {
				$response = Invoke-WebRequest -Uri $url -Method Head -UserAgent $userAgent
				if ($response.Headers["Content-Disposition"]) {
					$contentDisposition = $response.Headers["Content-Disposition"]
					$matches = [System.Text.RegularExpressions.Regex]::Match($contentDisposition, 'filename="?([^"]+)"?')
					if ($matches.Success) {
						$filename = $matches.Groups[1].Value
					}
				}
			} catch {
				Write-Host "HEAD request failed: $_"
			}

			# If not found, try to get the filename from the URL after potential redirection
			if (-not $filename) {
				try {
					# Use a GET request but don't download the body
					$response = Invoke-WebRequest -Uri $url -UserAgent $userAgent -Method Get -MaximumRedirection 5 -UseBasicParsing
					$finalUrl = $response.BaseResponse.ResponseUri.AbsoluteUri

					if ($response.Headers["Content-Disposition"]) {
						$contentDisposition = $response.Headers["Content-Disposition"]
						$matches = [System.Text.RegularExpressions.Regex]::Match($contentDisposition, 'filename="?([^"]+)"?')
						if ($matches.Success) {
							$filename = $matches.Groups[1].Value
						}
					} elseif ($finalUrl) {
						$filename = [System.IO.Path]::GetFileName((New-Object System.Uri($finalUrl)).LocalPath)
					}
				} catch {
					Write-Host "GET request failed: $_"
				}
			}

			# Fallback to using the original URL if no filename has been determined
			if (-not $filename -or $filename -eq '') {
				$uri = [System.Uri]$url
				$filename = [System.IO.Path]::GetFileName($uri.LocalPath)
				if (-not $filename -or $filename -eq '') {
					Write-Host "Could not determine filename, using default name."
					$filename = "defaultFilename.exe"
				}
			}

			$downloadPath = Join-Path $localStoragePath $filename

			Write-Host "Downloading '$filename' to '$downloadPath'..."
			try {
				Invoke-WebRequest -Uri $url -OutFile $downloadPath -UserAgent $userAgent
			} catch {
				Write-Host "Error downloading '$url': $_"
			}
		}

		if($localInstallerPaths -And $localInstallerPaths[$categoryToUse] -And $localInstallerPaths[$categoryToUse].Length -gt 0) { Write-Host "" }
		$registryPaths = @(
			'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
			'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
			'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		)

		$installedItems = foreach ($path in $registryPaths) {
			Get-ChildItem $path | ForEach-Object {
				$property = Get-ItemProperty $_.PsPath
				if (-not [string]::IsNullOrWhiteSpace($property.DisplayName)) {
					$property.DisplayName
				}
			}
		}

		$installedItems = $installedItems | Sort-Object -Unique

		foreach ($installLine in $localInstallerPaths[$categoryToUse]) {
			$split = $installLine -Split "`t"  # Using backtick t for tab character
			if($split.Length -eq 1) {
				$name = ""
				$installerFileName = $split[0].Trim()
			}
			else {
				$name = $split[0].Trim()
				$installerFileName = $split[1].Trim()
			}
			
			$alreadyInstalled = $false
			if (-not [string]::IsNullOrWhiteSpace($name)) {
				foreach ($item in $installedItems) {
					if ($item -match [regex]::Escape($name)) {
						$alreadyInstalled = $true
						break
					}
				}
			}
			
			if ($alreadyInstalled) { 
				Write-Host "'$name' already installed."
				continue 
			}
			
			Write-Host "Running '$installerPath'..."
			$installerPath = Join-Path $localStoragePath $installerFileName
			Start-Process -FilePath "$installerPath" -Wait
		}

		if($hostEntries -And $hostEntries[$categoryToUse] -And $hostEntries[$categoryToUse].Length -gt 0) { Write-Host "" }
		foreach ($entry in $hostsEntries[$categoryToUse]) {
			$content = Get-Content -Path $hostsPath
			if ($content -notcontains $entry) {
				Write-Host "Adding host entry '$entry'..."
				Add-Content -Path $hostsPath -Value $entry
			}
		}

		if($fontPaths -And $fontPaths[$categoryToUse] -And $fontPaths[$categoryToUse].Length -gt 0) { Write-Host "" }
		foreach ($fontPath in $fontPaths[$categoryToUse]) {
			$font = [System.IO.Path]::GetFileName($fontPath)
			$destPath = Join-Path $windowsFontsPath $font
			$fontName = [System.IO.Path]::GetFileNameWithoutExtension($font)
			$fontExt = [System.IO.Path]::GetExtension($font)
			$fontType = "TrueType"
			
			if ($fontExt -eq ".otf") {
				$fontType = "OpenType"
			}
			
			$regValue = "$fontName ($fontType)"

			Write-Host "Installing $fontType font '$fontName'..."

			if (-Not (Test-Path $destPath)) {
				$fullFontPath = Join-Path $localFontsPath $fontPath
				Copy-Item -Path $fullFontPath -Destination $destPath
			}
			if (-Not (Get-ItemProperty -Path $fontKeyPath -Name $regValue -ErrorAction SilentlyContinue)) {
				Set-ItemProperty -Path $fontKeyPath -Name $regValue -Value $font
			}
		}
	}
	catch [System.Exception]
	{
		Write-DisplayError $PSItem.ToString() -Exit
	}
}
End
{
	Write-DisplayHost "Done." -Style Done
}