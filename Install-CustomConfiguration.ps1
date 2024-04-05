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

	.PARAMETER LocalDownloadDir
	The directory containing the local storage for downloads. If not provided, "downloads" is assumed.
	
	.PARAMETER LocalPackageDir
	The directory containing the local storage for pre-made packages. If not provided, "packages" is assumed.

	.PARAMETER LocalFontsDir
	The directory containing the local fonts. If not provided, "fonts" is assumed.

	.PARAMETER DefaultCategoryName
	The default category name that is always installed. If not provided, "All" is assumed.
	
	.PARAMETER SkipWinget
	Skips winget installs.
	
	.PARAMETER SkipDownloads
	Skips downloading files.
	
	.PARAMETER SkipLocalInstalls
	Skips local installs.
	
	.PARAMETER SkipHosts
	Skips updating the HOSTS file.
	
	.PARAMETER SkipFonts
	Skips font installs.
	
	.PARAMETER SkipCopies
	Skips copies.
	
	.PARAMETER SkipRunCommands
	Skips running commands.
	
	.PARAMETER SkipStartups
	Skips enabling/disabling startup programs.
	
	.PARAMETER SkipLibraries
	Skips creating libraries.
	
	.PARAMETER SkipShortcuts
	Skips creating shortcuts.
	
	.PARAMETER DeleteDownloads
	If set, deletes downloaded files when installation is done.

	.PARAMETER ExcludeDefaultCategory
	Normally, items in the default category are always processed for any category. If this parameter is passed, they will not be processed. Cannot be used together with the -All parameter.
	
	.PARAMETER Interactive
	If set, presents a menu to run the different processes instead of running them automatically.
	
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
		[parameter()]											        [string]$ConfigurationPath,
		[parameter()]											        [string]$WindowsDir,
		[parameter()]											        [string]$ConfigDownloadDir,
		[parameter()]											        [string]$ConfigPackageDir,
		[parameter()]											        [string]$ConfigFontDir,
		[parameter()]											        [string]$DefaultCategoryName,
		[parameter(ParameterSetName='Category')]       					[switch]$ExcludeDefaultCategory,
		[parameter()]													[switch]$SkipWinget = $false,
		[parameter()]													[switch]$SkipDownloads = $false,
		[parameter()]													[switch]$SkipLocalInstalls = $false,
		[parameter()]													[switch]$SkipHosts = $false,
		[parameter()]													[switch]$SkipFonts = $false,
		[parameter()]													[switch]$SkipCopies = $false,
		[parameter()]													[switch]$SkipRunCommands = $false,
		[parameter()]													[switch]$SkipStartups = $false,
		[parameter()]													[switch]$SkipLibraries = $false,
		[parameter()]													[switch]$SkipShortcuts = $false,
		[parameter()]													[switch]$DeleteDownloads = $false,
		[parameter()]													[switch]$Interactive = $false
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

		function Import-CategorizedData {
			param (
				[string]$FilePath
			)

			$categories = @{}
			$currentCategory = $null
			$lines = Get-Content -Path $FilePath

			# first pass - to identify categories
			foreach ($line in $lines) {
				if ($line.Trim() -match '^[\S ]+:$') {
					$trimmedLine = $line.Trim()
					$categoryName = $trimmedLine -replace ':$'
					if (-not $categories.ContainsKey($categoryName)) {
						$categories[$categoryName] = @()
					}
				}
			}

			# second pass - to add items
			foreach ($line in $lines) {
				if([string]::IsNullOrWhiteSpace(($line -replace '--.*$','').Trim())) { continue }
				
				if ($line -match '^\t') {										# item line
					$trimmedLine = ($line -replace '--.*$','').Trim()			# remove any comments from end of line

					if ($currentCategory) {
						if($categories[$currentCategory] -Contains $trimmedLine) { continue }
						
						$categories[$currentCategory] += $trimmedLine			# add to current category if it doesn't already include it
						
						if ($currentCategory -ne $defaultCategoryName) {
							if($categories[$defaultCategoryName] -Contains $trimmedLine) { continue }

							$categories[$defaultCategoryName] += $trimmedLine	# add to default category if it doesn't already exist
						}
						
						if (-Not $ExcludeDefaultCategory -And ($currentCategory -eq $defaultCategoryName)) {
							$categoryKeys = $categories.Keys | Where-Object { $_ -ne $defaultCategoryName -and $_ -ne $currentCategory }
							
							foreach ($key in $categoryKeys) {
								if($categories[$key] -Contains $trimmedLine) { continue }
								$categories[$key] += $trimmedLine				# add to all other categories if it doesn't already exist
							}
						}
					}
					else {
						Write-Error "Invalid line, item not in a category: $line"
					}
				} elseif ($line.Trim() -match '^[\S ]+:$') {					# category line
					$trimmedLine = ($line -replace '--.*$','').Trim()			# remove any comments from end of line
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
					$allCategories += $categories.Keys
			}

			$sortedUniqueCategories = $allCategories | Where-Object { $_ -ne $defaultCategoryName } | Sort-Object -Unique -CaseSensitive:$false

			if ($allCategories -contains $defaultCategoryName) {
				$sortedUniqueCategories = @($defaultCategoryName) + $sortedUniqueCategories
			}

			return $sortedUniqueCategories
		}

		function ShowMenu {
			Write-Host " 1. list categories       - (l)"
			Write-Host " 2. process winget        - (w)"
			Write-Host " 3. process downloads     - (d)"
			Write-Host " 4. process installs      - (i)"
			Write-Host " 5. process hosts         - (h)"
			Write-Host " 6. process fonts         - (f)"
			Write-Host " 7. process copies        - (c)"
			Write-Host " 8. process run commands  - (r)"
			Write-Host " 9. process libraries     - (b)"
			Write-Host "10. process startups      - (s)"
			Write-Host "11. process shortcuts     - (t)"
			Write-Host ""
			Write-Host " 0. exit                  - (x)"
			Write-Host ""
			
			$choice = Read-Host "select option> "
			
			if($choice -eq "1" -or $choice.ToLower() -eq "l") {
				ListCategories
			}

			if($choice -eq "2" -or $choice.ToLower() -eq "w") {
				ProcessWinget
			}
			
			if($choice -eq "3" -or $choice.ToLower() -eq "d") {
				ProcessDownloads
			}
			
			if($choice -eq "4" -or $choice.ToLower() -eq "i") {
				ProcessInstalls
			}
			
			if($choice -eq "5" -or $choice.ToLower() -eq "h") {
				ProcessHosts
			}
			
			if($choice -eq "6" -or $choice.ToLower() -eq "f") {
				ProcessFonts
			}
			
			if($choice -eq "7" -or $choice.ToLower() -eq "c") {
				ProcessCopies
			}
			
			if($choice -eq "8" -or $choice.ToLower() -eq "r") {
				ProcessRunCommands
			}
			
			if($choice -eq "9" -or $choice.ToLower() -eq "s") {
				ProcessStartups
			}
			
			if($choice -eq "10" -or $choice.ToLower() -eq "b") {
				ProcessLibraries
			}
			
			if($choice -eq "11" -or $choice.ToLower() -eq "t") {
				ProcessShortcuts
			}
			
			if($choice -eq "0" -or $choice.ToLower() -eq "x") {
				exit
			}
			
			ShowMenu
		}
		
		function ListCategories {
			if (-Not $Interactive -And (-Not $ListCategories)) { return }

			Write-Host ""
			Write-Host "Categories:"
			$categoryList | ForEach-Object { Write-Output "  $_" }
			
			Write-Host ""
			Write-DisplayHost "Done." -Style Done
			
			if(-Not $Interactive) { exit }
		}
		
		function ProcessWinget {
			$do = $Interactive -or ($doWinget -And $wingetPackages -And $wingetPackages[$categoryToUse] -And $wingetPackages[$categoryToUse].Count -gt 0)
			Write-Host "do = $do"
			
			if(-Not $do) { return }

			Write-Host ""
			Write-host "Performin winget installs"
			
			foreach ($pkg in $wingetPackages[$categoryToUse]) {
				Write-Host "Installing '$pkg' via winget..."
				winget install $pkg --accept-package-agreements --accept-source-agreements --include-unknown
			}
		}
		
		function ProcessDownloads {
			$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"

			$do = $Interactive -or ($doDownload -And $downloadUrls -And $downloadUrls[$categoryToUse] -And $downloadUrls[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }

			Write-Host ""
			Write-Host "Performing downloads"
			
			foreach ($line in $downloadUrls[$categoryToUse]) {
				Write-Host ""
				$splitLine = $line -split "`t"
				
				if($splitLine.Count -eq 1) {
					$downloadName = $splitLine[0]
					$folderName = "DropboxUnknown"
					$url = $splitLine[0]
				}
				else {
					$downloadName = $splitLine[0]
					$folderName = $downloadName
					$url = $splitLine[1]
				}
				Write-Host "`Fetching filename for '$downloadName'..."

				$filename = $null
				$response = $null

				 if ($url -like "*dropbox.com*") {
					Write-Host "Dropbox folder detected, modifying URL for direct download..."

					$url = $url -replace 'dl=0', 'dl=1'
					$filename = "$folderName.zip"
				}
				
				# Try to get the filename from the Content-Disposition header using a HEAD request
				if (-not $filename) {
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

				$downloadPath = Join-Path $configDownloadPath $filename

				Write-Host "Downloading '$filename' to '$downloadPath'..."
				try {
					Invoke-WebRequest -Uri $url -OutFile $downloadPath -UserAgent $userAgent
				} catch {
					Write-Host "Error downloading '$url': $_"
				}
				
				if ($filename.EndsWith(".zip")) {
					Write-Host "Extracting '$filename'..."
					$extractPath = $downloadPath.Replace('.zip', '')
					if (Test-Path -Path $extractPath) {
						Remove-Item $extractPath -Force -Recurse
					}

					$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"
					$password = "abc123"

					$arguments = "x `"$downloadPath`" -o`"$extractPath`" -aoa"
					if ($password) {
						$arguments += " -p$password"
					}

					Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -NoNewWindow -Wait

					Write-Host "Deleting '$downloadPath'..."
					Remove-Item $downloadPath -Force
				}
			}
		}
		
		function ProcessInstalls {
			$do = $Interactive -or ($doLocalInstall -And $localInstallerPaths -And $localInstallerPaths[$categoryToUse] -And $localInstallerPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Running local installs"
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
				$downloadPath = Join-Path $configDownloadPath $installerFileName
				$packagePath = Join-Path $configPackagePath $installerFileName
				
				$installerPath = ""
				if(Test-Path -Path $downloadPath) {
					$installerPath = $downloadPath
				}
				
				if(Test-Path -Path $packagePath) {
					$installerPath = $packagePath
				}
				
				if(-Not [string]::IsNullOrWhiteSpace($installerPath)) {
					Start-Process -FilePath "$installerPath" -Wait
				}
			}
			
			if($DeleteDownloads) {
				Remove-Item -Path '$configDownloadPath\*' -Recurse -Force
			}
		}
		
		function ProcessHosts {
			$do = $Interactive -or ($doHosts -And $hostsEntries -And $hostsEntries[$categoryToUse] -And $hostsEntries[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Updating HOSTS"
			$content = Get-Content -Path $hostsPath
			
			foreach ($entry in $hostsEntries[$categoryToUse]) {	
				if ($content -notcontains $entry) {
					Write-Host "Adding host entry '$entry'..."
					Add-Content -Path $hostsPath -Value $entry
				}
			}
		}
		
		function ProcessFonts {
			$do = $Interactive -or ($doFonts -And $fontPaths -And $fontPaths[$categoryToUse] -And $fontPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Installing fonts"
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
					$fullFontPath = Join-Path $configFontPath $fontPath
					Copy-Item -Path $fullFontPath -Destination $destPath
				}
				if (-Not (Get-ItemProperty -Path $fontKeyPath -Name $regValue -ErrorAction SilentlyContinue)) {
					Set-ItemProperty -Path $fontKeyPath -Name $regValue -Value $font
				}
			}
		}
		
		function ProcessCopies {
			$do = $Interactive -or ($doCopies -And $copyPaths -And $copyPaths[$categoryToUse] -And $copyPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Copying packages"
			
			foreach ($copyPath in $copyPaths[$categoryToUse]) {
				$splitCopy = $copyPath -split "`t"
				if($splitCopy.Count -eq 2) {
					$sourceDirectory = $splitCopy[0]
					$destinationDirectory = $splitCopy[1]
				}
				elseif($splitCopy.Count -eq 3) {
					$sourceDirectory = $splitCopy[1]
					$destinationDirectory = $splitCopy[2]					
				}
				
				if([string]::IsNullOrWhiteSpace($sourceDirectory) -or [string]::IsNullOrWhiteSpace($destinationDirectory)) { 	continue 
				}
								
				$sourceDirectory = Join-Path $configPackagePath $sourceDirectory
				if(-Not (Test-Path -Path $sourceDirectory)) { continue }
				
				$sourceDirectory += "\*"
				
				if(-Not $destinationDirectory.EndsWith("\")) { $destinationDirectory += "\" }
				
				if(-Not (Test-Path -Path $destinationDirectory)) {
					Write-Host "Copying '$sourceDirectory'"
					New-Item -Type Directory $destinationDirectory
				}		
				
				Copy-item -Force -Recurse -Verbose $sourceDirectory -Destination $destinationDirectory				
			}
		}

		function ProcessRunCommands {
			$do = $Interactive -or ($doRunCommands -And $runCommandPaths -And $runCommandPaths[$categoryToUse] -And $runCommandPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
		
			Write-Host ""
			Write-Host "Running commands"
		
			foreach ($command in $runCommandPaths[$categoryToUse]) {
				$splitStart = $startItem -split "`t"
				
				if($splitStart.Count -ne 2) { continue }
				
				$name = $splitStart[0]
				$cmd = $splitStart[1]
				
				Write-Host "Executing $name"
				Invoke-Expression $cmd
			}
		}

		function ProcessStartups {
			$do = $Interactive -or ($doStartups -And $startupPaths -And $startupPaths[$categoryToUse] -And $startupPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Enabling/disabling startup programs"

			foreach ($startItem in $startupPaths[$categoryToUse]) {
				$splitStart = $startItem -split "`t"
				
				if($splitStart.Count -ne 3) { continue }
				
				$state = $splitStart[0].ToLower()
				$name = $splitStart[1]
				$exePath = $splitStart[2]
				
				if($state -ne "on" -And $state -ne "off") { continue }
				if([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($exePath)) { continue }
				
				if($state -eq "on") {
					Write-Host "Enabling $name"
					Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name '$name' -Value '$exePath'
				}
				else {
					Write-Host "Disabling $name"
					Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name '$name'
				}
			}		
		}	

		function ProcessLibraries {
			$do = $Interactive -or ($doLibraries -And $libraryPaths -And $libraryPaths[$categoryToUse] -And $libraryPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
			
			Write-Host ""
			Write-Host "Processing libraries"
			
			foreach ($libItem in $libraryPaths[$categoryToUse]) {
				$splitLib = $libItem -split "`t"
				
				if($splitLib.Count -ne 2) { continue }
				
				$name = $splitLib[0]
				$folderPath = $splitLib[1]
				
				if(-Not (Test-Path -Path $folderPath)) { continue }
				
				$shell = New-Object -ComObject Shell.Application
				$existinglibraries = $shell.NameSpace("shell:::{031E4825-7B94-4dc3-B131-E946B44C8DD5}")
				$library = $librariesFolder.Items() | Where-Object { $_.Name -eq $name }

				if (-Not $library) {
					Write-Host "Creating library $name"
					$library = $existinglibraries.Items().Add($name)
				}
				
				$folders = $shell.NameSpace($library.Path).Items()
				$folderInLibrary = $folders | Where-Object { $_.Path -eq $searchPath }

				if (-Not $folderInLibrary -And (Test-Path -Path $folderPath)) {
					Write-Host "Adding '$folderPath' to '$name'"
					
					$libraryShell = $shell.NameSpace($library.Path)
					$libraryShell.MoveHere($folderPath)
				}
			}	
		}		
		
		function ProcessShortcuts {
			$do = $Interactive -or ($doShortcuts -And $shortcutPaths -And $shortcutPaths[$categoryToUse] -And $shortcutPaths[$categoryToUse].Count -gt 0)
			
			if(-Not $do) { return }
		
			Write-Host ""
			Write-Host "Creating shortcuts"
		
			foreach ($shortcut in $shortcutPaths[$categoryToUse]) {
				$splitShortcut = $startItem -split "`t"
				
				if($splitStart.Count -ne 3) { continue }
				
				$name = $splitStart[0]
				$sourcePath = $splitStart[1]
				$targetPath = $splitStart[2]
				
				if ([string]::IsNullOrWhiteSpace($name) -or
					[string]::IsNullOrWhiteSpace($sourcePath) -or
					[string]::IsNullOrWhiteSpace($targetPath)) { 
					continue 
				}
				
				if(-Not (Test-Path -Path $sourcePath)) { continue }
				if(-Not (Test-Path -Path $targetPath)) { continue }
				
				if(-Not $targetPath.EndsWith("\")) {
					$targetPath += "\"
				}
				
				Write-Host "Creating $name"
				$WScriptShell = New-Object -ComObject WScript.Shell
				$Shortcut = $WScriptShell.CreateShortcut("$targetPath$name.lnk")
				$Shortcut.TargetPath = $sourcePath
				$Shortcut.Save()
			}
		}
		
		
		
		# ----- start config --------------------------------------------------
		$defaultWindowsPath = "C:\Windows"
		$defaultConfigPath = (Get-Location).Path
		$defaultConfigDownloadDir = "downloads"
		$defaultConfigPackageDir = "packages"
		$defaultConfigFontDir = "fonts"
		$defaultCategoryName = "All"
		
		if($WindowsDir) {
			$windowsPath = $WindowsDir
		}
		else {
			$windowsPath = $defaultWindowsPath
		}
		
		$hostsPath = Join-Path $windowsPath "System32\drivers\etc\HOSTS"
		$windowsFontsPath = Join-Path $windowsPath "Fonts"
		
		if($ConfigurationPath) {
			$configDir = $ConfigurationPath
		}
		else {
			$configDir = $defaultConfigPath
		}

		if($ConfigDownloadDir) {
			$downloadDir = $ConfigDownloadDir
		}
		else {
			$downloadDir = $defaultConfigDownloadDir
		}
		
		if($ConfigPackageDir) {
			$packageDir = $ConfigPackageDir
		}
		else {
			$packageDir = $defaultConfigPackageDir
		}
		
		if($ConfigFontDir) {
			$fontsDir = $ConfigFontDir
		}
		else {
			$fontsDir = $defaultConfigFontDir
		}
		
		if($DefaultCategory) {
			$defaultCategoryName = $DefaultCategory
		}
		else {
			$defaultCategoryName = $defaultCategoryName
		}
		
		if($Category) {
			$categoryToUse = $Category
		}
		
		if($All) {
			$categoryToUse = $defaultCategoryName
		}
		
		$configFontPath = Join-Path $configDir $fontsDir
		$configPackagePath = Join-Path $configDir $packageDir
		$configDownloadPath = Join-Path $configDir $downloadDir
		$fontKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
		
		if(-Not (Test-Path -PathType Container -Path $configDownloadPath)) {
			New-Item -ItemType Directory $configDownloadPath | Out-Null
		}
		
		$doWinget = -Not $SkipWinget
		$doDownload = -Not $SkipDownloads
		$doLocalInstall = -Not $SkipLocalInstalls
		$doHosts = -Not $SkipHosts
		$doFonts = -Not $SkipFonts
		$doCopies = -Not $SkipCopies
		$doRunCommands = -Not $SkipRunCommands
		$doStartups = -Not $SkipStartups
		$doLibraries = -Not $SkipLibraries
		$doShortcuts = -Not $SkipShortcuts
		
		Write-Host "---------------------------"
		Write-Host "Configuration settings:"
		Write-Host "  Windows directory       = '$windowsPath'"
		Write-Host "  configuration directory = '$configDir'"
		Write-Host "  config download path    = '$configDownloadPath'"
		Write-Host "  config package path     = '$configPackagePath'"
		Write-Host "  config fonts path       = '$configFontPath'"
		Write-Host "  default category name   = $defaultCategoryName"
		Write-Host "  interactive mode        = $Interactive".Replace("True", "yes").Replace("False", "no")
		if(-Not $Interactive) {
			Write-Host ("  perform winget          = $doWinget").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform downloads       = $doDownload").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform local installs  = $doLocalInstall").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform hosts           = $doHosts").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform fonts           = $doFonts").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform copies          = $doCopies").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform run commands    = $doRunCommands").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform startups        = $doStartups").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform libraries       = $doLibraries").Replace("True", "yes").Replace("False", "no")
			Write-Host ("  perform shortcuts       = $doShortcuts").Replace("True", "yes").Replace("False", "no")
		}
		if(-Not $ListCategories) {
			Write-Host "  processing for category = $categoryToUse"
		}
		Write-Host "---------------------------"
		
		if($packageDir -eq $downloadDir) {
			Write-Error "Package directory and download directory cannot be the same."
			exit
		}
		
		$wingetPackages = Import-CategorizedData "$configDir\wingetPackages.txt"
		$downloadUrls = Import-CategorizedData "$configDir\downloadUrls.txt"
		$localInstallerPaths = Import-CategorizedData "$configDir\localInstallers.txt"
		$hostsEntries = Import-CategorizedData "$configDir\hostsEntries.txt"
		$fontPaths = Import-CategorizedData "$configDir\fonts.txt"
		$copyPaths = Import-CategorizedData "$configDir\copy.txt"
		$runCommandPaths = Import-CategorizedData "$configDir\runCommands.txt"
		$startupPaths = Import-CategorizedData "$configDir\startups.txt"
		$libraryPaths = Import-CategorizedData "$configDir\libraries.txt"
		$shortcutsPaths = Import-CategorizedData "$configDir\shortcuts.txt"

		# ----- end config --------------------------------------------------

		$categoryList = Get-UniqueCategories @("$configDir\wingetPackages.txt", 
								 "$configDir\downloadUrls.txt", 
								 "$configDir\localInstallers.txt", 
								 "$configDir\hostsEntries.txt", 
								 "$configDir\fonts.txt",
								 "$configDir\copy.txt",
								 "$configDir\runCommands.txt",
								 "$configDir\startups.txt",
								 "$configDir\libraries.txt",
								 "$configDir\shortcuts.txt")
									 
		
		if($Interactive) {
			ShowMenu
		}
		
		ListCategories
		
		if(-Not ($categoryList -Contains $categoryToUse)) {
			Write-Host "Category '$categoryToUse' not present in any of the config files. Nothing to do."
		}
		
		if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
			Write-Host "This script needs to be run as Administrator." -ForegroundColor Red
			exit
		}
		
		ProcessWinget
		ProcessDownloads
		ProcessInstalls
		ProcessHosts
		ProcessFonts
		ProcessCopies
		ProcessRunCommands
		ProcessStartups
		ProcessLibraries
		ProcessShortcuts
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