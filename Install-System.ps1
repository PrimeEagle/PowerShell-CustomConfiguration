# ----- start config --------------------------------------------------

$windowsPath = "C:\Windows"
$hostsPath = Join-Path $windowsPath "System32\drivers\etc\HOSTS"
$windowsFontsPath = Join-Path $windowsPath "Fonts"
$rootPath = "D:\systemsetup"
$localStoragePath = Join-Path $rootPath "downloads"
$localFontsPath = Join-Path $rootPath "Fonts"
$fontKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

$wingetPackages = Get-Content -Path "$rootPath\wingetPackages.txt" | ForEach-Object { $_.Trim() }
$downloadUrls = Get-Content -Path "$rootPath\downloadUrls.txt" | ForEach-Object { $_.Trim() }
$localInstallerPaths = Get-Content -Path "$rootPath\localInstallers.txt" | ForEach-Object { $_.Trim() }
$hostsEntries = Get-Content -Path "$rootPath\hostsEntries.txt" | ForEach-Object { $_.Trim() }
$fontPaths = Get-Content -Path "$rootPath\fonts.txt"

# ----- end config --------------------------------------------------


if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "This script needs to be run as Administrator." -ForegroundColor Red
    exit
}

#foreach ($pkg in $wingetPackages) {
#	Write-Host "Installing '$pkg' via winget..."
#    winget install $pkg --accept-package-agreements --accept-source-agreements
#}

#$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
#
#foreach ($url in $downloadUrls) {
#    Write-Host "`nFetching filename for '$url'..."
#
#    $filename = $null
#    $response = $null
#
#    # Try to get the filename from the Content-Disposition header using a HEAD request
#    try {
#        $response = Invoke-WebRequest -Uri $url -Method Head -UserAgent $userAgent
#        if ($response.Headers["Content-Disposition"]) {
#            $contentDisposition = $response.Headers["Content-Disposition"]
#            $matches = [System.Text.RegularExpressions.Regex]::Match($contentDisposition, 'filename="?([^"]+)"?')
#            if ($matches.Success) {
#                $filename = $matches.Groups[1].Value
#            }
#        }
#    } catch {
#        Write-Host "HEAD request failed: $_"
#    }
#
#    # If not found, try to get the filename from the URL after potential redirection
#    if (-not $filename) {
#        try {
#            # Use a GET request but don't download the body
#            $response = Invoke-WebRequest -Uri $url -UserAgent $userAgent -Method Get -MaximumRedirection 5 -UseBasicParsing
#            $finalUrl = $response.BaseResponse.ResponseUri.AbsoluteUri
#
#            if ($response.Headers["Content-Disposition"]) {
#                $contentDisposition = $response.Headers["Content-Disposition"]
#                $matches = [System.Text.RegularExpressions.Regex]::Match($contentDisposition, 'filename="?([^"]+)"?')
#                if ($matches.Success) {
#                    $filename = $matches.Groups[1].Value
#                }
#            } elseif ($finalUrl) {
#                $filename = [System.IO.Path]::GetFileName((New-Object System.Uri($finalUrl)).LocalPath)
#            }
#        } catch {
#            Write-Host "GET request failed: $_"
#        }
#    }
#
#    # Fallback to using the original URL if no filename has been determined
#    if (-not $filename -or $filename -eq '') {
#        $uri = [System.Uri]$url
#        $filename = [System.IO.Path]::GetFileName($uri.LocalPath)
#        if (-not $filename -or $filename -eq '') {
#            Write-Host "Could not determine filename, using default name."
#            $filename = "defaultFilename.exe"
#        }
#    }
#
#    $downloadPath = Join-Path $localStoragePath $filename
#
#    Write-Host "Downloading '$filename' to '$downloadPath'..."
#    try {
#        Invoke-WebRequest -Uri $url -OutFile $downloadPath -UserAgent $userAgent
#    } catch {
#        Write-Host "Error downloading '$url': $_"
#    }
#}

Write-Host ""
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

foreach ($installLine in $localInstallerPaths) {
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

    Write-Host "name = $name"
    Write-Host "already installed = $alreadyInstalled"
    
    if ($alreadyInstalled) { 
        Write-Host "skipping '$name'"
        continue 
    }
    
    Write-Host "Running '$installerPath'..."
	$installerPath = Join-Path $localStoragePath $installerFileName
    Start-Process -FilePath "$installerPath" -Wait
}


exit
Write-Host ""
foreach ($entry in $hostsEntries) {
    $content = Get-Content -Path $hostsPath
    if ($content -notcontains $entry) {
		Write-Host "Adding host entry '$entry'..."
        Add-Content -Path $hostsPath -Value $entry
    }
}

Write-Host ""
foreach ($fontPath in $fontPaths) {
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

Write-Host ""
Write-Host "Done."