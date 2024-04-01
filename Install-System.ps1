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

foreach ($pkg in $wingetPackages) {
	Write-Host "Installing '$pkg' via winget..."
    winget install $pkg --accept-package-agreements --accept-source-agreements
}

Write-Host ""
foreach ($url in $downloadUrls) {
    Write-Host "Fetching filename for '$url'..."

    # Use Invoke-WebRequest to get the response headers
    $response = Invoke-WebRequest -Uri $url -Method Head

    $filename = $null
    if ($response.Headers["Content-Disposition"]) {
        # Extract the filename from the Content-Disposition header
        $contentDisposition = $response.Headers["Content-Disposition"]
        $filename = [System.Text.RegularExpressions.Regex]::Match($contentDisposition, 'filename="([^"]+)"').Groups[1].Value
    }

    $downloadPath = Join-Path $localStoragePath $filename
	
	Write-Host "Downloading '$filename' to '$downloadPath'..."
    Invoke-WebRequest -Uri $url -OutFile "$downloadPath"
}

Write-Host ""
foreach ($installer in $localInstallerPaths) {
	$installerPath = Join-Path $localStoragePath $installer
	
	Write-Host "Running '$installerPath'..."
    Start-Process -FilePath "$installerPath" -Wait
	#Start-Process -FilePath "$installerPath" -Args "/S" -Wait
}

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