# Check if the current PS session is running as an administrator
$IsAdmin = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$IsAdmin = $IsAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# if statment to check if result is true or false
if ($IsAdmin -eq $true) {
    # Tool can continue working.
}

else {
    Write-Output ""
    Write-Output "[!] LikeParser requires Administrator privileges to function."
    Write-Output "[!] Please run LikeParser from an elevated Administrator terminal."
    Write-Output ""
    exit
}

# First variable load
$RunPath = Get-Location
$Username = $env:USERNAME
$SystemDrive = $env:SystemDrive
$WebDriverPath = "$RunPath\ChromeDriverWin64\"
$ChromePath = "$SystemDrive\Program Files\Google\Chrome\Application\chrome.exe"

# Dotsource ReadConfig.ps1
. "$RunPath\DotSource\ReadConfig.ps1"

# Dotsource DownloadConvertExcel.ps1
. "$RunPath\DotSource\DownloadConvertExcel.ps1"

# Dotsource ExtractURLs.ps1
. "$RunPath\DotSource\ExtractURLs.ps1"

# Dotsource ForeachURL.ps1
. "$RunPath\DotSource\ForeachURL.ps1"

# Dotsource StartFiltering.ps1
. "$RunPath\DotSource\StartFiltering.ps1"