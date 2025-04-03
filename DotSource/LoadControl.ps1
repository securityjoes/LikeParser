# Check 1 - Google Chrome
#region

# Chrome path
$ChromePath = "$SystemDrive\Program Files\Google\Chrome\Application\chrome.exe"

# if statment to test-path of Chrome
if (Test-Path -Path $ChromePath) {
    # Do nothing
}

else {
    Write-Output "[!] Chrome is not installed on this system."
    Write-Output "[!] LikeParser requires Chrome to function properly."
    Write-Output "[!] Script execution has been canceled."
    $LoadControlFlag = "fail"
}


#endregion

# Check 2 - Selenium PowerShell Module
#region
# check if selenium is installed on the system
$ModuleCheck = Get-Module -Name Selenium -ListAvailable

# if statment to check the output of the execution
if ($ModuleCheck) {
    # Do nothing
}

else {
    # Install Selenium module
    Install-Module -Name Selenium -Force
    Write-Output "[!] The Selenium PowerShell module was not detected and has now been successfully installed."
    Write-Output "[!] If the script does not function as expected, please restart your computer and try again."

}

# Import the Selenium module
Import-Module -Name Selenium
#endregion

# Check 3 - ImportExcel PowerShell Module
#region
# check if ImportExcel is installed on the system
$ModuleCheck = Get-Module -Name ImportExcel -ListAvailable

# if statment to check the output of the execution
if ($ModuleCheck) {
    # Do nothing
}

else {
    # Install ImportExcel module
    Install-Module -Name ImportExcel -Force
    Write-Output "[!] The ImportExcel PowerShell module was not detected and has now been successfully installed."
    Write-Output "[!] If the script does not function as expected, please restart your computer and try again."

}

# Import the ImportExcel module
Import-Module -Name ImportExcel
#endregion

# Check Chrome browser versions
$ChromeVersion = (Get-Item "$SystemDrive\Program Files\Google\Chrome\Application\chrome.exe").VersionInfo.ProductVersion
$ChromeVersion = $ChromeVersion -replace '\..*',''

# Check Chrome driver version
$ChromeDriveVersion = (Get-Item "$RunPath\ChromeDriverWin64\chromedriver.exe").VersionInfo.FileVersion
$ChromeDriveVersion = $ChromeDriveVersion -replace '\..*',''

# if version are the same do:
if ($ChromeVersion -eq $ChromeDriveVersion) {
    # do nothing
}

# if versions are not the same do:
else {
    Write-Output ""
    Write-Output "[!] Chrome browser version is $ChromeVersion but your Chrome driver version is $ChromeDriveVersion"
    Write-Output "[!] Please download the matching Chrome driver for your Chrome browser"
    Write-Output "[!] And save it under $RunPath\ChromeDriverWin64\"
    Write-Output "[!] URL to download: https://getwebdriver.com/"
    Write-Output ""
    $LoadControlFlag = "fail"


}