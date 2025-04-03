# Check 1 - 1Password Application
#region

# 1Password path
$1PasswordPath = "$env:LocalAppData\1Password"

if (Test-Path -Path $1PasswordPath) {
    # Do nothing
}

else {
    Write-Output "[!] 1Password is not installed on this system."
    Write-Output "[!] LikeParser requires 1Password to function properly."
    Write-Output "[!] Script execution has been canceled."
    $LoadControlFlag = "fail"
    exit
}
#endregion

# Check 2 - 1Password CLI 
#region

# Catch and store response
$Respond = winget list 1password-cli

# if statment to check if not installed
if ($Respond -eq "No installed package found matching input criteria.") {
    
    # Install 1password-cli
    winget install 1password-cli | Out-Null
    Write-Output "[!] 1Password CLI was not detected and has now been successfully installed."
    Write-Output "[!] If the script does not function as expected, please restart your computer and try again."
    Write-Output ""


}
#endregion

# Currect DotSource variables
$LinkedInLoginURL = "https://www.linkedin.com/login"

# Catch output in a variable
op user list *> "$RunPath\temp.txt"
$CheckLogin = Get-Content -Path "$RunPath\temp.txt"

# User is not logined
if ($CheckLogin -match ".*\[ERROR\].*") {
    Invoke-Expression $(op signin)
    $Email = op item get $ItemName --fields label=username
    $Password = op item get $ItemName --fields label=password
}

# User is logined
else {
    $Email = op item get $ItemName --fields label=username
    $Password = op item get $ItemName --fields label=password
}

# Remove the temp.txt file
Remove-Item -Path "$RunPath\temp.txt" -ErrorAction SilentlyContinue -Force | Out-Null

# Open Google Chrome
$global:Driver = Start-SeChrome -BinaryPath $ChromePath -WebDriverDirectory $WebDriverPath -StartURL $LinkedInLoginURL

# Find "Email or Phone" input tab and inject $Email
Send-SeKeys -Element (Find-SeElement -Driver $Driver -id "username") -Keys $Email

# Find "Password" input tab and inject $Password
Send-SeKeys -Element (Find-SeElement -Driver $Driver -id "password") -Keys $Password

# Click on "Sign in"
Start-Sleep -Seconds 1.5
Invoke-SeClick -Element (Find-SeElement -Driver $Driver -XPath '/html/body/div/main/div[2]/div[1]/form/div[4]/button')