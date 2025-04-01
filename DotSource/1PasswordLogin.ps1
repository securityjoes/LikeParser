# Currect DotSource variables
$LinkedInLoginURL = "https://www.linkedin.com/login"

# Catch output in a variable
op user list *> "$RunPath\temp.txt"
$CheckLogin = Get-Content -Path "$RunPath\temp.txt"

# 1Password item name
$ItemName = "Personal-LinkedIn"

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