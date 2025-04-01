# Variables to load
$ConfigPath = "$RunPath\Config\Config.txt"

# Collect the LoginMethod value from the config.txt
$LoginMethod = (Select-String -Path $ConfigPath -Pattern "LoginMethod").Line
$LoginMethod = $LoginMethod -replace '.*\=',''

# if statment to lunch the login method base on the value of LoginMethod
if ($LoginMethod -like "1Password") {
    # Dotsource 1PasswordLogin.ps1
    . "$RunPath\DotSource\1PasswordLogin.ps1"
}

else {
    # Dotsource UserDataLogin.ps1
    . "$RunPath\DotSource\UserDataLogin.ps1"
}