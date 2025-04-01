# Variables to load
$LinkedInURL = "https://www.linkedin.com/feed/"
$UserDataDir = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data"

# Chrome options with profile
$chromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$chromeOptions.BinaryLocation = $ChromePath
$chromeOptions.AddArgument("--user-data-dir=$UserDataDir")
$chromeOptions.AddArgument("profile-directory=Default")
$chromeOptions.AddArgument("--start-maximized")

# Start driver using raw Selenium
$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($WebDriverPath)
$global:Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $chromeOptions)

# Go to LinkedIn
$global:Driver.Navigate().GoToUrl($LinkedInURL)
