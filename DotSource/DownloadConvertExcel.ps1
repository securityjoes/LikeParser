# Currect DotSource variables
$WorkingFolder = "$RunPath\WorkingFolder"
$DownloadPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$AnalyticsURL = "https://www.linkedin.com/company/18706304/admin/analytics/updates/"

# Naviagte to LinkedIn Analytics tab
Enter-SeUrl $AnalyticsURL -Driver $Driver
Start-Sleep -Seconds 1

# Create the working direcotry fodler
if (Test-Path -Path $WorkingFolder) {
    Remove-Item -Path $WorkingFolder -Force -Recurse -Confirm:$false | Out-Null
    New-Item -ItemType Directory -Path $WorkingFolder -ErrorAction SilentlyContinue | Out-Null
}
else {
    New-Item -ItemType Directory -Path $WorkingFolder -ErrorAction SilentlyContinue | Out-Null
}

# Display a pop-up to instruct the user to download the desired file
$Driver.ExecuteScript("alert('Choose the timeframe and click on Export to download the file.');")

# Set the timeout (max wait time in seconds)
$Timeout = 300  
$WaitTime = 1  # Time between checks (in seconds)
$ElapsedTime = 0

# Get the initial list of existing .xls or .xlsx files
$ExistingFiles = Get-ChildItem -Path $DownloadPath -Filter "*.xls" | Select-Object -ExpandProperty Name

# Initialize variable to store the detected file
$FileName = $null
$FilePath = $null

# Loop to wait for a truly new .xls
do {
    Start-Sleep -Seconds $WaitTime

    # Get the list of .xls or .xlsx files again
    $NewFiles = Get-ChildItem -Path $DownloadPath -Filter "*.xls*" | Select-Object -ExpandProperty Name

    # Find files that did not exist before
    $DetectedNewFile = $NewFiles | Where-Object { $_ -notin $ExistingFiles }

    if ($DetectedNewFile) {
        $FileName = $DetectedNewFile
        $FilePath = "$DownloadPath\$FileName"
        break
    }

    $ElapsedTime += $WaitTime
} while ($ElapsedTime -lt $Timeout)

# If no new file was found within the timeout period, exit with an error
if (-not $FileName) { exit 1 }

# Move $FilePath from downloads folder to $WorkingFolder
Start-Sleep -Seconds 3
Move-Item -Path $FilePath -Destination $WorkingFolder | Out-Null

# New path of the xls file after moving it
$FilePath = "$WorkingFolder\$FileName"

# Start Excel application
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

# Open the XLS file
$Workbook = $Excel.Workbooks.Open($FilePath)

# Define the new XLSX file path
$xlsxFile = "$([System.IO.Path]::ChangeExtension($FilePath, 'xlsx'))"

# Save as XLSX format (FileFormat 51 is for .xlsx)
$Workbook.SaveAs($xlsxFile, 51)

# Close and clean up
$Workbook.Close($false)
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# remove
Remove-Item -Path $FilePath