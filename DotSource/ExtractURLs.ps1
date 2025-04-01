# Import the ImportExcel module
Import-Module ImportExcel

# Define file paths
$SheetName = "All posts"
$URLsTextFile = "$RunPath\WorkingFolder\URLs.txt"

# Import data from the xlsx Excel file without assuming headers
$Data = Import-Excel -Path $xlsxFile -WorksheetName $SheetName -NoHeader

# Extract column B (Index 1, since PowerShell uses zero-based indexing)
$ColumnB = $Data | ForEach-Object { $_.PSObject.Properties.Value[1] }

# Save to text file
$ColumnB | Out-File -FilePath $URLsTextFile -Encoding UTF8

(Get-Content $URLsTextFile) | Select-Object -Skip 1 | Set-Content $URLsTextFile