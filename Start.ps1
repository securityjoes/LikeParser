# Ensure ImportExcel is installed (Run once if needed)
# Install-Module -Name ImportExcel -Scope CurrentUser
Write-Output "   __ _ _          ___                         "
Write-Output "  / /(_) | _____  / _ \__ _ _ __ ___  ___ _ __ "
Write-Output " / / | | |/ / _ \/ /_)/ _' | '__/ __|/ _ \ '__|"
Write-Output "/ /__| |   <  __/ ___/ (_| | |  \__ \  __/ |   "   
Write-Output "\____/_|_|\_\___\/    \__,_|_|  |___/\___|_|   "
Write-Output ""
Write-Output "      GitHub.com/securityjoes/LikeParser"
Write-Output "             Author: Eilay Yosfan"
Write-Output "                 Version: 1.0"
Write-Output ""   

# Create a date variable for the excel output.
$CurrentDate = Get-Date -Format "dd.MM.yy"
$FinalOutputFileName = "$($CurrentDate)-Like-Status.xlsx"

if (Test-Path -Path "$RunPath\$FinalOutputFileName") {
    Remove-Item -Path "$RunPath\$FinalOutputFileName" -Force -Recurse -Confirm:$false | Out-Null
}
                                     
# Get the current working directory
$RunPath = Get-Location | Select-Object -ExpandProperty Path
$TxtFiles = Get-ChildItem -Path $RunPath -Filter "*.txt" -File
$TxtFileCount = $TxtFiles.Count

if ($TxtFileCount -eq 0) {
    Write-Output "[!] LikeParser did not find any .txt file."
    Write-Output "[!] Script has been stopped."
    exit
}

# Create dictionaries to track extracted data and user statistics
$UserStats = @{}
$ExtractedData = @()
$OriginalContent = @{}

foreach ($TxtFile in $TxtFiles) {
    # Save the original content of the file before modification
    $OriginalContent[$TxtFile.FullName] = Get-Content -Path $TxtFile.FullName

    # Read the file content
    $FileContent = $OriginalContent[$TxtFile.FullName]
    $Chunk = @()
    $ModifiedContent = @()

    foreach ($Line in $FileContent) {
        if ($Line -match "^\s*$") {  # Detect empty lines
            if ($Chunk.Count -gt 0) { 
                # Add markers around non-empty blocks
                $ModifiedContent += "#####[TOP]#####"
                $ModifiedContent += $Chunk
                $ModifiedContent += "#####[LOW]#####"
                $Chunk = @()
            }
        } else {
            # Add the line to the current chunk
            $Chunk += $Line
        }
    }

    # Process the last chunk if the file does not end with an empty line
    if ($Chunk.Count -gt 0) {
        $ModifiedContent += "#####[TOP]#####"
        $ModifiedContent += $Chunk
        $ModifiedContent += "#####[LOW]#####"
    }

    # Overwrite the file with the modified content
    $ModifiedContent | Set-Content -Path $TxtFile.FullName

    # Initialize variables for extracting data
    $Processing = $false
    $Chunk = @()

    foreach ($Line in $ModifiedContent) {
        if ($Line -match "####\[TOP\]#####") {
            $Processing = $true
            $Chunk = @()
            continue
        }

        if ($Line -match "####\[LOW\]#####") {
            $Processing = $false

            # Process blocks with exactly 5 lines
            if ($Chunk.Count -eq 5) {
                $LikeType = $Chunk[0]  # Extract like type
                $UserName = $Chunk[1]  # Extract username
                $UserRelations = $Chunk[3]  # Extract user relations
                $UserTitle = $Chunk[4]  # Extract user title

                # Update user statistics
                if ($UserStats.ContainsKey($UserName)) {
                    $UserStats[$UserName]["LikeCount"] += 1
                    $UserStats[$UserName]["LikeTypes"] += ", $LikeType"
                } else {
                    $UserStats[$UserName] = @{
                        "LikeCount" = 1
                        "LikeTypes" = $LikeType
                    }
                }

                # Store extracted data
                $ExtractedData += [PSCustomObject]@{
                    "Username" = $UserName
                    "Like Type" = $LikeType
                    "User Relations" = $UserRelations
                    "User Title" = $UserTitle
                }
            }

            # Reset the chunk for the next block
            $Chunk = @()
            continue
        }

        if ($Processing) {
            # Add the line to the current chunk
            $Chunk += $Line
        }
    }
}

# Convert user statistics to an array for exporting
$UserStatsData = foreach ($User in $UserStats.Keys) {
    [PSCustomObject]@{
        "Username" = $User
        "Like Count" = $UserStats[$User]["LikeCount"]
        "Like Types" = $UserStats[$User]["LikeTypes"]
    }
}

# Export extracted data to an Excel file if data exists
if ($ExtractedData.Count -gt 0 -or $UserStatsData.Count -gt 0) {
    $OutputFile = "$RunPath\$FinalOutputFileName"

    $ExtractedData | Export-Excel -Path $OutputFile -WorksheetName "Extracted Data" -AutoSize
    $UserStatsData | Export-Excel -Path $OutputFile -WorksheetName "User Stats" -AutoSize
}

# Restore the original content of each text file
foreach ($TxtFile in $TxtFiles) {
    if ($OriginalContent.ContainsKey($TxtFile.FullName)) {
        $OriginalContent[$TxtFile.FullName] | Set-Content -Path $TxtFile.FullName
    }
}

Write-Output "Your Like Status Report is Here -> $RunPath\$FinalOutputFileName"