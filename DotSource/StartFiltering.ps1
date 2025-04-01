# Date for naming
$CurrentDate = Get-Date -Format "dd.MM.yy"
$FinalOutputFileName = "$($CurrentDate)-LikeParser-Report.xlsx"

# Rename existing file to new name
Rename-Item -Path $xlsxFile -NewName $FinalOutputFileName

$RunPath = "$RunPath\WorkingFolder"
$TxtFiles = Get-ChildItem -Path $RunPath -Filter "*.txt" -File

$UserStats = @{}
$ExtractedData = @()
$OriginalContent = @{}

foreach ($TxtFile in $TxtFiles) {
    $OriginalContent[$TxtFile.FullName] = Get-Content -Path $TxtFile.FullName
    $FileContent = $OriginalContent[$TxtFile.FullName]
    $Chunk = @()
    $ModifiedContent = @()

    foreach ($Line in $FileContent) {
        if ($Line -match "^\s*$") {
            if ($Chunk.Count -gt 0) { 
                $ModifiedContent += "#####[TOP]#####"
                $ModifiedContent += $Chunk
                $ModifiedContent += "#####[LOW]#####"
                $Chunk = @()
            }
        } else {
            $Chunk += $Line
        }
    }

    if ($Chunk.Count -gt 0) {
        $ModifiedContent += "#####[TOP]#####"
        $ModifiedContent += $Chunk
        $ModifiedContent += "#####[LOW]#####"
    }

    $ModifiedContent | Set-Content -Path $TxtFile.FullName

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

            if ($Chunk.Count -eq 5) {
                $LikeType = $Chunk[0]
                $UserName = $Chunk[1]
                $UserRelations = $Chunk[3]
                $UserTitle = $Chunk[4]

                if ($UserStats.ContainsKey($UserName)) {
                    $UserStats[$UserName]["LikeCount"] += 1
                    $UserStats[$UserName]["LikeTypes"] += ", $LikeType"
                } else {
                    $UserStats[$UserName] = @{
                        "LikeCount" = 1
                        "LikeTypes" = $LikeType
                    }
                }

                $ExtractedData += [PSCustomObject]@{
                    "Username" = $UserName
                    "Like Type" = $LikeType
                    "User Relations" = $UserRelations
                    "User Title" = $UserTitle
                }
            }

            $Chunk = @()
            continue
        }

        if ($Processing) {
            $Chunk += $Line
        }
    }
}

$UserStatsData = foreach ($User in $UserStats.Keys) {
    [PSCustomObject]@{
        "Username" = $User
        "Like Count" = $UserStats[$User]["LikeCount"]
        "Like Types" = $UserStats[$User]["LikeTypes"]
    }
}

$OutputFile = "$RunPath\$FinalOutputFileName"

# Usernames to exclude
$ExcludeUsernames = @(
    "Konstantin Markov",
    "Ido Naor 🇮🇱",
    "Jan Moronia",
    "Leo Valentić",
    "Jeff M",
    "Felipe Duarte",
    "Nrmn M.",
    "Donnie C.",
    "Nir Avron",
    "Jon Velardiez",
    "Anna Breeva",
    "Vedang Bhagwat",
    "Joao Andrade",
    "Reut Parshan",
    "Charles Lomboni",
    "Eilay Yosfan"
)

# Export sheets with new names
if ($ExtractedData.Count -gt 0 -or $UserStatsData.Count -gt 0) {
    $ExtractedData | Export-Excel -Path $OutputFile -WorksheetName "User Name Data Base" -AutoSize

    # Exclude usernames before exporting to "Hot Lead Detector"
    $FilteredStatsData = $UserStatsData | Where-Object { $_.Username.Trim() -notin $ExcludeUsernames }
    $FilteredStatsData | Export-Excel -Path $OutputFile -WorksheetName "Hot Lead Detector" -AutoSize

    # Explicitly import, sort, and export to prevent worksheet access errors
    $SortedStats = Import-Excel -Path $OutputFile -WorksheetName "Hot Lead Detector" |
        Sort-Object -Property "Like Count" -Descending

    $SortedStats | Export-Excel -Path $OutputFile -WorksheetName "Hot Lead Detector" -AutoSize -ClearSheet
}

# Restore original txt file content
foreach ($TxtFile in $TxtFiles) {
    if ($OriginalContent.ContainsKey($TxtFile.FullName)) {
        $OriginalContent[$TxtFile.FullName] | Set-Content -Path $TxtFile.FullName
    }
}

Write-Output ""
Write-Output "Your Like Status Report is Here -> $OutputFile"
