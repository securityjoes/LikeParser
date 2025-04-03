# Dotsource Banner.ps1 (Print it after the 'clear' prompt)
. "$RunPath\DotSource\Banner.ps1"

# Get text content of URLs.txt and save to a variable
$URLS = Get-Content -Path $URLsTextFile

# Loop Counter for file names
$LineCounter = 0

# foreach loop all the urls
foreach ($URL in $URLS) {
    # +1 the value each loop
    $LineCounter++
}

Write-Output "There is $LineCounter posts to scrap data from."

# Loop Counter for file names
$Counter = 0

# foreach loop all the urls
foreach ($URL in $URLS) {
    # +1 the value each loop
    $Counter++

    Write-Output "Scraping data from post number [$Counter] -> $URL"

    # Dotsource ScrollDragSelectCopy.ps1
    . "$RunPath\DotSource\ScrollDragSelectCopy.ps1"
}

Remove-Item -Path "$RunPath\WorkingFolder\URLs.txt"