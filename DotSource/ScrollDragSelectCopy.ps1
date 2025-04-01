Start-Sleep -Seconds 1
# Navigate to the LinkedIn post
Enter-SeUrl $URL -Driver $Driver
Start-Sleep -Seconds 3

# Open the reaction tab
Invoke-SeClick -Element (Find-SeElement -Driver $Driver -XPath '/html/body/div[5]/div[3]/div/div/div/div[2]/div/div/main/div/section/div/div/div/div/div/div[1]/div[4]/div[1]/div/div/ul/li[1]/button')
Start-Sleep -Seconds 1

# Scroll + drag + select
$js = @"
(function() {
    const SCROLL_DURATION_SEC = 10;
    const SCROLL_INTERVAL_SEC = 0.05;
    const SCROLL_STEP_PX = 50;

    const SCROLL_DURATION_MS = SCROLL_DURATION_SEC * 1000;
    const SCROLL_INTERVAL_MS = SCROLL_INTERVAL_SEC * 1000;

    const container = document.querySelector('div.artdeco-modal__content');
    if (!container) return;

    const selectable = container.querySelector('li, div, span, p');
    if (!selectable) return;

    const rect = selectable.getBoundingClientRect();
    const startX = rect.left + 10;
    const startY = rect.top + 10;

    const totalSteps = SCROLL_DURATION_MS / SCROLL_INTERVAL_MS;
    const endY = startY + (totalSteps * SCROLL_STEP_PX);

    const down = new MouseEvent('mousedown', {
        clientX: startX,
        clientY: startY,
        bubbles: true,
        cancelable: true,
        buttons: 1
    });
    selectable.dispatchEvent(down);

    let currentY = startY;
    let stepCount = 0;

    const interval = setInterval(() => {
        currentY += SCROLL_STEP_PX;
        stepCount++;

        const move = new MouseEvent('mousemove', {
            clientX: startX,
            clientY: currentY,
            bubbles: true,
            cancelable: true,
            buttons: 1
        });
        selectable.dispatchEvent(move);
        container.scrollBy(0, SCROLL_STEP_PX);

        if (stepCount >= totalSteps) {
            clearInterval(interval);

            const up = new MouseEvent('mouseup', {
                clientX: startX,
                clientY: currentY,
                bubbles: true,
                cancelable: true
            });
            selectable.dispatchEvent(up);

            // Select all text
            const selection = window.getSelection();
            selection.removeAllRanges();
            const range = document.createRange();
            range.selectNodeContents(container);
            selection.addRange(range);
        }
    }, SCROLL_INTERVAL_MS);
})();
"@

# Execute the Scroll + drag + select script
$Driver.ExecuteScript($js)

# Sleep 1 second more then the Scroll + drag + select script
Start-Sleep 11

# Initiate a copy highlited text and save inside $selectedText
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait("^c")  # Sends Ctrl+C

$copiedText = Get-Clipboard
$copiedText | Set-Content -Path "$RunPath\WorkingFolder\$Counter.txt" -Encoding UTF8
