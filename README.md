
# LikeParser v2.1
<img width="768" alt="LikeParser Banner" src="https://github.com/user-attachments/assets/db5cdde7-f496-4342-a4b1-3b6c9d410d25" />

## What is LikeParser?
LikeParser is a smart automation tool that turns your LinkedIn engagement into a goldmine for Marketing and Sales. It scrapes data from your LinkedIn profile or page and generates a detailed report with post-level insights — showing how each post performed and highlighting your top-performing content in terms of audience engagement.

You can choose the timeframe for scraping — from the past month to a full year — giving you the flexibility to analyze short-term trends or long-term performance.

But it doesn’t stop there — LikeParser also builds a ranked database of every user who interacted with your content, giving your Sales team a warm lead list ordered from most to least engaged. It’s a simple, powerful way to turn likes into leads.

Oh, and the best part? It’s completely free. LikeParser was designed to work without paying for LinkedIn’s API — no subscriptions, no tokens, just results.

## How to Use LikeParser?
Getting started with LikeParser is simple: just configure the `config.txt` file, then run the tool.

### Step 1 - Configuration
1. Open `config.txt` and select your preferred login method: <be>

   * 1Password: Set `LoginMethod=1Password` and `ItemName=YOUR_LINKEDIN_1PASSWORD_ITEM_NAME` to log in using saved credentials from 1Password. <be>
   
   * UserData (Chrome): If you're already logged into LinkedIn via Chrome, set `LoginMethod=UserData` to launch Chrome with your existing cookies and session. <br>
   
### Step 2 - Running
1. Download the tool, click on "<> Code" and then on "Download ZIP" and extract the folder to your desktop.
2. Lunch `PowerShell.exe`, navigate to the tool folder, and execute the start file like this `.\Start.ps1`

## Errors When Running LikeParser?
LikeParser has a built-in error handling system that displays clear, actionable messages directly in the PowerShell terminal when something goes wrong. These messages are designed to guide you through resolving common issues quickly. <be>

For example: <br>
`
[!] Chrome browser version is 114.0.5735.110 but your Chrome driver version is 113.0.5672.63
[!] Please download the matching Chrome driver for your Chrome browser
[!] And save it under C:\LikeParser\ChromeDriverWin64\
[!] URL to download: https://getwebdriver.com/
` <br>
How to fix it: Download the correct ChromeDriver version that matches your Chrome browser and replace the existing one in the specified folder.

Another example: <br>
`
[!] The ImportExcel PowerShell module was not detected and has now been successfully installed.
[!] If the script does not function as expected, please restart your computer and try again.
` <br>
How to fix it: In this example, the ImportExcel PowerShell module was not initially detected, but LikeParser automatically installed it.
In most cases, restarting your PowerShell session is enough to proceed. However, a full system restart may be more effective if issues persist.

This is the typical style of LikeParser’s error messages — informative and solution-oriented.
If the script encounters an error, always read the PowerShell terminal output first; it usually includes clear guidance on how to resolve the issue.
