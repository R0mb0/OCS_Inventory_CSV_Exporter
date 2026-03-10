#requires -Version 5.1
[CmdletBinding()]
param(
    # Example: https://your-server/ocsreports
    [Parameter(Mandatory = $true)]
    [string]$BaseUrl,

    # Optional override (defaults to "$BaseUrl/index.php")
    [string]$LoginUrl,

    # Optional override for the target search page
    # If omitted, it uses a default query similar to your original one.
    [string]$SearchUrl,

    [Parameter(Mandatory = $true)]
    [string]$Username,

    # Use either -Password or -PromptForPassword
    [string]$Password,

    [switch]$PromptForPassword,

    # Output directory for CSV file (default: script folder)
    [string]$OutputDirectory,

    # File name format
    [string]$TimestampFormat = "yyyy_MM_dd-HH_mm",

    # Headless browser mode
    [switch]$ShowBrowser
)

function Write-Log([string]$Message) {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message"
}

try {
    # Normalize base URL
    $BaseUrl = $BaseUrl.TrimEnd('/')

    if (-not $LoginUrl) {
        $LoginUrl = "$BaseUrl/index.php"
    }

    if (-not $SearchUrl) {
        $SearchUrl = "$BaseUrl/index.php?function=visu_search&fields=HARDWARE-LASTCOME&comp=tall&values=&values2=all&type_field="
    }

    if ($PromptForPassword -and [string]::IsNullOrWhiteSpace($Password)) {
        $secure = Read-Host "Enter password" -AsSecureString
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { $Password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
    }

    if ([string]::IsNullOrWhiteSpace($Password)) {
        throw "Password is required. Pass -Password or use -PromptForPassword."
    }

    if (-not $OutputDirectory) {
        $OutputDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
        if (-not $OutputDirectory) { $OutputDirectory = (Get-Location).Path }
    }
    New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

    $timestamp = Get-Date -Format $TimestampFormat
    $outCsv = Join-Path $OutputDirectory "$timestamp.csv"

    $tmpDir = Join-Path $env:TEMP "ocs_pw_export"
    $nodeScript = Join-Path $tmpDir "export.js"
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

    if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
        throw "Node.js not found. Please install Node.js LTS."
    }

    if (-not (Test-Path (Join-Path $tmpDir "package.json"))) {
        Write-Log "First run: installing Playwright dependencies..."
        Push-Location $tmpDir
        npm init -y | Out-Null
        npm i playwright | Out-Null
        npx playwright install chromium | Out-Null
        Pop-Location
    }

    $headless = if ($ShowBrowser) { "false" } else { "true" }

    # Single-quoted here-string avoids PowerShell variable expansion in JS
    $js = @'
const { chromium } = require("playwright");

function log(msg){ console.log(`[${new Date().toLocaleTimeString()}] ${msg}`); }

(async () => {
  const loginUrl  = process.env.OCS_LOGIN;
  const searchUrl = process.env.OCS_SEARCH;
  const user      = process.env.OCS_USER;
  const pass      = process.env.OCS_PASS;
  const outFile   = process.env.OCS_OUTFILE;
  const headless  = process.env.OCS_HEADLESS === "true";

  const browser = await chromium.launch({ headless });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  try {
    log("Opening login page...");
    await page.goto(loginUrl, { waitUntil: "domcontentloaded" });

    log("Filling credentials...");
    await page.fill('input[name="LOGIN"]', user);
    await page.fill('input[name="PASSWD"]', pass);

    log("Submitting login form...");
    await Promise.all([
      page.waitForLoadState("networkidle"),
      page.click('input[name="Valid_CNX"], #btn-logon')
    ]);

    log("Opening search page...");
    await page.goto(searchUrl, { waitUntil: "networkidle" });

    const select = "#select_colaffich_multi_crit";
    await page.waitForSelector(select, { timeout: 30000 });

    const values = await page.$$eval(select + " option", opts =>
      opts.map(o => o.value).filter(v => v && v !== "default")
    );

    log(`Columns found in selector: ${values.length}`);

    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      log(`Toggling column ${i + 1}/${values.length}: ${v}`);
      await page.selectOption(select, v);
      await page.waitForTimeout(400);
    }

    log("Waiting for table stabilization...");
    await page.waitForTimeout(1500);

    const exportSel = 'a[href*="function=export_csv"][href*="tablename=affich_multi_crit"][href*="nolimit=true"]';
    await page.waitForSelector(exportSel, { timeout: 30000 });

    log("Starting CSV download...");
    const [download] = await Promise.all([
      page.waitForEvent("download", { timeout: 60000 }),
      page.click(exportSel)
    ]);

    await download.saveAs(outFile);
    log("Download completed.");
  } catch (err) {
    console.error("ERR:" + err.message);
    process.exitCode = 1;
  } finally {
    log("Cleaning browser context...");
    await context.close();
    await browser.close();
  }
})();
'@

    Set-Content -Path $nodeScript -Value $js -Encoding UTF8

    $env:OCS_LOGIN    = $LoginUrl
    $env:OCS_SEARCH   = $SearchUrl
    $env:OCS_USER     = $Username
    $env:OCS_PASS     = $Password
    $env:OCS_OUTFILE  = $outCsv
    $env:OCS_HEADLESS = $headless

    Write-Log "Starting automation..."
    Push-Location $tmpDir
    try {
        node $nodeScript
        if ($LASTEXITCODE -ne 0) { throw "Playwright execution failed." }

        if (-not (Test-Path $outCsv)) { throw "CSV file not found." }

        $head = (Get-Content $outCsv -TotalCount 5 -ErrorAction SilentlyContinue) -join "`n"
        if ($head -match '<html|<!DOCTYPE') {
            throw "Downloaded file is HTML, not CSV."
        }

        Write-Log "SUCCESS: CSV saved at $outCsv"
    }
    finally {
        Pop-Location
        Remove-Item Env:OCS_LOGIN,Env:OCS_SEARCH,Env:OCS_USER,Env:OCS_PASS,Env:OCS_OUTFILE,Env:OCS_HEADLESS -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}