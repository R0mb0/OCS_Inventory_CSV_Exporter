<div align="center">
<h1>📦 OCS Inventory CSV Exporter</h1>

<p>
A production-ready <strong>PowerShell + Playwright</strong> automation tool that logs into OCS Inventory, opens the advanced search page, enables all selectable table columns, and downloads a full CSV export with timestamped filenames.
<br/><br/>
Built for reliability, privacy, and unattended execution (Windows Task Scheduler friendly).
</p>

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/OCS_Inventory_CSV_Exporter)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Playwright](https://img.shields.io/badge/Playwright-Automation-2EAD33?logo=playwright&logoColor=white)](https://playwright.dev/)
[![Node.js](https://img.shields.io/badge/Node.js-LTS-339933?logo=node.js&logoColor=white)](https://nodejs.org/)

</div>

<hr>

<h2>🚀 Why this project exists</h2>
<p>
This project was created to solve a real operational constraint: in some enterprise environments, the server hosting OCS is heavily restricted and direct database access is not allowed by policy.
</p>
<p>
In this scenario, CSV export from the web interface becomes the only approved and realistic way to extract the data needed to build downstream internal applications, reports, or integrations.
</p>
<p>
The repository is published to help anyone facing the same limitation: automating reliable CSV extraction when the application database cannot be contacted directly.
</p>

<h2>✨ Features</h2>
<ul>
<li><strong>Full-column export workflow</strong>: toggles all options in <code>#select_colaffich_multi_crit</code> before download.</li>
<li><strong>Timestamped output files</strong>: e.g. <code>2026_03_10-14_30.csv</code>.</li>
<li><strong>Headless automation</strong> (default) for scheduled/background use.</li>
<li><strong>Optional visible browser mode</strong> for troubleshooting and demos.</li>
<li><strong>No hardcoded environment details</strong>: base URL and credentials are parameterized.</li>
<li><strong>Detailed runtime logs</strong> printed to terminal with timestamps.</li>
<li><strong>First-run dependency bootstrap</strong>: auto-installs Playwright when needed.</li>
</ul>

<h2>📁 Repository structure</h2>
<ul>
<li><code>Script/ocs_export.ps1</code> → main automation script</li>
<li><code>Script/examples.ps1</code> → usage examples</li>
<li><code>README.md</code> → project documentation</li>
<li><code>LICENSE</code> → MIT license</li>
</ul>

<h2>🔧 Requirements</h2>
<ul>
<li><strong>Windows PowerShell 5.1+</strong> (or PowerShell 7+)</li>
<li><strong>Node.js LTS</strong> installed and available in PATH</li>
<li>Network access to your OCS instance</li>
<li>Valid OCS user credentials</li>
</ul>

<h2>⚙️ Parameters</h2>
<p>The script is fully parameterized to avoid exposing private infrastructure details in code.</p>

<ul>
<li><code>-BaseUrl</code> (required): OCS base URL, e.g. <code>https://your-server/ocsreports</code></li>
<li><code>-LoginUrl</code> (optional): defaults to <code>$BaseUrl/index.php</code></li>
<li><code>-SearchUrl</code> (optional): defaults to standard OCS “lastcome” search URL</li>
<li><code>-Username</code> (required): OCS username</li>
<li><code>-Password</code> (optional): plain text password (less secure)</li>
<li><code>-PromptForPassword</code> (optional switch): secure password prompt (recommended)</li>
<li><code>-OutputDirectory</code> (optional): CSV destination directory</li>
<li><code>-TimestampFormat</code> (optional): filename datetime format</li>
<li><code>-ShowBrowser</code> (optional switch): run non-headless for visual debugging</li>
</ul>

<h2>▶️ Quick start</h2>

<h3>1) Open PowerShell in the repository script folder</h3>
<pre><code>cd .\OCS_Inventory_CSV_Exporter\Script</code></pre>

<h3>2) Run using secure password prompt (recommended)</h3>
<pre><code>.\ocs_export.ps1 `
  -BaseUrl "https://your-server/ocsreports" `
  -Username "your_user" `
  -PromptForPassword `
  -OutputDirectory "C:\Exports\OCS"</code></pre>

<h3>3) Run with visible browser (debug mode)</h3>
<pre><code>.\ocs_export.ps1 `
  -BaseUrl "https://your-server/ocsreports" `
  -Username "your_user" `
  -PromptForPassword `
  -ShowBrowser</code></pre>

<h3>4) Run with custom search URL (optional)</h3>
<pre><code>.\ocs_export.ps1 `
  -BaseUrl "https://your-server/ocsreports" `
  -SearchUrl "https://your-server/ocsreports/index.php?function=visu_search&fields=HARDWARE-LASTCOME&comp=tall&values=&values2=all&type_field=" `
  -Username "your_user" `
  -PromptForPassword</code></pre>

<h2>🧠 How the script works (technical flow)</h2>
<ol>
<li>Validates prerequisites (<code>node</code>, Playwright runtime, output folder).</li>
<li>Builds a temporary Node.js Playwright runner script.</li>
<li>Opens OCS login page and authenticates using standard fields:
<ul>
<li><code>LOGIN</code></li>
<li><code>PASSWD</code></li>
<li><code>Valid_CNX</code> submit action</li>
</ul>
</li>
<li>Loads the target search page.</li>
<li>Finds <code>#select_colaffich_multi_crit</code> and iterates all options except <code>default</code>.</li>
<li>Triggers selector changes to toggle/show all relevant columns.</li>
<li>Clicks the export link with <code>function=export_csv</code>, <code>tablename=affich_multi_crit</code>, <code>nolimit=true</code>.</li>
<li>Saves the downloaded CSV with a timestamped filename.</li>
<li>Disposes browser context at the end to prevent stale cookie/session behavior.</li>
</ol>

<h2>🗓️ Automate daily execution (Task Scheduler)</h2>
<p>Recommended action:</p>
<pre><code>powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\Script\ocs_export.ps1" -BaseUrl "https://your-server/ocsreports" -Username "your_user" -PromptForPassword -OutputDirectory "C:\Exports\OCS"</code></pre>

<h2>🧪 Troubleshooting</h2>
<ul>
<li><strong>Node.js not found</strong> → install Node.js LTS and reopen terminal.</li>
<li><strong>Playwright first run is slow</strong> → expected; browser binaries are being installed.</li>
<li><strong>HTML downloaded instead of CSV</strong> → verify credentials, selectors, and current OCS UI layout.</li>
<li><strong>Columns not fully expanded</strong> → run with <code>-ShowBrowser</code> and inspect selector behavior in your OCS version.</li>
<li><strong>Execution policy errors</strong> → use <code>-ExecutionPolicy Bypass</code> in scheduled command.</li>
</ul>

<h2>🛠️ Development notes</h2>
<ul>
<li>The solution intentionally uses browser automation because some OCS setups depend on runtime UI state for export correctness.</li>
<li>Pure HTTP approaches may fail in environments with CSRF/session-dependent DataTables behavior.</li>
<li>This tool is especially useful when policy constraints prevent direct database-level extraction.</li>
</ul>
