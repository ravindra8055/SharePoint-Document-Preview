---
name: sharepoint-powershell
description: "Use when creating SharePoint PowerShell scripts: migration validation, user/group management, automation, permission auditing, or any SP2016/SP2019 on-premises to SharePoint Online task. Enforces authentication, streaming CSV output, and PS 5.1 compatibility patterns from the project reference scripts."
argument-hint: "Describe the script to create (e.g. 'audit site permissions', 'copy list items', 'validate migration')"
---

# SharePoint PowerShell Script Generation

This skill is kept in sync with the Claude Code companion at [../../../.claude/skills/sharepoint-powershell/SKILL.md](../../../.claude/skills/sharepoint-powershell/SKILL.md). If the two ever disagree, [CLAUDE.md](../../../CLAUDE.md) is the tiebreaker.

## When to Use
- Creating any new `.ps1` automation script for SharePoint Online or SP2016/SP2019
- Validation scripts (migration, permissions, content)
- User/group management scripts
- List/library automation scripts
- Any script that touches SP on-prem (source) and SPO (target)

---

## Critical Rule: Stream Results To Disk

**Write results to the output CSV incrementally — one row at a time, as soon as that row finishes processing. Do NOT accumulate results in a `$script:Results` array and write at the end.**

Why: input sets are large (100s of sites, 10s of thousands of folders/items), runs can take hours, and prior versions of these scripts lost everything when the process was killed, the network blipped, or `Export-Csv` ran out of memory at the end. Streaming writes mean partial results survive a crash, memory stays flat, and the operator can `Import-Csv` the output while the script is still running to spot-check progress.

How to apply: every script opens one `System.IO.StreamWriter` over the result CSV during initialization, writes the header once, and appends one CSV-encoded line per processed input row inside the per-row loop. The writer lives in `$script:` scope and is disposed in the `finally` block. Counters (`$script:ProcessedRows`, etc.) still live in memory — only row data is streamed. See Sections 8 and 9.

---

## Mandatory Conventions — Follow Every Time

### 1. File Header

```powershell
#Requires -Version 5.1
<#
.SYNOPSIS
    One-line summary.

.DESCRIPTION
    Multi-line description explaining what the script does,
    inputs it expects, and outputs it produces.

.PARAMETER CsvInputPath
    Path to CSV file with required columns: ...

.PARAMETER OutputFolder
    Folder path for output files.
    Default: ./ScriptNameLog-{timestamp}

.EXAMPLE
    .\Script-Name.ps1 -CsvInputPath ".\input.csv" `
        -SourceSiteUrl "http://sp2019.contoso.local/sites/Finance" `
        -TargetSiteUrl "https://tenant.sharepoint.com/sites/Finance" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TargetUsername "admin@tenant.onmicrosoft.com" `
        -TargetPassword "Password123!"
#>
```

- **Always include a full `.EXAMPLE` block with ALL parameters filled in with realistic placeholder values.**
- The example must be runnable as-is (no `<required>` placeholders).

---

### 2. Parameters Block

```powershell
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvInputPath,

    [Parameter(Mandatory = $true)]
    [string]$SourceSiteUrl,        # SP2016/SP2019 base URL — only when on-prem source is needed

    [Parameter(Mandatory = $true)]
    [string]$TargetSiteUrl,        # SharePoint Online site URL

    [Parameter(Mandatory = $true)]
    [string]$ClientId,             # Azure AD app client ID for SPO PnP authentication

    [Parameter(Mandatory = $true)]
    [string]$TargetUsername,       # SPO username (hardcoded credential)

    [Parameter(Mandatory = $true)]
    [string]$TargetPassword,       # SPO password (hardcoded credential)

    [string]$OutputFolder = "./ScriptNameLog-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
)
```

- Omit `$SourceSiteUrl` and CSOM sections when the script is SPO-only.
- Add extra `Mandatory` parameters specific to the script purpose after the standard set.
- `$OutputFolder` is always optional with a timestamped default.

---

### 3. Strict Mode and Error Preference

```powershell
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
```

Always place these immediately after the `param()` block.

---

### 4. Global Script Variables

```powershell
# ==========================================
# Global Variables & Initialization
# ==========================================
$script:ProcessedRows  = 0
$script:SuccessfulRows = 0
$script:FailedRows     = 0
$script:Credential     = $null
$script:SourceContexts = @{}   # only when CSOM / SP2016/SP2019 is needed

# Streaming writer + canonical column order for the result CSV.
# Set in Initialize-ResultWriter (Section 9), disposed in the finally block.
$script:ResultWriter   = $null
$script:ResultFilePath = $null
$script:ResultColumns  = @(
    "RowNumber",
    # ... your input-specific columns ...
    "Status",
    "IsSuccessful",
    "Message",
    "Timestamp"
)

if (-not (Test-Path $CsvInputPath)) {
    throw "CSV input file not found: $CsvInputPath"
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "Created output folder: $OutputFolder"
}
```

- Use `$script:` scope for all shared mutable state.
- **No `$script:Results` array.** Rows go straight to disk (see Section 8).
- **Initialize every variable that is used in a `catch` block BEFORE the `try` block** — StrictMode will throw on uninitialized variables even in catch blocks.

---

### 5. CSOM Assembly Loading — SP2016/SP2019 Source Only

Include this function whenever the script reads from SP2016 or SP2019 on-prem:

```powershell
# ==========================================
# Function: Import CSOM Assemblies
# ==========================================
function Import-CSOMAssemblies {
    [CmdletBinding()]
    param()

    $candidatePaths = @(
        "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI",
        "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI"
    )

    $loaded = $false
    foreach ($path in $candidatePaths) {
        $runtimeDll = Join-Path $path "Microsoft.SharePoint.Client.Runtime.dll"
        $clientDll  = Join-Path $path "Microsoft.SharePoint.Client.dll"

        if ((Test-Path $runtimeDll) -and (Test-Path $clientDll)) {
            Add-Type -Path $runtimeDll -ErrorAction SilentlyContinue
            Add-Type -Path $clientDll  -ErrorAction SilentlyContinue
            $loaded = $true
            break
        }
    }

    if (-not $loaded) {
        throw "CSOM assemblies not found. Ensure SharePoint Server CSOM or client libraries are installed."
    }
}
```

**SP2016/SP2019 context — uses current logged-in Windows user (DefaultNetworkCredentials):**

```powershell
# ==========================================
# Function: Get or Create Source Context
# ==========================================
function Get-SourceContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    $siteKey = $SiteUrl.ToLowerInvariant()

    if (-not $script:SourceContexts.ContainsKey($siteKey)) {
        try {
            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
            $ctx.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            $script:SourceContexts[$siteKey] = $ctx
        }
        catch {
            throw "Failed to create source context for $SiteUrl : $_"
        }
    }

    return $script:SourceContexts[$siteKey]
}
```

- Context is cached in `$script:SourceContexts` hashtable keyed by lowercase URL.
- **Never ask the user for SP2016/SP2019 credentials.** Always use `DefaultNetworkCredentials`.
- When loading CSOM items, always use plain `.Load()` + `.ExecuteQuery()` — **never use ScriptBlock-based lambda expressions** (e.g. `ctx.Load(list, l => l.Title)`) as they fail silently or crash in PS 5.1.

---

### 6. SPO Authentication — PnP PowerShell

**NEVER install or import PnP.PowerShell. It is already present in the execution environment.**

```powershell
# ==========================================
# Function: Initialize Credential
# ==========================================
function Initialize-Credential {
    [CmdletBinding()]
    param()

    $securePassword = ConvertTo-SecureString $TargetPassword -AsPlainText -Force
    $script:Credential = New-Object System.Management.Automation.PSCredential($TargetUsername, $securePassword)
}

# ==========================================
# Function: Connect to SPO Site
# ==========================================
function Get-PnPConnectionForSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    try {
        Write-Verbose "Connecting to $SiteUrl"
        Connect-PnPOnline -Url $SiteUrl -Credentials $script:Credential -ClientId $ClientId -ErrorAction Stop
        Write-Host "Connected: $SiteUrl"
    }
    catch {
        throw "Failed to connect to site $SiteUrl : $_"
    }
}
```

- Call `Get-PnPConnectionForSite` before every PnP operation that targets a new site URL.
- If the script only ever targets a single SPO site, connect once during initialization.

---

### 7. CSV Validation

```powershell
# ==========================================
# Function: Validate CSV Columns
# ==========================================
function Test-CsvFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Rows
    )

    if ($Rows.Count -eq 0) {
        throw "CSV file is empty: $CsvInputPath"
    }

    $first = $Rows[0]
    $requiredColumns = @("Column1", "Column2")   # replace with actual required columns

    foreach ($column in $requiredColumns) {
        if (-not ($first.PSObject.Properties.Name -contains $column)) {
            throw "CSV must contain required column: $column"
        }
    }
}
```

- Always skip rows where required fields are blank (`[string]::IsNullOrWhiteSpace`).
- Log skipped rows as `"Failed"` (or `"Skipped"`) with a descriptive `Message` — and still **write the row to the result CSV immediately**.

---

### 8. Result Object Schema + Streaming Write

Every input row produces exactly one result row, written to the result CSV **as soon as that row finishes processing** (success or failure). Build a `[PSCustomObject]`, hand it to `Write-ResultRow`, do not keep it in memory afterwards.

```powershell
# Inside the per-row loop:
$row = [PSCustomObject]@{
    RowNumber    = $rowIndex
    # ... input field columns ...
    Status       = "Added"      # "Added" | "AlreadyExists" | "Skipped" | "Failed"
                                # (or comparison vocabulary: "Match" | "Mismatch" | "MissingInTarget")
    IsSuccessful = $true
    Message      = "Human-readable detail"
    Timestamp    = (Get-Date).ToString("s")
}

Write-ResultRow -Row $row

$script:ProcessedRows++
if ($row.IsSuccessful) { $script:SuccessfulRows++ } else { $script:FailedRows++ }
```

Rules:
- `Status` is a string enum, not a boolean. Per-script vocabulary is fine (comparison scripts use `Match`/`Mismatch`/`MissingInTarget`/`MissingInSource`), but document it in the `.DESCRIPTION` block.
- `IsSuccessful = $true` iff `Status` is not `Failed`.
- The set of property names on `$row` must exactly match `$script:ResultColumns` (declared in Section 4). Mismatches break CSV alignment.
- **Never** accumulate rows into `$script:Results = @()` and write at the end. If you find yourself doing `$script:Results += ...`, stop and use `Write-ResultRow` instead.
- Long-running per-row work (e.g. listing 50,000 items inside one folder) still produces **one** result row at the end of that work. The streaming applies between input rows, not within a single row's processing.
- If the script also emits a secondary detail CSV (e.g. `MissingItems_Part<n>.csv`), it must use the same streaming pattern with its own writer (see Section 9).

---

### 9. Output Files — Streaming Writer

Open one `StreamWriter` per output CSV during initialization, write the header immediately, and dispose in `finally`.

```powershell
# ==========================================
# Function: Initialize Result Writer
# ==========================================
function Initialize-ResultWriter {
    [CmdletBinding()]
    param()

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:ResultFilePath = Join-Path $OutputFolder "Results_$timestamp.csv"

    # UTF8 without BOM (matches Export-Csv -Encoding UTF8 on PS 5.1+ for downstream tooling).
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $script:ResultWriter = New-Object System.IO.StreamWriter($script:ResultFilePath, $false, $encoding)
    $script:ResultWriter.AutoFlush = $true   # row-level durability; trades a bit of throughput for crash safety

    # Header
    $script:ResultWriter.WriteLine((ConvertTo-CsvLine -Values $script:ResultColumns))
}

# ==========================================
# Function: Write One Result Row
# ==========================================
function Write-ResultRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Row
    )

    $values = foreach ($col in $script:ResultColumns) {
        if ($Row.PSObject.Properties.Name -contains $col) { $Row.$col } else { "" }
    }
    $script:ResultWriter.WriteLine((ConvertTo-CsvLine -Values $values))
}

# ==========================================
# Function: CSV-Quote One Line
# ==========================================
function ConvertTo-CsvLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Values
    )

    $escaped = foreach ($v in $Values) {
        if ($null -eq $v) {
            ''
        }
        else {
            $s = [string]$v
            # RFC 4180: quote if value contains comma, quote, CR, or LF; double internal quotes.
            if ($s -match '[,"\r\n]') {
                '"' + ($s -replace '"', '""') + '"'
            }
            else {
                $s
            }
        }
    }
    [string]::Join(',', $escaped)
}
```

In the `finally` block:

```powershell
if ($script:ResultWriter) {
    try { $script:ResultWriter.Flush() }   catch { }
    try { $script:ResultWriter.Dispose() } catch { }
}
```

Summary JSON is still written **once** at the end (it's a single small object):

```powershell
$summaryFile = Join-Path $OutputFolder "Summary_$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

$summary = @{
    ExecutionTime  = "$($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
    TotalRows      = $rows.Count
    ProcessedRows  = $script:ProcessedRows
    SuccessfulRows = $script:SuccessfulRows
    FailedRows     = $script:FailedRows
    ResultFile     = $script:ResultFilePath
    TimestampStart = $startTime.ToString("s")
    TimestampEnd   = $endTime.ToString("s")
}

$summary | ConvertTo-Json | Set-Content -Path $summaryFile -Encoding UTF8
```

#### Rotating detail files (`MissingItems_Part<n>.csv`)

Comparison scripts that may emit huge detail rows still stream — they just rotate the writer when the current file hits `$MaxRecordsPerFile`:

```powershell
function Write-DetailRow {
    param([PSCustomObject]$Row)

    if ($script:DetailRowsInCurrentFile -ge $MaxRecordsPerFile) {
        $script:DetailWriter.Dispose()
        $script:DetailPartIndex++
        $script:DetailRowsInCurrentFile = 0
        Open-DetailWriter -PartIndex $script:DetailPartIndex   # writes header for new part
    }

    $values = foreach ($col in $script:DetailColumns) {
        if ($Row.PSObject.Properties.Name -contains $col) { $Row.$col } else { "" }
    }
    $script:DetailWriter.WriteLine((ConvertTo-CsvLine -Values $values))
    $script:DetailRowsInCurrentFile++
}
```

Reuse the same `ConvertTo-CsvLine` helper so quoting rules stay consistent across files.

---

### 10. Summary Console Output

```powershell
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Script Name Summary" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Total Rows:       $($summary.TotalRows)"
Write-Host "Processed Rows:   $($summary.ProcessedRows)"
Write-Host "Successful Rows:  $($summary.SuccessfulRows)" -ForegroundColor Green
Write-Host "Failed Rows:      $($summary.FailedRows)" -ForegroundColor $(if ($summary.FailedRows -gt 0) { "Red" } else { "Green" })
Write-Host "Result CSV:       $($script:ResultFilePath)"
Write-Host "Summary JSON:     $summaryFile"
Write-Host "Execution Time:   $($summary.ExecutionTime)"
Write-Host "==========================================" -ForegroundColor Cyan
```

Inside the per-row loop, emit a one-line progress note every N rows (e.g. every 50) so a long-running script visibly makes progress.

---

### 11. Main Execution Block

```powershell
# ==========================================
# Main Execution
# ==========================================
try {
    Write-Host "Starting <script purpose>..." -ForegroundColor Cyan
    Import-CSOMAssemblies      # only when SP2016/SP2019 source is used
    Initialize-Credential
    Initialize-ResultWriter    # MUST run before Invoke-MainFunction so per-row writes have a sink
    Invoke-MainFunction
}
catch {
    Write-Error "Fatal error: $_"
    exit 1
}
finally {
    if ($script:ResultWriter) {
        try { $script:ResultWriter.Flush() }   catch { }
        try { $script:ResultWriter.Dispose() } catch { }
    }
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Host "Script execution completed." -ForegroundColor Cyan
}
```

The `finally` block runs even on Ctrl+C or an unhandled exception — that's what makes streaming writes durable. Don't skip `Dispose()`.

---

### 12. Runspace Pools (Parallel Scripts)

When using runspace pools for parallelism (see `Compare-MigratedFolders.ps1`), the writer cannot be shared across threads safely. Two acceptable patterns:

1. **Per-thread temp file, merge at end.** Each runspace writes its own `Results_thread<n>.csv`; the main thread concatenates them in the `finally` block (header from thread 0, body only from the rest).
2. **Producer/consumer with a thread-safe queue.** Runspaces enqueue `[PSCustomObject]` rows into a `System.Collections.Concurrent.ConcurrentQueue`; a single dedicated writer thread drains the queue and calls `Write-ResultRow`. Simpler than locking the StreamWriter directly.

Do **not** wrap `$script:ResultWriter.WriteLine(...)` in a `[System.Threading.Monitor]::Enter`/`Exit` and call it from multiple runspaces — `$script:` scope does not cross runspace boundaries cleanly in PS 5.1 and you will lose rows silently.

---

## PowerShell 5.1 Compatibility — Hard Rules

| Rule | Detail |
|------|--------|
| No `??` null-coalescing | Use `if ($x -eq $null) { $x = $default }` |
| No `?:` ternary | Use `if (...) { ... } else { ... }` or inline `$(if (...) {...} else {...})` in strings |
| No `&&` / `\|\|` operators | Use `-and` / `-or` |
| No `ForEach-Object -Parallel` | Use runspace pools (see Section 12) or sequential loops |
| No CSOM lambda/ScriptBlock loads | Use `$ctx.Load($obj); $ctx.ExecuteQuery()` |
| No `#Requires -PSEdition Core` | Script must run on Desktop edition |
| No `[System.Collections.Generic.List[T]]` shorthand | Use `New-Object 'System.Collections.Generic.List[pscustomobject]'` or just `@()` arrays |
| Initialize before try/catch | Variables referenced in `catch` must be declared before the `try` block |
| No `class` keyword | Use functions and `[PSCustomObject]@{}` instead |
| No `using module` | Use `Import-Module` (or omit if module auto-loads) |
| No `$script:Results += ...` accumulator | Stream each row via `Write-ResultRow` (see Section 8) |

---

## Section Comment Headers

Separate every logical section with:

```powershell
# ==========================================
# Section Name
# ==========================================
```

Use these standard section names (in order):
1. Global Variables & Initialization
2. Function: Import CSOM Assemblies *(on-prem only)*
3. Function: Get or Create Source Context *(on-prem only)*
4. Function: Initialize Credential
5. Function: Connect to SPO Site
6. Function: Validate CSV Columns
7. Function: Initialize Result Writer
8. Function: Write One Result Row
9. Function: CSV-Quote One Line
10. Function: \<Core domain functions\>
11. Function: \<Main processing function — named `Invoke-*`\>
12. Main Execution

---

## Module Policy

| Module | Rule |
|--------|------|
| `PnP.PowerShell` | Pre-installed. **Do not** call `Install-Module` or `Import-Module`. Use PnP cmdlets directly. |
| `SharePointPnPPowerShellOnline` | **Do not use.** PS 5.1 uses `PnP.PowerShell` only. |
| CSOM DLLs | Load via `Add-Type` from SharePoint hive ISAPI folder (see `Import-CSOMAssemblies`). |

---

## Quick Self-Review Before Returning A Script

Before reporting a generated script as done, scan it for these specifically:

- [ ] `Initialize-ResultWriter` is called in the main `try` block **before** any row processing.
- [ ] Every place a result row is produced calls `Write-ResultRow` (or the secondary detail-file equivalent). Search for `$script:Results` and `+= [PSCustomObject]` — both should return zero hits.
- [ ] The `finally` block disposes every writer the script opened.
- [ ] `$script:ResultColumns` and the property names set on each `[PSCustomObject]` row match exactly.
- [ ] No PS 7-only syntax (re-check the table above).
- [ ] All variables referenced from `catch` blocks are initialized before their `try`.
- [ ] If reading SP2016/SP2019, no `Get-Credential` prompt — `DefaultNetworkCredentials` only.
- [ ] No `Install-Module` / `Import-Module PnP.PowerShell` anywhere.

---

## Reference Templates

See the project reference scripts for complete working implementations:
- [Add-SharePointUsersToGroups.ps1](../../../Add-SharePointUsersToGroups.ps1) — SPO-only, CSV-driven user/group management
- [Compare-MigratedFolders.ps1](../../../Compare-MigratedFolders.ps1) — SP2019 CSOM source + SPO PnP target, folder comparison, runspace pool

Note: the reference scripts predate the streaming-CSV rule in this skill. If you modify one of them, migrate it to streaming writes at the same time; if you create a new script, follow this skill (not the legacy pattern in the reference).
