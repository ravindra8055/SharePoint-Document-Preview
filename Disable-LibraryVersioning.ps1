#Requires -Version 5.1
<#
.SYNOPSIS
    Disables document versioning on SharePoint Online libraries.

.DESCRIPTION
    This script reads a CSV file containing SharePoint Online site URLs and library names,
    then disables version history on each specified library.

    Status values in the output:
    - "Updated"       — Versioning was disabled successfully
    - "AlreadyDisabled" — Versioning was already disabled, no change made
    - "Skipped"       — Row was skipped due to missing required fields
    - "Failed"        — An error occurred while processing the library

.PARAMETER CsvInputPath
    Path to CSV file with required columns: SiteUrl, LibraryName

.PARAMETER ClientId
    Azure AD app client ID for SPO PnP authentication.

.PARAMETER TargetUsername
    SPO username for authentication.

.PARAMETER TargetPassword
    SPO password for authentication.

.PARAMETER OutputFolder
    Folder path for output files.
    Default: ./DisableVersioningLog-{timestamp}

.EXAMPLE
    .\Disable-LibraryVersioning.ps1 -CsvInputPath ".\Libraries.csv" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TargetUsername "admin@tenant.onmicrosoft.com" `
        -TargetPassword "Password123!"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvInputPath,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TargetUsername,

    [Parameter(Mandatory = $true)]
    [string]$TargetPassword,

    [string]$OutputFolder = "./DisableVersioningLog-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ==========================================
# Global Variables & Initialization
# ==========================================
$script:ProcessedRows  = 0
$script:SuccessfulRows = 0
$script:FailedRows     = 0
$script:Credential     = $null
$script:CurrentSiteUrl = $null

# Streaming writer + canonical column order for the result CSV.
$script:ResultWriter   = $null
$script:ResultFilePath = $null
$script:ResultColumns  = @(
    "RowNumber",
    "SiteUrl",
    "LibraryName",
    "PreviousVersioningEnabled",
    "PreviousMajorVersionLimit",
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

    # Only reconnect if site URL changed
    if ($script:CurrentSiteUrl -eq $SiteUrl) {
        return
    }

    try {
        Write-Verbose "Connecting to $SiteUrl"
        Connect-PnPOnline -Url $SiteUrl -Credentials $script:Credential -ClientId $ClientId -ErrorAction Stop
        $script:CurrentSiteUrl = $SiteUrl
        Write-Host "Connected: $SiteUrl"
    }
    catch {
        throw "Failed to connect to site $SiteUrl : $_"
    }
}

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
    $requiredColumns = @("SiteUrl", "LibraryName")

    foreach ($column in $requiredColumns) {
        if (-not ($first.PSObject.Properties.Name -contains $column)) {
            throw "CSV must contain required column: $column"
        }
    }
}

# ==========================================
# Function: Initialize Result Writer
# ==========================================
function Initialize-ResultWriter {
    [CmdletBinding()]
    param()

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:ResultFilePath = Join-Path $OutputFolder "Results_$timestamp.csv"

    # UTF8 without BOM
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $script:ResultWriter = New-Object System.IO.StreamWriter($script:ResultFilePath, $false, $encoding)
    $script:ResultWriter.AutoFlush = $true

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

# ==========================================
# Function: Disable Library Versioning
# ==========================================
function Disable-LibraryVersioning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LibraryName
    )

    # Use explicit variable names instead of hashtable to avoid property access issues
    $resultSuccess = $false
    $resultStatus = "Failed"
    $resultMessage = ""
    $resultPrevVersioningEnabled = ""
    $resultPrevMajorVersionLimit = ""

    $list = $null
    $versioningEnabled = $null
    $majorVersionLimit = $null

    try {
        # Get the list first
        $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop

        if ($null -eq $list) {
            $resultMessage = "Library not found: $LibraryName"
            return @{
                Success                   = $resultSuccess
                Status                    = $resultStatus
                Message                   = $resultMessage
                PreviousVersioningEnabled = $resultPrevVersioningEnabled
                PreviousMajorVersionLimit = $resultPrevMajorVersionLimit
            }
        }
    }
    catch {
        $resultMessage = "Error getting library: $_"
        return @{
            Success                   = $resultSuccess
            Status                    = $resultStatus
            Message                   = $resultMessage
            PreviousVersioningEnabled = $resultPrevVersioningEnabled
            PreviousMajorVersionLimit = $resultPrevMajorVersionLimit
        }
    }

    # Explicitly load the versioning properties using Get-PnPProperty
    try {
        Get-PnPProperty -ClientObject $list -Property EnableVersioning, MajorVersionLimit -ErrorAction Stop | Out-Null
        $versioningEnabled = $list.EnableVersioning
        $majorVersionLimit = $list.MajorVersionLimit
        $resultPrevVersioningEnabled = [string]$versioningEnabled
        $resultPrevMajorVersionLimit = [string]$majorVersionLimit
    }
    catch {
        # Properties couldn't be loaded - try to proceed anyway
        Write-Warning "Could not load versioning properties for $LibraryName - will attempt to disable anyway"
        $resultPrevVersioningEnabled = "Unknown"
        $resultPrevMajorVersionLimit = "Unknown"
        $versioningEnabled = $true  # Assume enabled and try to disable
    }

    # Check if versioning is already disabled
    if ($versioningEnabled -eq $false) {
        $resultSuccess = $true
        $resultStatus = "AlreadyDisabled"
        $resultMessage = "Versioning was already disabled on this library"
        return @{
            Success                   = $resultSuccess
            Status                    = $resultStatus
            Message                   = $resultMessage
            PreviousVersioningEnabled = $resultPrevVersioningEnabled
            PreviousMajorVersionLimit = $resultPrevMajorVersionLimit
        }
    }

    # Disable versioning
    try {
        Set-PnPList -Identity $LibraryName -EnableVersioning $false -ErrorAction Stop

        $resultSuccess = $true
        $resultStatus = "Updated"
        if ($null -ne $majorVersionLimit) {
            $resultMessage = "Versioning disabled successfully. Previous setting: Enabled with $majorVersionLimit major versions"
        }
        else {
            $resultMessage = "Versioning disabled successfully"
        }
    }
    catch {
        $resultMessage = "Error disabling versioning: $_"
    }

    return @{
        Success                   = $resultSuccess
        Status                    = $resultStatus
        Message                   = $resultMessage
        PreviousVersioningEnabled = $resultPrevVersioningEnabled
        PreviousMajorVersionLimit = $resultPrevMajorVersionLimit
    }
}

# ==========================================
# Function: Invoke-DisableVersioning
# ==========================================
function Invoke-DisableVersioning {
    [CmdletBinding()]
    param()

    $startTime = Get-Date
    Write-Host "Reading CSV file: $CsvInputPath" -ForegroundColor Cyan

    $rows = Import-Csv -Path $CsvInputPath
    Test-CsvFormat -Rows $rows

    $totalRows = $rows.Count
    Write-Host "Found $totalRows rows to process" -ForegroundColor Cyan

    $rowIndex = 0
    foreach ($csvRow in $rows) {
        $rowIndex++
        $siteUrl = $csvRow.SiteUrl
        $libraryName = $csvRow.LibraryName

        # Progress indicator every 10 rows
        if ($rowIndex % 10 -eq 0) {
            Write-Host "Processing row $rowIndex of $totalRows..." -ForegroundColor Gray
        }

        # Skip rows with missing required fields
        if ([string]::IsNullOrWhiteSpace($siteUrl) -or [string]::IsNullOrWhiteSpace($libraryName)) {
            $row = [PSCustomObject]@{
                RowNumber                 = $rowIndex
                SiteUrl                   = $siteUrl
                LibraryName               = $libraryName
                PreviousVersioningEnabled = ""
                PreviousMajorVersionLimit = ""
                Status                    = "Skipped"
                IsSuccessful              = $false
                Message                   = "Missing required field (SiteUrl or LibraryName)"
                Timestamp                 = (Get-Date).ToString("s")
            }
            Write-ResultRow -Row $row
            $script:ProcessedRows++
            $script:FailedRows++
            Write-Warning "Row $rowIndex skipped: Missing required fields"
            continue
        }

        # Connect to site
        $connectionError = $null
        try {
            Get-PnPConnectionForSite -SiteUrl $siteUrl
        }
        catch {
            $connectionError = $_.Exception.Message
        }

        if ($connectionError) {
            $row = [PSCustomObject]@{
                RowNumber                 = $rowIndex
                SiteUrl                   = $siteUrl
                LibraryName               = $libraryName
                PreviousVersioningEnabled = ""
                PreviousMajorVersionLimit = ""
                Status                    = "Failed"
                IsSuccessful              = $false
                Message                   = "Connection failed: $connectionError"
                Timestamp                 = (Get-Date).ToString("s")
            }
            Write-ResultRow -Row $row
            $script:ProcessedRows++
            $script:FailedRows++
            Write-Warning "Row $rowIndex failed: Could not connect to $siteUrl"
            continue
        }

        # Disable versioning
        $result = Disable-LibraryVersioning -LibraryName $libraryName

        $row = [PSCustomObject]@{
            RowNumber                 = $rowIndex
            SiteUrl                   = $siteUrl
            LibraryName               = $libraryName
            PreviousVersioningEnabled = $result["PreviousVersioningEnabled"]
            PreviousMajorVersionLimit = $result["PreviousMajorVersionLimit"]
            Status                    = $result["Status"]
            IsSuccessful              = $result["Success"]
            Message                   = $result["Message"]
            Timestamp                 = (Get-Date).ToString("s")
        }

        Write-ResultRow -Row $row
        $script:ProcessedRows++

        if ($result["Success"]) {
            $script:SuccessfulRows++
            Write-Verbose "Row $rowIndex ($libraryName): $($result["Status"])"
        }
        else {
            $script:FailedRows++
            Write-Warning "Row $rowIndex ($libraryName): $($result["Message"])"
        }
    }

    # Write summary
    $endTime = Get-Date
    $duration = $endTime - $startTime

    $summaryFile = Join-Path $OutputFolder "Summary_$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    $summary = @{
        ExecutionTime  = "$($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
        TotalRows      = $totalRows
        ProcessedRows  = $script:ProcessedRows
        SuccessfulRows = $script:SuccessfulRows
        FailedRows     = $script:FailedRows
        ResultFile     = $script:ResultFilePath
        TimestampStart = $startTime.ToString("s")
        TimestampEnd   = $endTime.ToString("s")
    }

    $summary | ConvertTo-Json | Set-Content -Path $summaryFile -Encoding UTF8

    # Console summary
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Disable Library Versioning Summary" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Total Rows:       $($summary.TotalRows)"
    Write-Host "Processed Rows:   $($summary.ProcessedRows)"
    Write-Host "Successful Rows:  $($summary.SuccessfulRows)" -ForegroundColor Green
    if ($summary.FailedRows -gt 0) {
        Write-Host "Failed Rows:      $($summary.FailedRows)" -ForegroundColor Red
    }
    else {
        Write-Host "Failed Rows:      $($summary.FailedRows)" -ForegroundColor Green
    }
    Write-Host "Result CSV:       $($script:ResultFilePath)"
    Write-Host "Summary JSON:     $summaryFile"
    Write-Host "Execution Time:   $($summary.ExecutionTime)"
    Write-Host "==========================================" -ForegroundColor Cyan
}

# ==========================================
# Main Execution
# ==========================================
try {
    Write-Host "Starting Disable Library Versioning..." -ForegroundColor Cyan
    Initialize-Credential
    Initialize-ResultWriter
    Invoke-DisableVersioning
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
