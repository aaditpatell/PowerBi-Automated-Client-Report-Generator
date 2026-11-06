<#
.SYNOPSIS
    Creates one ready-to-load Power BI template (.pbit) per client from a folder of quarterly billing workbooks.

.DESCRIPTION
    1. Reads every Users_Report_*.xlsx in the input folder.
    2. Works out the reporting group, fiscal year and quarter from each file name
       (same rules as the template, e.g. Users_Report_LTC_2026-2027_Q3.xlsx -> LTC, 2026-2027, Q3;
       no group in the name -> "Overall").
    3. For each reporting group and fiscal year, picks the report quarter (the latest one found,
       or the one given with -Quarter) and the workbooks it needs:
           Q1 report -> Q1         Q2 report -> Q1, Q2
           Q3 report -> Q1, Q2, Q3 Q4 report -> Q1, Q2, Q3 (Q3 is used as Q4, no Q4 workbook needed)
    4. Writes a copy of the template into the output folder with Report Quarter and the
       Q1/Q2/Q3 Excel Workbook values built in, e.g. 2026-2027-LTC-Q3.pbit. The parameters are
       turned into fixed values, so the file opens without asking for anything.

    Each client's file only points at that client's own workbooks, so the finished report
    contains only that client's data.

    After running: open each .pbit (Power BI loads the data by itself), then File > Save as .pbix.

.PARAMETER InputFolder
    Folder containing the client workbooks (Users_Report_*.xlsx).

.PARAMETER TemplatePath
    The Automated Client User Billing Report Template (.pbit).

.PARAMETER OutputFolder
    Folder for the generated .pbit files and the generation log. Created if it does not exist.

.PARAMETER Quarter
    Optional. Q1, Q2, Q3 or Q4. Generate this quarter for every client instead of each client's latest.

.PARAMETER FiscalYear
    Optional. e.g. 2026-2027. Only generate reports for this fiscal year.

.PARAMETER NameFormat
    Optional. How output files are named, using {FiscalYear}, {Client} and {Quarter}.
    Default: {FiscalYear}-{Client}-{Quarter}  ->  2026-2027-JTS-Q2.pbit
    Example: -NameFormat '{FiscalYear}-{Quarter}-{Client}'  ->  2026-2027-Q2-JTS.pbit

.EXAMPLE
    .\Generate-ClientReports.ps1 -InputFolder "C:\Billing\2026-2027 Q3" -TemplatePath "C:\Billing\Automated_Client_User_Billing_Report_Template.pbit" -OutputFolder "C:\Billing\Reports Q3"

.EXAMPLE
    .\Generate-ClientReports.ps1 -InputFolder "C:\Billing\All" -TemplatePath ".\Template.pbit" -OutputFolder ".\Out" -Quarter Q4 -FiscalYear 2026-2027
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $InputFolder,
    [Parameter(Mandatory = $true)] [string] $TemplatePath,
    [Parameter(Mandatory = $true)] [string] $OutputFolder,
    [ValidateSet('Q1', 'Q2', 'Q3', 'Q4')] [string] $Quarter,
    [string] $FiscalYear,
    # Output file name pattern. Change the order here to rename every generated file.
    [string] $NameFormat = '{FiscalYear}-{Client}-{Quarter}'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------------------------
# File name parsing (mirrors fnParseReportFileName in the template)
# ---------------------------------------------------------------------------------------------
function Test-FiscalYear([string] $Text) {
    if ($Text -notmatch '^\d{4}-\d{4}$') { return $false }
    return ([int]$Text.Substring(5, 4) -eq [int]$Text.Substring(0, 4) + 1)
}

function Get-ReportInfo([System.IO.FileInfo] $File) {
    $parts = @($File.BaseName -split '[_ ]' | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() })
    $fyPos = -1
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if (Test-FiscalYear $parts[$i]) { $fyPos = $i; break }
    }
    if ($fyPos -lt 0 -or $fyPos + 1 -ge $parts.Count) { return $null }

    $quarterText = $parts[$fyPos + 1]
    if ($quarterText.Length -lt 2) { return $null }
    $q = $quarterText.Substring(0, 2).ToUpper()
    if ($q -notmatch '^Q[1-4]$') { return $null }

    $groupParts = @()
    for ($i = 0; $i -lt $fyPos; $i++) {
        if ($parts[$i] -ne 'Users' -and $parts[$i] -ne 'Report') { $groupParts += $parts[$i] }
    }
    $group = 'Overall'
    if ($groupParts.Count -gt 0) { $group = ($groupParts -join ' ') }

    return [PSCustomObject]@{
        Group      = $group
        FiscalYear = $parts[$fyPos]
        Quarter    = $q
        QuarterNum = [int]$q.Substring(1, 1)
        File       = $File
    }
}

# ---------------------------------------------------------------------------------------------
# Template editing: set parameter values in the query definitions (UTF-16 JSON).
# A .pbit can hold the queries twice: in DataModelSchema (the model) and in UnappliedChanges
# (Power Query's copy, which Power BI uses when present). Both are updated the same way.
# ---------------------------------------------------------------------------------------------
function ConvertTo-JsonEscapedMText([string] $Value) {
    # M text literal ("" escapes a quote), then JSON-escape it for the schema file
    $m = '"' + $Value.Replace('"', '""') + '"'
    return $m.Replace('\', '\\').Replace('"', '\"')
}

# matches the current value of a parameter's expression: null, or an existing "text" value
$ValuePattern = '(?:null|\\"(?:[^\\]|\\\\)*?\\")'

function Set-NamedParameter([string] $Schema, [string] $Name, [string] $Value) {
    $pattern = '("name":\s*"' + [regex]::Escape($Name) + '",(?:(?!"name":)[\s\S])*?"(?:expression|text)":\s*\[?\s*")' + $ValuePattern + '( meta \[IsParameterQuery=true)'
    $found = [regex]::Matches($Schema, $pattern).Count
    if ($found -ne 1) { throw "Template check failed: expected 1 parameter named '$Name', found $found." }
    $replacement = if ($Value) { ConvertTo-JsonEscapedMText $Value } else { 'null' }
    return [regex]::Replace($Schema, $pattern, { param($m) $m.Groups[1].Value + $replacement + $m.Groups[2].Value })
}

function Set-ReportQuarter([string] $Schema, [string] $Value) {
    $pattern = '("(?:expression|text)":\s*\[?\s*")' + $ValuePattern + '( meta \[IsParameterQuery=true, List=\{\\"Q1\\")'
    $found = [regex]::Matches($Schema, $pattern).Count
    if ($found -ne 1) { throw "Template check failed: expected 1 Report Quarter parameter, found $found." }
    $replacement = ConvertTo-JsonEscapedMText $Value
    return [regex]::Replace($Schema, $pattern, { param($m) $m.Groups[1].Value + $replacement + $m.Groups[2].Value })
}

function Update-Queries([string] $Text, [hashtable] $Values) {
    $Text = Set-ReportQuarter $Text $Values['Report Quarter']
    foreach ($name in @('Q1 Excel Workbook', 'Q2 Excel Workbook', 'Q3 Excel Workbook')) {
        $Text = Set-NamedParameter $Text $name $Values[$name]
    }
    # Turn every parameter into a fixed value. A template that still has parameters always opens
    # the "enter parameters" window; without them Power BI opens the file and loads the data directly.
    $paramPattern = ' meta \[IsParameterQuery=true[^\]]*\](?=")'
    $paramCount = [regex]::Matches($Text, $paramPattern).Count
    if ($paramCount -lt 4) { throw "Template check failed: expected the template's parameters, found $paramCount." }
    $Text = [regex]::Replace($Text, $paramPattern, '')
    # also clear the cached "last loaded" copy of the formulas (stored with an extra level of escaping)
    return [regex]::Replace($Text, ' meta \[IsParameterQuery=true[^\]]*\](?=\\+")', '')
}

function Read-EntryBytes([System.IO.Compression.ZipArchiveEntry] $Entry) {
    $stream = $Entry.Open()
    try {
        $buffer = New-Object System.IO.MemoryStream
        $stream.CopyTo($buffer)
        return , $buffer.ToArray()
    }
    finally { $stream.Dispose() }
}

function New-ClientTemplate([string] $Template, [string] $Destination, [hashtable] $Values) {
    $source = [System.IO.Compression.ZipFile]::OpenRead($Template)
    try {
        $schemaEntry = $source.GetEntry('DataModelSchema')
        if (-not $schemaEntry) { throw "The template has no DataModelSchema. Make sure it is a .pbit file, not a .pbix." }
        $schema = [System.Text.Encoding]::Unicode.GetString((Read-EntryBytes $schemaEntry))

        $schemaBytes = [System.Text.Encoding]::Unicode.GetBytes((Update-Queries $schema $Values))   # UTF-16 LE, no BOM, like Power BI

        $unappliedBytes = $null
        $unappliedEntry = $source.GetEntry('UnappliedChanges')
        if ($unappliedEntry) {
            $unapplied = [System.Text.Encoding]::Unicode.GetString((Read-EntryBytes $unappliedEntry))
            $unappliedBytes = [System.Text.Encoding]::Unicode.GetBytes((Update-Queries $unapplied $Values))
        }

        if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
        $target = [System.IO.Compression.ZipFile]::Open($Destination, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($entry in $source.Entries) {
                # SecurityBindings signs the original contents; Power BI rejects the file if it no longer matches
                if ($entry.FullName -eq 'SecurityBindings') { continue }
                $bytes = Read-EntryBytes $entry
                if ($entry.FullName -eq 'DataModelSchema') { $bytes = $schemaBytes }
                elseif ($entry.FullName -eq 'UnappliedChanges') { $bytes = $unappliedBytes }
                elseif ($entry.FullName -eq '[Content_Types].xml') {
                    $xml = [System.Text.Encoding]::UTF8.GetString($bytes)
                    $xml = [regex]::Replace($xml, '<Override PartName="/SecurityBindings" ContentType="[^"]*"\s*/>', '')
                    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($xml)
                }
                $newEntry = $target.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                $newEntry.LastWriteTime = $entry.LastWriteTime
                $out = $newEntry.Open()
                try { $out.Write($bytes, 0, $bytes.Length) } finally { $out.Dispose() }
            }
        }
        finally { $target.Dispose() }
    }
    finally { $source.Dispose() }
}

# ---------------------------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------------------------
$InputFolder = (Resolve-Path -LiteralPath $InputFolder).Path
$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path

$files = @(Get-ChildItem -LiteralPath $InputFolder -Filter '*.xlsx' -File | Where-Object { -not $_.Name.StartsWith('~$') })
if ($files.Count -eq 0) { throw "No .xlsx files found in $InputFolder" }

$reports = @{}
$skippedFiles = @()
foreach ($f in $files) {
    $info = Get-ReportInfo $f
    if (-not $info) { $skippedFiles += $f.Name; continue }
    if ($FiscalYear -and $info.FiscalYear -ne $FiscalYear) { continue }
    $key = $info.Group + '|' + $info.FiscalYear
    if (-not $reports.ContainsKey($key)) { $reports[$key] = @() }
    $reports[$key] += $info
}

$log = @()
foreach ($key in ($reports.Keys | Sort-Object)) {
    $items = $reports[$key]
    $group = $items[0].Group
    $fy = $items[0].FiscalYear

    # report quarter: the one requested, or the latest one available for this client
    if ($Quarter) { $reportQ = [int]$Quarter.Substring(1, 1) }
    else { $reportQ = ($items | Measure-Object -Property QuarterNum -Maximum).Maximum }
    $needed = 1..([Math]::Min($reportQ, 3))

    $values = @{ 'Report Quarter' = "Q$reportQ"; 'Q1 Excel Workbook' = $null; 'Q2 Excel Workbook' = $null; 'Q3 Excel Workbook' = $null }
    $missing = @()
    $notes = @()
    foreach ($q in $needed) {
        $candidates = @($items | Where-Object { $_.QuarterNum -eq $q } | Sort-Object { $_.File.LastWriteTime } -Descending)
        if ($candidates.Count -eq 0) { $missing += "Q$q"; continue }
        if ($candidates.Count -gt 1) {
            $notes += "Q$q has $($candidates.Count) versions, used newest: $($candidates[0].File.Name)"
        }
        $values["Q$q Excel Workbook"] = $candidates[0].File.FullName
    }

    $baseName = $NameFormat.Replace('{FiscalYear}', $fy).Replace('{Client}', $group).Replace('{Quarter}', "Q$reportQ")
    $safeName = ($baseName -replace '[\\/:*?"<>|]', '-') + '.pbit'
    $destination = Join-Path $OutputFolder $safeName

    if ($missing.Count -gt 0) {
        $status = 'SKIPPED - missing ' + ($missing -join ', ')
    }
    else {
        try {
            New-ClientTemplate $TemplatePath $destination $values
            $status = 'Created'
        }
        catch { $status = 'FAILED - ' + $_.Exception.Message }
    }

    $log += [PSCustomObject]@{
        ReportingGroup = $group
        FiscalYear     = $fy
        ReportQuarter  = "Q$reportQ"
        Status         = $status
        OutputFile     = $(if ($status -eq 'Created') { $safeName } else { '' })
        Q1Workbook     = $(if ($values['Q1 Excel Workbook']) { Split-Path $values['Q1 Excel Workbook'] -Leaf } else { '' })
        Q2Workbook     = $(if ($values['Q2 Excel Workbook']) { Split-Path $values['Q2 Excel Workbook'] -Leaf } else { '' })
        Q3Workbook     = $(if ($values['Q3 Excel Workbook']) { Split-Path $values['Q3 Excel Workbook'] -Leaf } else { '' })
        Notes          = ($notes -join '; ')
    }
}

$logPath = Join-Path $OutputFolder 'generation_log.csv'
$log | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8

$log | Format-Table ReportingGroup, FiscalYear, ReportQuarter, Status, OutputFile -AutoSize | Out-String -Width 250 | Write-Host
if ($skippedFiles.Count -gt 0) {
    Write-Host "Ignored (name not in the Users_Report_<group>_<YYYY-YYYY>_<Qn> format):" -ForegroundColor Yellow
    $skippedFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
}
$created = @($log | Where-Object { $_.Status -eq 'Created' }).Count
Write-Host "$created template(s) created in $OutputFolder. Log: $logPath"
Write-Host "Next: open each .pbit (it loads by itself, no paths to enter), then File > Save as .pbix."
