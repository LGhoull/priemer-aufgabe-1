
<# 
.SYNOPSIS
  Summarize ANUMMER_01..ANUMMER_10 columns into one list per BESTELLIDENT.

.DESCRIPTION
  - Reads a semicolon-delimited CSV.
  - Groups by BESTELLIDENT (handles multi-row orders).
  - Aggregates all non-empty & non-'NULL' ANUMMER_xx values into ANUMMER_LIST as [123,456,...].
  - Exports a cleaned CSV with semicolons.
  - By default, drops the original ANUMMER_xx columns (toggle with -KeepAnummerColumns).

.PARAMETER InputCsv
  Path to the input CSV file.

.PARAMETER OutputCsv
  Path to the output CSV file.

.EXAMPLE
  .\Summarize-Anummer.ps1 -InputCsv .\orders.csv -OutputCsv .\orders_summarized.csv

.EXAMPLE
  .\Summarize-Anummer.ps1 -InputCsv .\orders.csv -OutputCsv .\orders_summarized.csv -KeepAnummerColumns
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputCsv
)

$KeepAnummerColumns = $false;
# --- Config: ANUMMER columns (01..10) ---
$anumCols = 1..10 | ForEach-Object { "ANUMMER_{0:D2}" -f $_ }

# --- Import CSV (semicolon-delimited). Adjust -Encoding if needed (e.g., Default / ISO-8859-1) ---
try {
    $data = Import-Csv -Path $InputCsv -Delimiter ';' -Encoding UTF8
} catch {
    Write-Warning "UTF-8 import failed, retrying with Default encoding..."
    $data = Import-Csv -Path $InputCsv -Delimiter ';' -Encoding Default
}

if (-not $data) {
    Write-Error "No rows loaded from '$InputCsv'. Check path and delimiter."
    exit 1
}

# --- Helper: normalize values (trim, ignore empty/NULL) ---
function Get-ValidAnummers {
    param([object[]]$rows, [string[]]$columns)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($row in $rows) {
        foreach ($col in $columns) {
            if ($row.PSObject.Properties.Name -contains $col) {
                $val = $row.$col
                if ($null -ne $val) {
                    $val = "$val".Trim()
                    if ($val -and $val -ne 'NULL') {
                        $items.Add($val)
                    }
                }
            }
        }
    }
    # unique preserve order
    $seen = @{}
    $unique = foreach ($i in $items) {
        if (-not $seen.ContainsKey($i)) { $seen[$i] = $true; $i }
    }
    return $unique
}

# --- Optional: choose first non-empty per field across grouped rows (for non-ANUMMER fields) ---
function Merge-RowProperties {
    param([object[]]$rows, [string[]]$excludeCols)

    # Start from first row as base
    $base = $rows[0].PSObject.Copy()

    # Consider all property names present in any row
    $allProps = ($rows | ForEach-Object { $_.PSObject.Properties.Name }) | Select-Object -Unique

    foreach ($prop in $allProps) {
        if ($excludeCols -contains $prop) { continue }

        # If base is empty, try to fill from other rows
        $current = $base.$prop
        $currentStr = if ($null -ne $current) { "$current".Trim() } else { "" }

        if (-not $currentStr) {
            foreach ($r in $rows) {
                $val = $r.$prop
                $valStr = if ($null -ne $val) { "$val".Trim() } else { "" }
                if ($valStr) {
                    $base.$prop = $val
                    break
                }
            }
        }
    }

    return $base
}

# --- Group by BESTELLIDENT (if present); otherwise, process row-by-row ---
$groupKey = 'BESTELLIDENT'
$hasGroupKey = ($data | Select-Object -First 1).PSObject.Properties.Name -contains $groupKey

$results = New-Object System.Collections.Generic.List[object]

if ($hasGroupKey) {
    $grouped = $data | Group-Object -Property $groupKey
    foreach ($g in $grouped) {
        $rows = $g.Group

        # Build ANUMMER_LIST
        $anums = Get-ValidAnummers -rows $rows -columns $anumCols
        $anumList = if ($anums.Count -gt 0) { "[" + ($anums -join ',') + "]" } else { "[]" }

        # Merge other properties (prefer first row, fill blanks from rest)
        $merged = Merge-RowProperties -rows $rows -excludeCols $anumCols
        $merged | Add-Member -NotePropertyName 'ANUMMER_LIST' -NotePropertyValue $anumList -Force

        if (-not $KeepAnummerColumns) {
            foreach ($col in $anumCols) {
                if ($merged.PSObject.Properties.Name -contains $col) {
                    $merged.PSObject.Properties.Remove($col)
                }
            }
        }

        $results.Add($merged)
    }
}
else {
    # No BESTELLIDENT column -> just collapse ANUMMER_* per row
    foreach ($row in $data) {
        $anums = Get-ValidAnummers -rows @($row) -columns $anumCols
        $anumList = if ($anums.Count -gt 0) { "[" + ($anums -join ',') + "]" } else { "[]" }
        $row | Add-Member -NotePropertyName 'ANUMMER_LIST' -NotePropertyValue $anumList -Force

        if (-not $KeepAnummerColumns) {
            foreach ($col in $anumCols) {
                if ($row.PSObject.Properties.Name -contains $col) {
                    $row.PSObject.Properties.Remove($col)
                }
            }
        }
        $results.Add($row)
    }
}

# --- Export CSV (semicolon-delimited) ---
# UseQuotes AsNeeded is available in PowerShell 7+. If you're on Windows PowerShell 5.1, omit it.
try {
    $results | Export-Csv -Path $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8 -UseQuotes AsNeeded
} catch {
    Write-Warning "Export with UTF-8 failed, retrying with Default encoding..."
    $results | Export-Csv -Path $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding Default
}

Write-Host "Done. Wrote $($results.Count) summarized row(s) to '$OutputCsv'."
