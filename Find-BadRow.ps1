# ==============================
# CONFIG
# ==============================

$BasePath = "." # . means current folder
$BackupFile = Get-ChildItem "$BasePath\mdd_map_R*.ORIGINAL.xlsx" | Select-Object -First 1
$WorkingFile = $BackupFile.FullName -replace "\.ORIGINAL\.xlsx", ".xlsx"

$ProjectNum = ($WorkingFile -match "mdd_map_R(.+)\.xlsx") | Out-Null
$ProjectNum = $Matches[1]

# $Hawkeye = "S:\IPS\Voltron\LIONS\bin\exe\batch_stable\hawkeye.bat"
$FlatOut = "S:\IPS\Voltron\LIONS\bin\exe\batch_stable\flat_out"

$FlatOutToolArguments = @(
    "-a", "GexaFrontier_Unstacked"
)

$LogFile = "$BasePath\binary_search_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# not try to uncheck this - required to be always punched (for example, the row with ID); However, we are checking AN against 'id', so the row with ID should be skipped anyway, even if not listed here
$patternColBSkip = @(
    'Respondent\.ID'
    'CodingID'
    'uuid'
)

# ==============================
# LOG FUNCTION
# ==============================

function Write-Log($text) {
    $text | Tee-Object -FilePath $LogFile -Append
}

# ==============================
# GET X ROWS
# ==============================

function Get-XRows {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($BackupFile.FullName)
    $ws = $wb.Worksheets("variables")

    $lastRow = $ws.UsedRange.Rows.Count
    $xRows = @()
    $ExcludedByName = @()

    for ($row = 4; $row -le $lastRow; $row++) {

        $colBValue = $ws.Cells.Item($row, 2).Text # Column B
        $colAIValue = $ws.Cells.Item($row, 35).Text   # Column AI
        $colANValue = $ws.Cells.Item($row, 40).Text   # Column AN

        if ($colAIValue -eq "x") {

            $shouldExclude = $false
            foreach ($pattern in $patternColBSkip) {
                if ($colBValue -match $pattern) {
                    $shouldExclude = $true
                    break
                }
            }
            if ($colANValue -match '^\s*id\s*$') { # checking "format" against "id"
                $shouldExclude = $true
            }
            if ($shouldExclude) {
                $ExcludedByName += $row
            }
            else {
                $xRows += $row
            }
        }
    }

    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    return $xRows
}

# ==============================
# MODIFY EXCEL
# ==============================

function Modify-Excel($rowsToClear) {
    Copy-Item $BackupFile.FullName $WorkingFile -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($WorkingFile)
    $ws = $wb.Worksheets("variables")

    foreach ($row in $rowsToClear) {
        $ws.Cells.Item($row, 35).Value2 = "" # 35 is column AI - we are unsetting it
    }

    $wb.Save()
    $wb.Close($true)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# ==============================
# RUN TOOL
# ==============================

function Run-Tool {

    # the call is suppressed - I believe this is unnecessary
    # & $Hawkeye "R$ProjectNum.mdd" | Out-Null

    $output = & $FlatOut "R$ProjectNum.mdd" @FlatOutToolArguments 2>&1
    $text = $output | Out-String

    if ($text -match "\.\.\.aaaaaannd\.\.\.it's a photo finish!") {
        return @{ Status = "SUCCESS"; Output = $text }
    }

    if ($text -match "UH-HO!" -and $text -match "Exit:") {
        return @{ Status = "TARGET_FAILURE"; Output = $text }
    }

    if ($text -match "Exit:") {
        return @{ Status = "OTHER_FAILURE"; Output = $text }
    }

    return @{ Status = "UNKNOWN"; Output = $text }
}
# ==============================
# BINARY SEARCH
# ==============================

Write-Log "Project: $ProjectNum"
Write-Log "Started: $(Get-Date)"
Write-Log "======================================"

$xRows = Get-XRows
Write-Log "Found $($xRows.Count) rows marked with x"

$low = 0
$high = $xRows.Count - 1
$attempt = 1
$badRow = $null

while ($low -le $high) {

    $mid = [int](($low + $high) / 2)
    $testRows = $xRows[$low..$mid]

    Write-Log ""
    Write-Log "Attempt $attempt"
    Write-Log "Clearing rows: $($testRows -join ',')"

    Modify-Excel $testRows
    $result = Run-Tool

    switch ($result.Status) {

        "SUCCESS" {
            Write-Log "Result: SUCCESS"
            $badRow = $mid
            $high = $mid - 1
        }

        "TARGET_FAILURE" {
            Write-Log "Result: TARGET_FAILURE"
            $low = $mid + 1
        }

        "OTHER_FAILURE" {
            Write-Log "Result: OTHER_FAILURE - We broke something structural"
            Write-Log "STOP"
            Write-Log "Possibly unchecked some necessary row, like the row with ID."
            Write-Log "Aborting search to prevent corruption."
            break
        }

        "UNKNOWN" {
            Write-Log "Result: UNKNOWN - success or failure status not captured from outputs"
            Write-Log "STOP"
            Write-Log "Aborting search."
            break
        }
    }
    Write-Log "--------------------------------------"
    $attempt++
}

Write-Log ""
Write-Log "Finished: $(Get-Date)"

if ($badRow -ne $null) {
    Write-Log "Likely problematic Excel row: AI$($xRows[$badRow])"
    Write-Host "Likely problematic Excel row: AI$($xRows[$badRow])"
}
else {
    Write-Log "No single bad row found."
    Write-Host "No single bad row found."
}
 