<#
  Opens an exported workbook in Excel (hidden), reads the first formula cells of the dashboard sheet whose formula
  contains -Pattern, sets the dropdown cell whose list holds -Value to it, recalculates and reads them again.
  Prints one JSON line: { "repaired": bool, "before": [...], "after": [...] } (cell values as numbers / text).
  Used by tests/workbooks.test.js when TFX_EXCEL=1 (Windows with Excel only).
#>
param([string]$Xlsx, [string]$Value, [string]$Pattern = "Slicer Data", [int]$Count = 4)
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Xlsx); $xl.CalculateFull(); $ws = $wb.Worksheets.Item(1)
  $repaired = $wb.Windows.Item(1).Caption -like "*Repaired*"
  $cells = @(); $dropdown = $null
  foreach ($c in $ws.UsedRange.Cells) {
    if ($c.HasFormula -and $c.Formula -like "*$Pattern*" -and $cells.Count -lt $Count) { $cells += $c }
    if (-not $dropdown) {
      try {
        if ($c.Validation.Type -eq 3) {
          $list = @($xl.Evaluate($c.Validation.Formula1).Value2 | ForEach-Object { [string]$_ })
          if ($list -contains $Value) { $dropdown = $c }
        }
      } catch {}
    }
  }
  $read = { @($cells | ForEach-Object { $v = $_.Value2; if ($v -is [double]) { [math]::Round($v, 2) } else { [string]$_.Text } }) }
  $before = & $read
  if ($dropdown) { $dropdown.Value2 = $Value; $xl.CalculateFull() }
  $after = & $read
  @{ repaired = $repaired; dropdown = [bool]$dropdown; before = $before; after = $after } | ConvertTo-Json -Compress
  $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
