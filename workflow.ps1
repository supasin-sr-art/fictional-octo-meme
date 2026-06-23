param(
  [string]$InputDir = "",
  [string]$OutputDir = "",
  [datetime]$StartDate = [datetime]::MinValue,
  [datetime]$EndDate = [datetime]::MinValue,
  [datetime]$TodayDate = [datetime]::Today,
  [switch]$ApplyPvFilterOnly,
  [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($InputDir)) { $InputDir = Join-Path $ProjectRoot "input" }
if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Join-Path $ProjectRoot "output" }

$xlUp = -4162
$xlValues = -4163
$xlPart = 2
$xlByRows = 1
$xlPrevious = 2
$xlCalculationAutomatic = -4105
$xlCalculationManual = -4135


function Write-ProgressLine($percent, $message) {
  $line = "PROGRESS|$percent|$message"
  try {
    [Console]::Out.WriteLine($line)
    [Console]::Out.Flush()
  }
  catch {
    Write-Output $line
  }
}

function Save-ExcelProcessId($excel, [string]$dir) {
  try {
    if (-not ("Blackwolf.NativeMethods" -as [type])) {
      Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace Blackwolf {
  public static class NativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
  }
}
"@
    }
    [uint32]$excelPid = 0
    [void][Blackwolf.NativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    if ($excelPid -gt 0) {
      New-Item -ItemType Directory -Force -Path $dir | Out-Null
      Set-Content -LiteralPath (Join-Path $dir 'excel.pid') -Value ([string]$excelPid) -Encoding ASCII
    }
  } catch {}
}

function Set-ExcelTurboMode($excel) {
  try { $excel.Visible = $false } catch {}
  try { $excel.DisplayAlerts = $false } catch {}
  try { $excel.ScreenUpdating = $false } catch {}
  try { $excel.EnableEvents = $false } catch {}
  try { $excel.DisplayStatusBar = $false } catch {}
  try { $excel.AskToUpdateLinks = $false } catch {}
  try { $excel.AlertBeforeOverwriting = $false } catch {}
  try { $excel.AutomationSecurity = 3 } catch {}
  try { $excel.Calculation = $xlCalculationManual } catch {}
  try { $excel.CalculateBeforeSave = $false } catch {}
}

function Get-MatrixValue($data, $row, $col) {
  if ($null -eq $data) { return $null }
  if ($data -is [System.Array]) {
    try { return $data.GetValue($row, $col) } catch { return $null }
  }
  if ($row -eq 1 -and $col -eq 1) { return $data }
  return $null
}

function Normalize-Header($value) {
  if ($null -eq $value) { return "" }
  return ([string]$value -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
}

function Normalize-Id($value) {
  if ($null -eq $value) { return "" }
  if ($value -is [double] -or $value -is [single] -or $value -is [decimal] -or $value -is [int] -or $value -is [long]) {
    return ([decimal]$value).ToString("0")
  }
  $text = ([string]$value).Trim()
  if ($text -eq "") { return "" }
  if ($text -match "^\d+(\.0+)?$") {
    return ([decimal]$text).ToString("0")
  }
  return $text
}

function Convert-ToExcelIdValue($id) {
  if ($null -eq $id -or $id -eq "") { return $null }
  if ($id -match "^\d+$") { return [double]$id }
  return $id
}

function Convert-ToExcelText($value) {
  $text = Normalize-Id $value
  if ($text -eq "") { return $null }
  return "'" + $text
}

function Convert-ExcelDate($value) {
  if ($null -eq $value -or $value -eq "") { return $null }
  if ($value -is [datetime]) { return $value }
  if ($value -is [double] -or $value -is [single] -or $value -is [decimal] -or $value -is [int] -or $value -is [long]) {
    return [datetime]::FromOADate([double]$value)
  }
  $text = ([string]$value).Trim()
  if ($text -eq "") { return $null }
  return [datetime]::Parse($text)
}

function Find-OneFile($pattern, $required) {
  $files = @(Get-ChildItem -LiteralPath $InputDir -File -Filter $pattern | Sort-Object LastWriteTime -Descending)
  if ($files.Count -eq 0) {
    if ($required) { throw "Missing required input file: $pattern" }
    return $null
  }
  return $files[0]
}

function Repair-XlsxTableMetadata($path) {
  if (-not (Test-Path -LiteralPath $path)) { return 0 }

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $changedEntries = 0
  $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $tableEntries = @($archive.Entries | Where-Object { $_.FullName -like "xl/tables/*.xml" })
    foreach ($entry in $tableEntries) {
      $reader = New-Object System.IO.StreamReader($entry.Open())
      $xml = $reader.ReadToEnd()
      $reader.Close()

      $updated = $xml
      $updated = [regex]::Replace($updated, "<sortState\b[\s\S]*?</sortState>", "")

      $tableColumnsMatch = [regex]::Match($updated, "<tableColumns\b[^>]*>[\s\S]*?</tableColumns>")
      if ($tableColumnsMatch.Success) {
        $script:RepairTableColumnIndex = 0
        $fixedColumns = [regex]::Replace(
          $tableColumnsMatch.Value,
          "<tableColumn\b([^>]*?)\bid=""\d+""",
          {
            param($match)
            $script:RepairTableColumnIndex++
            return [regex]::Replace($match.Value, "id=""\d+""", "id=""$script:RepairTableColumnIndex""")
          }
        )
        $fixedColumns = [regex]::Replace($fixedColumns, "<tableColumns\b([^>]*)\bcount=""\d+""", "<tableColumns`$1count=""$script:RepairTableColumnIndex""")
        Remove-Variable -Name RepairTableColumnIndex -Scope Script -ErrorAction SilentlyContinue
        $updated = $updated.Substring(0, $tableColumnsMatch.Index) + $fixedColumns + $updated.Substring($tableColumnsMatch.Index + $tableColumnsMatch.Length)
      }

      if ($updated -ne $xml) {
        $entryName = $entry.FullName
        $entry.Delete()
        $newEntry = $archive.CreateEntry($entryName)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
        $writer.Write($updated)
        $writer.Close()
        $changedEntries++
      }
    }
  }
  finally {
    $archive.Dispose()
  }

  return $changedEntries
}

function Read-ZipEntryText($archive, $entryName) {
  $entry = $archive.GetEntry($entryName)
  if ($null -eq $entry) { throw "Missing workbook XML entry: $entryName" }
  $reader = New-Object System.IO.StreamReader($entry.Open())
  try {
    return $reader.ReadToEnd()
  }
  finally {
    $reader.Close()
  }
}

function Write-ZipEntryText($archive, $entryName, $text) {
  $entry = $archive.GetEntry($entryName)
  if ($null -ne $entry) { $entry.Delete() }
  $newEntry = $archive.CreateEntry($entryName)
  $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
  try {
    $writer.Write($text)
  }
  finally {
    $writer.Close()
  }
}

function Get-XlsxWorksheetEntryName($archive, $sheetName) {
  [xml]$workbookXml = Read-ZipEntryText $archive "xl/workbook.xml"
  [xml]$relationshipsXml = Read-ZipEntryText $archive "xl/_rels/workbook.xml.rels"

  $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
  $workbookNs.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $workbookNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $sheetNode = $workbookXml.SelectSingleNode("//m:sheet[@name='$sheetName']", $workbookNs)
  if ($null -eq $sheetNode) { throw "Cannot find worksheet XML mapping: $sheetName" }

  $relationshipId = $sheetNode.GetAttribute(
    "id",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  )
  $relationshipsNs = New-Object System.Xml.XmlNamespaceManager($relationshipsXml.NameTable)
  $relationshipsNs.AddNamespace("p", "http://schemas.openxmlformats.org/package/2006/relationships")
  $relationshipNode = $relationshipsXml.SelectSingleNode(
    "//p:Relationship[@Id='$relationshipId']",
    $relationshipsNs
  )
  if ($null -eq $relationshipNode) { throw "Cannot find worksheet relationship: $relationshipId" }

  $target = [string]$relationshipNode.Target
  if ($target.StartsWith("/")) { return $target.TrimStart("/") }
  return "xl/" + $target.TrimStart("/")
}

function Suspend-XlsxSheetProtection($path, $sheetName) {
  if (-not (Test-Path -LiteralPath $path)) { throw "Workbook not found: $path" }

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $entryName = Get-XlsxWorksheetEntryName $archive $sheetName
    $xml = Read-ZipEntryText $archive $entryName
    $match = [regex]::Match($xml, "<sheetProtection\b[^>]*/>")
    if (-not $match.Success) { return $null }

    $updated = $xml.Remove($match.Index, $match.Length)
    Write-ZipEntryText $archive $entryName $updated
    return @{
      EntryName = $entryName
      ProtectionXml = $match.Value
    }
  }
  finally {
    $archive.Dispose()
  }
}

function Restore-XlsxSheetProtection($path, $protectionInfo) {
  if ($null -eq $protectionInfo) { return }

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $entryName = [string]$protectionInfo.EntryName
    $protectionXml = [string]$protectionInfo.ProtectionXml
    $xml = Read-ZipEntryText $archive $entryName

    if ($xml -match "<sheetProtection\b") {
      $updated = [regex]::Replace($xml, "<sheetProtection\b[^>]*/>", $protectionXml, 1)
    }
    elseif ($xml.Contains("</sheetData>")) {
      $updated = $xml.Replace("</sheetData>", "</sheetData>$protectionXml")
    }
    else {
      throw "Cannot restore sheet protection in: $entryName"
    }

    Write-ZipEntryText $archive $entryName $updated
  }
  finally {
    $archive.Dispose()
  }
}

function Ensure-XlsxNonIssueStatusFormula($path) {
  if (-not (Test-Path -LiteralPath $path)) { throw "Workbook not found: $path" }

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $changedEntries = 0
  $matchedEntries = 0
  $alreadyUpdated = $false
  $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    # Do not assume table1.xml or sheet1.xml. Excel can renumber table/sheet XML
    # entries after a save, repair, or template change. Search all table and sheet XML.
    $entryNames = @(
      $archive.Entries |
        Where-Object { $_.FullName -like "xl/tables/*.xml" -or $_.FullName -like "xl/worksheets/*.xml" } |
        ForEach-Object { $_.FullName }
    )

    foreach ($entryName in $entryNames) {
      $entry = $archive.GetEntry($entryName)
      if ($null -eq $entry) { continue }

      $reader = New-Object System.IO.StreamReader($entry.Open())
      $xml = $reader.ReadToEnd()
      $reader.Close()

      $hasOldDateEProp = $xml.Contains('Table1[[#This Row],[CreateDate]], Table1[[#This Row],[EPropID]]')
      $hasOldDiscount = $xml.Contains('Table1[[#This Row],[Discount]])&lt;12')
      $hasNewDateStatusEProp = $xml.Contains('Table1[[#This Row],[CreateDate]], Table1[[#This Row],[Status]], Table1[[#This Row],[EPropID]]')
      $hasNewDiscount = $xml.Contains('Table1[[#This Row],[Discount]])&lt;13')

      if ($hasOldDateEProp -or $hasOldDiscount -or $hasNewDateStatusEProp -or $hasNewDiscount) {
        $matchedEntries++
      }
      if ($hasNewDateStatusEProp -and $hasNewDiscount) {
        $alreadyUpdated = $true
      }

      $updated = $xml.Replace(
        'Table1[[#This Row],[CreateDate]], Table1[[#This Row],[EPropID]]',
        'Table1[[#This Row],[CreateDate]], Table1[[#This Row],[Status]], Table1[[#This Row],[EPropID]]'
      )
      $updated = $updated.Replace(
        'Table1[[#This Row],[Discount]])&lt;12',
        'Table1[[#This Row],[Discount]])&lt;13'
      )

      if ($updated -ne $xml) {
        $entry.Delete()
        $newEntry = $archive.CreateEntry($entryName)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
        $writer.Write($updated)
        $writer.Close()
        $changedEntries++
      }
    }
  }
  finally {
    $archive.Dispose()
  }

  # A template can legitimately have a different table name/formula layout.
  # Continue instead of stopping the whole workflow; Excel validation later will
  # still detect a genuinely unusable workbook.
  if ($changedEntries -eq 0 -and -not $alreadyUpdated -and $matchedEntries -eq 0) {
    Write-Output "WARN: Non-issue formula XML pattern was not found; skipped formula patch safely."
  }

  return $changedEntries
}

function Get-HeaderMap($ws, $row, $maxCol) {
  $map = @{}
  for ($col = 1; $col -le $maxCol; $col++) {
    $key = Normalize-Header $ws.Cells.Item($row, $col).Value2
    if ($key -ne "" -and -not $map.ContainsKey($key)) {
      $map[$key] = $col
    }
  }
  return $map
}


function Get-HeaderMapFromListObject($table) {
  $map = @{}
  for ($i = 1; $i -le $table.ListColumns.Count; $i++) {
    $column = $table.ListColumns.Item($i)
    $key = Normalize-Header $column.Name
    if ($key -ne "" -and -not $map.ContainsKey($key)) {
      $map[$key] = $column.Range.Column
    }
  }
  return $map
}

function Get-WorksheetHeaderInfo($ws, $headerRow, $fallbackMaxCol) {
  if ($fallbackMaxCol -lt 1) { $fallbackMaxCol = 80 }
  $lastRow = [Math]::Max($headerRow, 1)
  $lastCol = $fallbackMaxCol
  $lastMap = @{}

  for ($try = 1; $try -le 12; $try++) {
    try {
      $used = $ws.UsedRange
      $candidateLastRow = $used.Row + $used.Rows.Count - 1
      $candidateLastCol = $used.Column + $used.Columns.Count - 1
      if ($candidateLastRow -ge $headerRow) { $lastRow = $candidateLastRow }
      if ($candidateLastCol -ge 1) { $lastCol = [Math]::Min([Math]::Max($candidateLastCol, $fallbackMaxCol), 120) }
    }
    catch {
      $lastRow = [Math]::Max($lastRow, $headerRow)
      $lastCol = [Math]::Max($lastCol, $fallbackMaxCol)
    }

    try {
      $lastMap = Get-HeaderMap $ws $headerRow $lastCol
      if ($lastMap.Count -gt 0) {
        return @{ Map = $lastMap; LastRow = $lastRow; LastCol = $lastCol }
      }
    }
    catch {}

    Start-Sleep -Milliseconds ([Math]::Min(1500, 150 * $try))
  }

  return @{ Map = $lastMap; LastRow = $lastRow; LastCol = $lastCol }
}
function Require-Col($map, $name) {
  if (-not $map.ContainsKey($name)) {
    $available = @($map.Keys | Sort-Object) -join ", "
    throw "Missing required column: $name. Available headers: $available"
  }
  return $map[$name]
}

function Find-HeaderCell($ws, $headerText) {
  $lastRow = 40
  $lastCol = 80
  for ($try = 1; $try -le 8; $try++) {
    try {
      $used = $ws.UsedRange
      $lastRow = [Math]::Max($used.Row + $used.Rows.Count - 1, 40)
      $lastCol = [Math]::Min([Math]::Max($used.Column + $used.Columns.Count - 1, 80), 120)
      $scanLastRow = [Math]::Min($lastRow, 40)
      for ($row = 1; $row -le $scanLastRow; $row++) {
        for ($col = 1; $col -le $lastCol; $col++) {
          if ((Normalize-Header $ws.Cells.Item($row, $col).Value2) -eq $headerText) {
            return @{ Row = $row; Col = $col; LastRow = $lastRow; LastCol = $lastCol }
          }
        }
      }
    }
    catch {}
    Start-Sleep -Milliseconds ([Math]::Min(1500, 150 * $try))
  }
  return $null
}

function Close-WorkbookSafe($wb, $saveChanges) {
  for ($try = 1; $try -le 5; $try++) {
    try {
      $wb.Close($saveChanges)
      return
    }
    catch {
      Start-Sleep -Milliseconds (400 * $try)
    }
  }
  try { $wb.Close($saveChanges) } catch {}
}

function Open-WorkbookSafe($excel, $path, $readOnly) {
  for ($try = 1; $try -le 6; $try++) {
    try {
      $wb = $excel.Workbooks.Open($path, 0, $readOnly)
      if ($null -ne $wb) {
        for ($wait = 1; $wait -le 30; $wait++) {
          try {
            if ($wb.Worksheets.Count -gt 0) {
              Write-Output -NoEnumerate $wb
              return
            }
          }
          catch {}
          Start-Sleep -Milliseconds 300
        }
        Close-WorkbookSafe $wb $false
      }
    }
    catch {
      Start-Sleep -Milliseconds (500 * $try)
    }
    Start-Sleep -Milliseconds (300 * $try)
  }
  throw "Cannot open workbook: $path"
}

function Get-WorksheetSafe($wb, $sheetName) {
  for ($try = 1; $try -le 8; $try++) {
    try {
      $ws = $wb.Worksheets.Item($sheetName)
      if ($null -ne $ws) {
        Write-Output -NoEnumerate $ws
        return
      }
    }
    catch {}
    Start-Sleep -Milliseconds (250 * $try)
  }
  throw "Cannot find worksheet: $sheetName"
}

function Extract-M190PropIds($excel, $file) {
  $ids = New-Object System.Collections.ArrayList
  $wb = Open-WorkbookSafe $excel $file.FullName $true
  try {
    $ws = Get-WorksheetSafe $wb "Policy Detail"
    $found = Find-HeaderCell $ws "Prop Id"
    if ($null -eq $found) { throw "Cannot find Prop Id header in $($file.Name)" }
    $rowCount = [Math]::Max(0, $found.LastRow - $found.Row)
    if ($rowCount -gt 0) {
      $data = $ws.Range(
        $ws.Cells.Item($found.Row + 1, $found.Col),
        $ws.Cells.Item($found.LastRow, $found.Col)
      ).Value2
      for ($i = 1; $i -le $rowCount; $i++) {
        $id = Normalize-Id (Get-MatrixValue $data $i 1)
        if ($id -ne "") { [void]$ids.Add($id) }
      }
    }
  }
  finally {
    Close-WorkbookSafe $wb $false
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }
  return @($ids.ToArray())
}

function Extract-CPropIds($excel, $file) {
  $ids = New-Object System.Collections.ArrayList
  if ($null -eq $file) { return @($ids.ToArray()) }

  $wb = Open-WorkbookSafe $excel $file.FullName $true
  try {
    foreach ($ws in $wb.Worksheets) {
      $found = Find-HeaderCell $ws "CPROP_ID"
      if ($null -ne $found) {
        $rowCount = [Math]::Max(0, $found.LastRow - $found.Row)
        if ($rowCount -gt 0) {
          $data = $ws.Range(
            $ws.Cells.Item($found.Row + 1, $found.Col),
            $ws.Cells.Item($found.LastRow, $found.Col)
          ).Value2
          for ($i = 1; $i -le $rowCount; $i++) {
            $id = Normalize-Id (Get-MatrixValue $data $i 1)
            if ($id -ne "") { [void]$ids.Add($id) }
          }
        }
        return @($ids.ToArray())
      }
    }
  }
  finally {
    Close-WorkbookSafe $wb $false
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }

  throw "Cannot find CPROP_ID header in $($file.Name)"
}

function Extract-EtlPropIds($file) {
  $records = Extract-EtlRecords $file
  return @($records | ForEach-Object { $_.PropId })
}

function Extract-EtlRecords($file) {
  $records = New-Object System.Collections.ArrayList
  if ($null -eq $file) { return @($records.ToArray()) }

  $lines = Get-Content -LiteralPath $file.FullName
  foreach ($line in $lines) {
    if ($line -match "^\s*(\d+)\.(\d+):([^:]+):(.+?)\s*$") {
      [void]$records.Add([pscustomobject]@{
        No = [int]$matches[1]
        PropId = $matches[2]
        Policy = $matches[3].Trim()
        Group = $matches[4].Trim()
      })
    }
  }
  return @($records.ToArray())
}

function Extract-ReportRows($excel, $file, $startDate, $endDate) {
  $allowedStatus = @{
    "เสร็จสมบูรณ์" = $true
    "เสร็จสมบูรณ์(ติดปัญหา Upload File)" = $true
    "ระบบขัดข้องกรุณาติดต่อไอที" = $true
    "" = $true
  }

  # Large Data Mode: อ่านเป็นช่วงเพื่อลดหน่วยความจำ ไม่โหลด 5 แสนแถวพร้อมกัน
  $chunkRows = 20000
  $rows = New-Object System.Collections.ArrayList
  $wb = Open-WorkbookSafe $excel $file.FullName $true
  try {
    $ws = Get-WorksheetSafe $wb "Data"
    $headerInfo = Get-WorksheetHeaderInfo $ws 1 80
    $lastRow = [int]$headerInfo["LastRow"]
    $headers = $headerInfo["Map"]

    $cAgencyCode = Require-Col $headers "AgencyCode"
    $cMticode = Require-Col $headers "Mticode"
    $cAgencyName = Require-Col $headers "AgencyName"
    $cRequestCode = Require-Col $headers "RequestCode"
    $cEmployeeName = Require-Col $headers "employeeName"
    $cAlienCode = Require-Col $headers "alienCode"
    $cAlienNameEn = Require-Col $headers "alienNameEn"
    $cCertificateNo = Require-Col $headers "CertificateNo"
    $cTotalPremium = Require-Col $headers "TotalPremium"
    $cProposalID = Require-Col $headers "ProposalID"
    $cCreateDate = Require-Col $headers "CreateDate"
    $cStatus = Require-Col $headers "Status"
    $cEPropID = Require-Col $headers "EPropID"
    $cDiscount = Require-Col $headers "Discount"

    $requiredCols = @(
      $cAgencyCode, $cMticode, $cAgencyName, $cRequestCode, $cEmployeeName,
      $cAlienCode, $cAlienNameEn, $cCertificateNo, $cTotalPremium,
      $cProposalID, $cCreateDate, $cStatus, $cEPropID, $cDiscount
    )
    $readLastCol = [int](($requiredCols | Measure-Object -Maximum).Maximum)
    $totalRows = [Math]::Max(0, $lastRow - 1)
    if ($totalRows -eq 0) { return @($rows.ToArray()) }

    $processedRows = 0
    for ($chunkStart = 2; $chunkStart -le $lastRow; $chunkStart += $chunkRows) {
      $chunkEnd = [Math]::Min($lastRow, $chunkStart + $chunkRows - 1)
      $chunkCount = $chunkEnd - $chunkStart + 1
      $data = $ws.Range(
        $ws.Cells.Item($chunkStart, 1),
        $ws.Cells.Item($chunkEnd, $readLastCol)
      ).Value2

      for ($i = 1; $i -le $chunkCount; $i++) {
        $createDate = Convert-ExcelDate (Get-MatrixValue $data $i $cCreateDate)
        if ($null -eq $createDate) { continue }
        if ($createDate.Date -lt $startDate.Date -or $createDate.Date -gt $endDate.Date) { continue }

        $status = Normalize-Header (Get-MatrixValue $data $i $cStatus)
        if (-not $allowedStatus.ContainsKey($status)) { continue }

        $certificateNo = Get-MatrixValue $data $i $cCertificateNo
        $policy = ""
        if ($null -ne $certificateNo) {
          $certText = ([string]$certificateNo).Trim()
          if ($certText.Length -ge 8) { $policy = $certText.Substring(0, 8) }
        }

        $proposalId = Normalize-Id (Get-MatrixValue $data $i $cProposalID)
        if ($proposalId -eq "") { continue }

        $excelRow = @(
          (Get-MatrixValue $data $i $cAgencyCode),
          (Get-MatrixValue $data $i $cMticode),
          (Get-MatrixValue $data $i $cAgencyName),
          (Convert-ToExcelText (Get-MatrixValue $data $i $cRequestCode)),
          (Get-MatrixValue $data $i $cEmployeeName),
          (Get-MatrixValue $data $i $cAlienCode),
          (Get-MatrixValue $data $i $cAlienNameEn),
          $certificateNo,
          $policy,
          (Get-MatrixValue $data $i $cTotalPremium),
          (Convert-ToExcelText $proposalId),
          $createDate.ToOADate(),
          $status,
          (Convert-ToExcelText (Get-MatrixValue $data $i $cEPropID)),
          (Get-MatrixValue $data $i $cDiscount)
        )

        [void]$rows.Add([pscustomobject]@{
          ProposalId = $proposalId
          Values = $excelRow
        })
      }

      $processedRows += $chunkCount
      $pct = 34 + [int][Math]::Floor(([double]$processedRows / [double]$totalRows) * 8)
      if ($pct -gt 42) { $pct = 42 }
      Write-ProgressLine $pct "กำลังอ่าน Daily Report แบบ Large Data: $processedRows / $totalRows แถว"
      $data = $null
      [GC]::Collect(0)
    }
  }
  finally {
    Close-WorkbookSafe $wb $false
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }

  return @($rows.ToArray())
}

function Get-MasterStartDate($excel, $file) {
  $wb = Open-WorkbookSafe $excel $file.FullName $true
  try {
    $ws = Get-WorksheetSafe $wb "Data"
    $table = $ws.ListObjects.Item("Table1")
    $lastRow = $table.Range.Row + $table.Range.Rows.Count - 1
    $headers = Get-HeaderMapFromListObject $table
    $cCreateDate = Require-Col $headers "CreateDate"
    $cProposalId = Require-Col $headers "ProposalID"
    $rowCount = [Math]::Max(0, $lastRow - 1)
    if ($rowCount -eq 0) { throw "Cannot determine start date from master Data sheet" }

    $proposalData = $ws.Range($ws.Cells.Item(2, $cProposalId), $ws.Cells.Item($lastRow, $cProposalId)).Value2
    $dateData = $ws.Range($ws.Cells.Item(2, $cCreateDate), $ws.Cells.Item($lastRow, $cCreateDate)).Value2
    $maxDate = $null
    for ($i = 1; $i -le $rowCount; $i++) {
      $proposalId = Normalize-Id (Get-MatrixValue $proposalData $i 1)
      if ($proposalId -eq "") { continue }
      $createDate = Convert-ExcelDate (Get-MatrixValue $dateData $i 1)
      if ($null -eq $createDate) { continue }
      if ($null -eq $maxDate -or $createDate.Date -gt $maxDate) { $maxDate = $createDate.Date }
    }
    if ($null -eq $maxDate) { throw "Cannot determine start date from master Data sheet" }
    return $maxDate
  }
  finally {
    Close-WorkbookSafe $wb $false
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }
}

function Get-ReportDateRange($excel, $file) {
  $chunkRows = 50000
  $wb = Open-WorkbookSafe $excel $file.FullName $true
  try {
    $ws = Get-WorksheetSafe $wb "Data"
    $headerInfo = Get-WorksheetHeaderInfo $ws 1 80
    $lastRow = [int]$headerInfo["LastRow"]
    $headers = $headerInfo["Map"]
    $cCreateDate = Require-Col $headers "CreateDate"
    $rowCount = [Math]::Max(0, $lastRow - 1)
    if ($rowCount -eq 0) { throw "Cannot determine date range from report Data sheet" }

    $minDate = $null
    $maxDate = $null
    $validDateRows = 0
    for ($chunkStart = 2; $chunkStart -le $lastRow; $chunkStart += $chunkRows) {
      $chunkEnd = [Math]::Min($lastRow, $chunkStart + $chunkRows - 1)
      $chunkCount = $chunkEnd - $chunkStart + 1
      $dateData = $ws.Range(
        $ws.Cells.Item($chunkStart, $cCreateDate),
        $ws.Cells.Item($chunkEnd, $cCreateDate)
      ).Value2
      for ($i = 1; $i -le $chunkCount; $i++) {
        $createDate = Convert-ExcelDate (Get-MatrixValue $dateData $i 1)
        if ($null -eq $createDate) { continue }
        $d = $createDate.Date
        $validDateRows++
        if ($null -eq $minDate -or $d -lt $minDate) { $minDate = $d }
        if ($null -eq $maxDate -or $d -gt $maxDate) { $maxDate = $d }
      }
      $dateData = $null
    }
    if ($null -eq $minDate -or $null -eq $maxDate) { throw "Cannot determine date range from report Data sheet" }
    return [pscustomobject]@{
      MinDate = $minDate
      MaxDate = $maxDate
      ValidDateRows = $validDateRows
      WorksheetRows = $rowCount
    }
  }
  finally {
    Close-WorkbookSafe $wb $false
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }
}

function Get-ReportLatestDate($excel, $file) {
  return (Get-ReportDateRange $excel $file).MaxDate
}

function Write-2DRange($range, $rows) {
  if ($rows.Count -eq 0) { return }
  $chunkSize = 10000
  $colCount = $rows[0].Count
  $ws = $range.Worksheet
  $baseRow = $range.Row
  $baseCol = $range.Column

  for ($offset = 0; $offset -lt $rows.Count; $offset += $chunkSize) {
    $count = [Math]::Min($chunkSize, $rows.Count - $offset)
    $data = New-Object "object[,]" $count, $colCount
    for ($r = 0; $r -lt $count; $r++) {
      $sourceRow = $rows[$offset + $r]
      for ($c = 0; $c -lt $colCount; $c++) {
        $data[$r, $c] = $sourceRow[$c]
      }
    }
    $dest = $ws.Range(
      $ws.Cells.Item($baseRow + $offset, $baseCol),
      $ws.Cells.Item($baseRow + $offset + $count - 1, $baseCol + $colCount - 1)
    )
    $dest.Value2 = $data
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dest)
    $data = $null
  }
}

function Update-Table1($wb, $pendingRows) {
  $ws = $wb.Worksheets.Item("Data")
  $table = $ws.ListObjects.Item("Table1")
  $oldLastRow = $table.Range.Row + $table.Range.Rows.Count - 1
  $dataCount = $pendingRows.Count

  if ($oldLastRow -ge 2) {
    [void]$ws.Range("A2:O$oldLastRow").ClearContents()
  }

  $requiredLastRow = [Math]::Max($dataCount, 1) + 1
  if ($oldLastRow -lt $requiredLastRow) {
    try {
      [void]$table.Resize($ws.Range("A1:W$requiredLastRow"))
      $oldLastRow = $table.Range.Row + $table.Range.Rows.Count - 1
    }
    catch {
      $capacity = [Math]::Max($oldLastRow - 1, 0)
      throw "Data Table1 has capacity $capacity rows, but needs $dataCount rows. Expand Table1 / unlock table expansion before running."
    }
  }

  if ($dataCount -gt 0) {
    $values = New-Object System.Collections.ArrayList
    foreach ($item in $pendingRows) { [void]$values.Add($item.Values) }
    try { $ws.Range("D2:D$oldLastRow").NumberFormat = "@" } catch {}
    try { $ws.Range("J2:J$oldLastRow").NumberFormat = "#,##0" } catch {}
    try { $ws.Range("K2:K$oldLastRow").NumberFormat = "@" } catch {}
    try { $ws.Range("N2:N$oldLastRow").NumberFormat = "@" } catch {}
    Write-2DRange $ws.Range("A2") $values
  }

  if ($oldLastRow -ge 2) {
    try { $ws.Range("D2:D$oldLastRow").NumberFormat = "@" } catch {}
    try { $ws.Range("J2:J$oldLastRow").NumberFormat = "#,##0" } catch {}
    try { $ws.Range("K2:K$oldLastRow").NumberFormat = "@" } catch {}
    try { $ws.Range("N2:N$oldLastRow").NumberFormat = "@" } catch {}
    try { $ws.Range("L2:L$oldLastRow").NumberFormat = "d/m/yyyy h:mm" } catch {}
  }
}

function Update-MenuEStatus($wb, $pendingRows, $smIds, $blIds, $issueCutoffDate) {
  $ws = $wb.Worksheets.Item("Data")
  $table = $ws.ListObjects.Item("Table1")
  $oldLastRow = $table.Range.Row + $table.Range.Rows.Count - 1
  $menuECol = $table.ListColumns.Item("ติดปัญหาไม่เข้าในเมนู E").Range.Column
  if ($oldLastRow -ge 2) {
    [void]$ws.Range($ws.Cells.Item(2, $menuECol), $ws.Cells.Item($oldLastRow, $menuECol)).ClearContents()
  }

  if ($null -eq $issueCutoffDate -or $issueCutoffDate -eq [datetime]::MinValue) {
    return @{ MenuEWritten = 0; MenuEClearedRows = [Math]::Max(0, $oldLastRow - 1); MenuEReferenceDate = "" }
  }

  $referenceDate = $issueCutoffDate.Date
  $smSet = @{}
  foreach ($id in $smIds) { $norm = Normalize-Id $id; if ($norm -ne "") { $smSet[$norm] = $true } }
  $blSet = @{}
  foreach ($id in $blIds) { $norm = Normalize-Id $id; if ($norm -ne "") { $blSet[$norm] = $true } }

  $written = 0
  $count = $pendingRows.Count
  if ($count -gt 0) {
    $values = New-Object "object[,]" $count, 1
    for ($i = 0; $i -lt $count; $i++) {
      $item = $pendingRows[$i]
      $proposalId = Normalize-Id $item.ProposalId
      $value = $null
      if ($proposalId -ne "" -and -not $smSet.ContainsKey($proposalId) -and -not $blSet.ContainsKey($proposalId)) {
        $createValue = $item.Values[11]
        if ($null -ne $createValue -and $createValue -ne "") {
          $createDate = Convert-ExcelDate $createValue
          if ($null -ne $createDate -and $createDate.Date -lt $referenceDate) {
            $value = "ติดปัญหาไม่เข้าในเมนู E"
            $written++
          }
        }
      }
      $values[$i, 0] = $value
    }
    $ws.Range($ws.Cells.Item(2, $menuECol), $ws.Cells.Item($count + 1, $menuECol)).Value2 = $values
  }

  return @{
    MenuEWritten = $written
    MenuEClearedRows = [Math]::Max(0, $oldLastRow - 1)
    MenuEReferenceDate = $referenceDate.ToString("yyyy-MM-dd")
  }
}

function Update-IssueWorkbook($excel, $issueFile, $pendingRows, $m190Ids, $etlRecords) {
  $wb = Open-WorkbookSafe $excel $issueFile.FullName $false
  $saveWorkbook = $false
  try {
    $wsData = $wb.Worksheets.Item("Data")
    $dataTable = $wsData.ListObjects.Item("Table1")
    $oldDataLastRow = $dataTable.Range.Row + $dataTable.Range.Rows.Count - 1
    if ($oldDataLastRow -ge 2) { [void]$wsData.Range("A2:P$oldDataLastRow").ClearContents() }
    $dataCount = $pendingRows.Count
    $requiredDataLastRow = [Math]::Max($dataCount, 1) + 1
    [void]$dataTable.Resize($wsData.Range("A1:P$requiredDataLastRow"))
    if ($dataCount -gt 0) {
      $dataValues = New-Object System.Collections.ArrayList
      foreach ($item in $pendingRows) {
        $rowValues = New-Object object[] 16
        for ($j = 0; $j -lt 15; $j++) { $rowValues[$j] = $item.Values[$j] }
        $rowValues[15] = "#N/A"
        [void]$dataValues.Add($rowValues)
      }
      try { $wsData.Range("D2:D$requiredDataLastRow").NumberFormat = "@" } catch {}
      try { $wsData.Range("J2:J$requiredDataLastRow").NumberFormat = "#,##0" } catch {}
      try { $wsData.Range("K2:K$requiredDataLastRow").NumberFormat = "@" } catch {}
      try { $wsData.Range("L2:L$requiredDataLastRow").NumberFormat = "d/m/yyyy h:mm" } catch {}
      try { $wsData.Range("N2:N$requiredDataLastRow").NumberFormat = "@" } catch {}
      Write-2DRange $wsData.Range("A2") $dataValues
    }

    $wsCheck = $wb.Worksheets.Item("Check")
    $checkTable = $wsCheck.ListObjects.Item("Table2")
    $oldCheckLastRow = $checkTable.Range.Row + $checkTable.Range.Rows.Count - 1
    if ($oldCheckLastRow -ge 2) { [void]$wsCheck.Range("A2:B$oldCheckLastRow").ClearContents() }
    $checkIds = @($m190Ids) + @($etlRecords | ForEach-Object { $_.PropId })
    $checkCount = $checkIds.Count
    $requiredCheckLastRow = [Math]::Max($checkCount, 1) + 1
    [void]$checkTable.Resize($wsCheck.Range("A1:B$requiredCheckLastRow"))
    if ($checkCount -gt 0) {
      $checkValues = New-Object System.Collections.ArrayList
      foreach ($id in $checkIds) { [void]$checkValues.Add(@((Convert-ToExcelText $id), "ออกกรมธรรม์")) }
      try { $wsCheck.Range("A2:A$requiredCheckLastRow").NumberFormat = "@" } catch {}
      Write-2DRange $wsCheck.Range("A2") $checkValues
    }

    $wsEtl = $wb.Worksheets.Item("ETL")
    $etlTable = $wsEtl.ListObjects.Item("Table4")
    $oldEtlLastRow = $etlTable.Range.Row + $etlTable.Range.Rows.Count - 1
    if ($oldEtlLastRow -ge 2) { [void]$wsEtl.Range("A2:D$oldEtlLastRow").ClearContents() }
    $etlCount = $etlRecords.Count
    $requiredEtlLastRow = [Math]::Max($etlCount, 1) + 1
    [void]$etlTable.Resize($wsEtl.Range("A1:D$requiredEtlLastRow"))
    if ($etlCount -gt 0) {
      $etlValues = New-Object System.Collections.ArrayList
      foreach ($record in $etlRecords) { [void]$etlValues.Add(@($record.No, (Convert-ToExcelText $record.PropId), $record.Policy, $record.Group)) }
      try { $wsEtl.Range("B2:B$requiredEtlLastRow").NumberFormat = "@" } catch {}
      Write-2DRange $wsEtl.Range("A2") $etlValues
    }

    $wb.Save()
    $saveWorkbook = $true
    return @{ IssueDataRows = $dataCount; IssueCheckRows = $checkCount; IssueM190Rows = $m190Ids.Count; IssueEtlRows = $etlCount }
  }
  finally {
    Close-WorkbookSafe $wb $saveWorkbook
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }
}

function Update-TwoColumnTable($wb, $sheetName, $tableName, $statusText, $ids) {
  $ws = $wb.Worksheets.Item($sheetName)
  $table = $ws.ListObjects.Item($tableName)
  $oldLastRow = $table.Range.Row + $table.Range.Rows.Count - 1
  if ($oldLastRow -ge 2) {
    [void]$ws.Range("A2:B$oldLastRow").ClearContents()
  }

  $dataCount = $ids.Count
  $requiredLastRow = [Math]::Max($dataCount, 1) + 1
  if ($oldLastRow -lt $requiredLastRow) {
    try {
      [void]$table.Resize($ws.Range("A1:B$requiredLastRow"))
      $oldLastRow = $table.Range.Row + $table.Range.Rows.Count - 1
    }
    catch {
      $capacity = [Math]::Max($oldLastRow - 1, 0)
      throw "$sheetName table $tableName has capacity $capacity rows, but needs $dataCount rows. Expand the table before running."
    }
  }

  if ($dataCount -gt 0) {
    $values = @()
    foreach ($id in $ids) {
      $values += ,@($statusText, (Convert-ToExcelText $id))
    }
    try { $ws.Range("B2:B$oldLastRow").NumberFormat = "@" } catch {}
    Write-2DRange $ws.Range("A2") $values
  }

  if ($oldLastRow -ge 2) {
    try { $ws.Range("B2:B$oldLastRow").NumberFormat = "@" } catch {}
  }
}

function Get-LastNonEmptyRow($ws, $firstCol, $lastCol, $minRow) {
  try {
    $lastRow = $ws.Cells.Item($ws.Rows.Count, $firstCol).End($xlUp).Row
    if ($lastRow -lt $minRow) { return $minRow }
    return $lastRow
  }
  catch {
    return $minRow
  }
}

function Get-UsedLastRow($ws, $minRow) {
  try {
    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    if ($lastRow -lt $minRow) { return $minRow }
    return $lastRow
  }
  catch {
    return $minRow
  }
}

function Is-PvBlankValue($value) {
  if ($null -eq $value) { return $true }
  $text = (Normalize-Header $value)
  return ($text -eq "" -or $text -eq "(blank)")
}

function Apply-PvProposalIdFilter($wb) {
  $result = [ordered]@{
    Applied = $false
    BlankItemsHidden = 0
    PivotTablesChecked = 0
    Error = ""
  }

  try {
    $wsPv = $wb.Worksheets.Item("PV ")
  }
  catch {
    $result.Error = "Cannot find PV sheet"
    return $result
  }

  try {
    foreach ($pt in $wsPv.PivotTables()) {
      $result.PivotTablesChecked++

      $field = $null
      try { $field = $pt.PivotFields("ProposalID") } catch { continue }
      if ($null -eq $field) { continue }

      try { $pt.ManualUpdate = $true } catch {}
      try { [void]$field.ClearAllFilters() } catch {}

      $hiddenInThisPivot = $false
      try {
        foreach ($item in $field.PivotItems()) {
          $itemName = ""
          try { $itemName = Normalize-Header $item.Name } catch {}

          if (Is-PvBlankValue $itemName) {
            try {
              $item.Visible = $false
              $hiddenInThisPivot = $true
              $result.BlankItemsHidden++
            }
            catch {}
          }
        }
      }
      finally {
        try { $pt.ManualUpdate = $false } catch {}
      }

      if ($hiddenInThisPivot) {
        $result.Applied = $true
        try { [void]$pt.RefreshTable() } catch {}
      }
    }
  }
  catch {
    $result.Error = $_.Exception.Message
  }

  return $result
}

function Format-ReportManual($wb, $wsFinal, $sourceRows) {
  $result = @{
    ReportPivotsRefreshed = 0
    ReportStatusBlocksVisible = 0
    ReportStatusBlocksHidden = 0
    ReportStatusFilterErrors = ""
  }

  try {
    $wsReport = $wb.Worksheets.Item("Report")
    $refreshedCaches = @{}
    foreach ($pt in $wsReport.PivotTables()) {
      $cacheIndex = [string]$pt.CacheIndex
      if ($refreshedCaches.ContainsKey($cacheIndex)) { continue }

      $cache = $pt.PivotCache()
      try { $cache.MissingItemsLimit = 0 } catch {}
      [void]$cache.Refresh()
      $refreshedCaches[$cacheIndex] = $true
    }

    $statusCounts = @{}
    if ($sourceRows -ge 2) {
      $statusData = $wsFinal.Range($wsFinal.Cells.Item(2, 6), $wsFinal.Cells.Item($sourceRows, 6)).Value2
      $statusRows = $sourceRows - 1
      for ($i = 1; $i -le $statusRows; $i++) {
        $status = Normalize-Header (Get-MatrixValue $statusData $i 1)
        if ($status -ne "") {
          if (-not $statusCounts.ContainsKey($status)) { $statusCounts[$status] = 0 }
          $statusCounts[$status]++
        }
      }
    }

    $blocks = @(
      @{ Pivot = "PivotTable14"; Status = "รอ Issue" },
      @{ Pivot = "PivotTable5"; Status = "ติดปัญหาไม่เข้าในเมนู E" },
      @{ Pivot = "PivotTable3"; Status = "ข้อมูลไม่สมบูรณ์" },
      @{ Pivot = "PivotTable1"; Status = "Blacklist" }
    )

    $blockRanges = New-Object System.Collections.ArrayList
    foreach ($block in $blocks) {
      try {
        $pt = $wsReport.PivotTables().Item($block.Pivot)
        $hasData = ($statusCounts.ContainsKey($block.Status) -and $statusCounts[$block.Status] -gt 0)
        try {
          $pf = $pt.PivotFields("สถานะไม่ issue")
          try { $pf.ClearAllFilters() } catch {}
          if ($hasData) {
            $filterApplied = $false
            try {
              $pf.EnableMultiplePageItems = $false
              $pf.CurrentPage = $block.Status
              $filterApplied = $true
            }
            catch {}

            if (-not $filterApplied) {
              try {
                $pf.EnableMultiplePageItems = $true
                $targetItem = $pf.PivotItems($block.Status)
                $targetItem.Visible = $true
                foreach ($item in $pf.PivotItems()) {
                  if ($item.Name -ne $block.Status) {
                    try { $item.Visible = $false } catch {}
                  }
                }
                $filterApplied = $true
              }
              catch {}
            }

            if (-not $filterApplied) {
              throw "Cannot apply status filter: $($block.Status)"
            }
          }
        }
        catch {
          if ($hasData) {
            $result.ReportStatusFilterErrors = ($result.ReportStatusFilterErrors + " " + $block.Pivot + ": " + $_.Exception.Message).Trim()
          }
        }

        [void]$pt.RefreshTable()
        $result.ReportPivotsRefreshed++
        $top = [int]$pt.TableRange2.Row
        $bottom = [int]($pt.TableRange2.Row + $pt.TableRange2.Rows.Count - 1)
        [void]$blockRanges.Add(@{
          Pivot = $block.Pivot
          Status = $block.Status
          HasData = $hasData
          Top = $top
          Bottom = $bottom
        })
      }
      catch {
        $result.ReportStatusFilterErrors = ($result.ReportStatusFilterErrors + " " + $block.Pivot + ": " + $_.Exception.Message).Trim()
      }
    }

    $wsReport.Rows.Hidden = $false
    $wsReport.Rows("11:300").Hidden = $true

    $visibleBlocks = @($blockRanges | Where-Object { $_.HasData } | Sort-Object { [int]$_.Top })
    $previousBottom = $null
    foreach ($block in $visibleBlocks) {
      $titleRow = [int]$block.Top + 1
      $valuesRow = [int]$block.Top + 2
      $bottom = [int]$block.Bottom
      if ($titleRow -le $bottom) {
        $wsReport.Rows("$titleRow`:$bottom").Hidden = $false
        if ($valuesRow -le $bottom) { $wsReport.Rows("$valuesRow`:$valuesRow").Hidden = $true }
        if ($null -ne $previousBottom -and ($titleRow -gt 12)) {
          $blankRow = $titleRow - 2
          $wsReport.Rows("$blankRow`:$blankRow").Hidden = $false
          $wsReport.Rows("$blankRow`:$blankRow").RowHeight = 20
        }
        $previousBottom = $bottom
        $result.ReportStatusBlocksVisible++
      }
    }

    $result.ReportStatusBlocksHidden = @($blockRanges | Where-Object { -not $_.HasData }).Count

    try {
      $wsReport.Cells.Font.Name = "Tahoma"
      $wsReport.Cells.Font.Size = 10
      $wsReport.Cells.VerticalAlignment = -4108
      $wsReport.Columns.Item(1).ColumnWidth = 22
      $wsReport.Columns.Item(2).ColumnWidth = 32
      $wsReport.Columns.Item(3).ColumnWidth = 18
      $wsReport.Columns.Item(4).ColumnWidth = 18

      $wsReport.Rows("4:4").Hidden = $false
      $wsReport.Rows("4:4").RowHeight = 20
      $wsReport.Rows("5:5").Hidden = $true

      foreach ($rowNumber in @(1, 6)) {
        $rowRange = $wsReport.Rows("$rowNumber`:$rowNumber")
        $rowRange.RowHeight = 30
        $rowRange.Font.Name = "Tahoma"
        $rowRange.Font.Size = 20
        $rowRange.Font.Bold = $true
        $rowRange.HorizontalAlignment = -4108
        $rowRange.VerticalAlignment = -4108
      }

      try {
        $wsReport.Rows("2:3").Font.Size = 10
        $wsReport.Rows("7:7").RowHeight = 20
        $wsReport.Rows("7:7").Font.Size = 12
        $wsReport.Rows("7:7").Font.Bold = $true
        $wsReport.Rows("8:9").Font.Size = 10
        $wsReport.Rows("8:9").Font.Bold = $true
      } catch {}

      foreach ($block in $blockRanges) {
        $titleRow = [int]$block.Top + 1
        $valuesRow = [int]$block.Top + 2
        $headerRow = [int]$block.Top + 3
        $bottom = [int]$block.Bottom

        $titleRange = $wsReport.Rows("$titleRow`:$titleRow")
        $titleRange.RowHeight = 30
        $titleRange.Font.Name = "Tahoma"
        $titleRange.Font.Size = 20
        $titleRange.Font.Bold = $true
        $titleRange.HorizontalAlignment = -4108
        $titleRange.VerticalAlignment = -4108

        if ($valuesRow -le $bottom) {
          $wsReport.Rows("$valuesRow`:$valuesRow").Font.Name = "Tahoma"
          $wsReport.Rows("$valuesRow`:$valuesRow").Font.Size = 9
        }

        if ($headerRow -le $bottom) {
          $headerRange = $wsReport.Rows("$headerRow`:$headerRow")
          $headerRange.RowHeight = 20
          $headerRange.Font.Name = "Tahoma"
          $headerRange.Font.Size = 12
          $headerRange.Font.Bold = $true
          $headerRange.HorizontalAlignment = -4108
          $headerRange.VerticalAlignment = -4108
        }

        if ($bottom -ge ($headerRow + 1)) {
          $dataRange = $wsReport.Rows("$($headerRow + 1)`:$bottom")
          $dataRange.Font.Name = "Tahoma"
          $dataRange.Font.Size = 10
          $dataRange.VerticalAlignment = -4108
        }

        $grandRange = $wsReport.Rows("$bottom`:$bottom")
        $grandRange.Font.Bold = $true
      }

      $wsReport.Rows("11:300").Hidden = $true
      $previousBottom = $null
      foreach ($block in $visibleBlocks) {
        $titleRow = [int]$block.Top + 1
        $valuesRow = [int]$block.Top + 2
        $bottom = [int]$block.Bottom
        if ($titleRow -le $bottom) {
          $wsReport.Rows("$titleRow`:$bottom").Hidden = $false
          if ($valuesRow -le $bottom) { $wsReport.Rows("$valuesRow`:$valuesRow").Hidden = $true }
          if ($null -ne $previousBottom -and ($titleRow -gt 12)) {
            $blankRow = $titleRow - 2
            $wsReport.Rows("$blankRow`:$blankRow").Hidden = $false
            $wsReport.Rows("$blankRow`:$blankRow").RowHeight = 20
          }
          $previousBottom = $bottom
        }
      }
    } catch {}
  }
  catch {
    $result.ReportStatusFilterErrors = ($result.ReportStatusFilterErrors + " " + $_.Exception.Message).Trim()
  }

  return $result
}

function Refresh-And-CopyPv($excel, $wb) {
  try { $excel.Calculation = $xlCalculationAutomatic } catch {}
  try { [void]$excel.CalculateFull() } catch {}

  $refreshedCaches = @{}
  foreach ($ws in $wb.Worksheets) {
    try {
      foreach ($pt in $ws.PivotTables()) {
        $cacheIndex = [string]$pt.CacheIndex
        if ($refreshedCaches.ContainsKey($cacheIndex)) { continue }
        $cache = $pt.PivotCache()
        try { $cache.MissingItemsLimit = 0 } catch {}
        [void]$cache.Refresh()
        $refreshedCaches[$cacheIndex] = $true
      }
    } catch {}
  }
  try { [void]$excel.CalculateFull() } catch {}

  $filterInfo = Apply-PvProposalIdFilter $wb
  $wsPv = $wb.Worksheets.Item("PV ")
  $headerRow = $null
  for ($row = 1; $row -le 30; $row++) {
    if ((Normalize-Header $wsPv.Cells.Item($row, 1).Value2) -eq "Date" -and
        (Normalize-Header $wsPv.Cells.Item($row, 2).Value2) -eq "Policy") {
      $headerRow = $row
      break
    }
  }
  if ($null -eq $headerRow) { throw "Cannot find PV header row" }

  $lastRow = Get-LastNonEmptyRow $wsPv 1 9 $headerRow
  if ($lastRow -lt $headerRow) { $lastRow = $headerRow }
  $sourceRange = $wsPv.Range($wsPv.Cells.Item($headerRow, 1), $wsPv.Cells.Item($lastRow, 9))
  $sourceValues = $sourceRange.Value2
  $sourceFormats = $sourceRange.NumberFormat
  $sourceCount = $lastRow - $headerRow + 1

  $include = New-Object System.Collections.ArrayList
  [void]$include.Add(1)
  for ($i = 2; $i -le $sourceCount; $i++) {
    $proposalIdText = Normalize-Header (Get-MatrixValue $sourceValues $i 5)
    if (-not (Is-PvBlankValue $proposalIdText)) { [void]$include.Add($i) }
  }
  $sourceRows = $include.Count

  $wsFinal = $wb.Worksheets.Item("PV final")
  $oldLastRow = Get-UsedLastRow $wsFinal 1
  try {
    $table = $wsFinal.ListObjects.Item("Table15")
    [void]$table.Resize($wsFinal.Range("A1:I$sourceRows"))
  } catch {}
  if ($oldLastRow -gt $sourceRows) { [void]$wsFinal.Rows("$($sourceRows + 1):$oldLastRow").Delete() }

  $outValues = New-Object "object[,]" $sourceRows, 9
  $outFormats = New-Object "object[,]" $sourceRows, 9
  for ($r = 0; $r -lt $sourceRows; $r++) {
    $srcRow = [int]$include[$r]
    for ($c = 1; $c -le 9; $c++) {
      $outValues[$r, ($c - 1)] = Get-MatrixValue $sourceValues $srcRow $c
      $outFormats[$r, ($c - 1)] = Get-MatrixValue $sourceFormats $srcRow $c
    }
  }
  $dest = $wsFinal.Range($wsFinal.Cells.Item(1, 1), $wsFinal.Cells.Item($sourceRows, 9))
  $dest.Value2 = $outValues
  try { $dest.NumberFormat = $outFormats } catch {}

  for ($col = 1; $col -le 9; $col++) {
    $wsFinal.Columns.Item($col).ColumnWidth = $wsPv.Columns.Item($col).ColumnWidth
  }

  $reportInfo = Format-ReportManual $wb $wsFinal $sourceRows
  try { [void]$excel.CalculateFull() } catch {}

  return @{
    HeaderRow = $headerRow
    LastRow = $lastRow
    RowsCopied = $sourceRows
    DataRowsCopied = [Math]::Max(0, $sourceRows - 1)
    PvBlankRowsSkipped = (($lastRow - $headerRow + 1) - $sourceRows)
    PvProposalFilterApplied = $filterInfo["Applied"]
    PvProposalBlankItemsHidden = $filterInfo["BlankItemsHidden"]
    PvProposalFilterError = $filterInfo["Error"]
    ReportPivotsRefreshed = $reportInfo["ReportPivotsRefreshed"]
    ReportStatusBlocksVisible = $reportInfo["ReportStatusBlocksVisible"]
    ReportStatusBlocksHidden = $reportInfo["ReportStatusBlocksHidden"]
    ReportStatusFilterErrors = $reportInfo["ReportStatusFilterErrors"]
  }
}

if ($ApplyPvFilterOnly) {
  if ($WorkbookPath -eq "") { throw "WorkbookPath is required when using ApplyPvFilterOnly." }
  $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
  $reportProtection = Suspend-XlsxSheetProtection $resolvedWorkbookPath "Report"

  Write-ProgressLine 20 "กำลังเปิด Microsoft Excel"
  $excel = New-Object -ComObject Excel.Application
  Save-ExcelProcessId $excel $OutputDir
  Set-ExcelTurboMode $excel

  try {
    $wb = Open-WorkbookSafe $excel $resolvedWorkbookPath $false
    $saveWorkbook = $false
    try {
      $pvInfo = Refresh-And-CopyPv $excel $wb
      $wb.Save()
      $saveWorkbook = $true
    }
    finally {
      Close-WorkbookSafe $wb $saveWorkbook
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    [void](Repair-XlsxTableMetadata $resolvedWorkbookPath)

    [pscustomobject]@{
      WorkbookPath = $resolvedWorkbookPath
      PvRowsCopiedToPvFinal = $pvInfo.RowsCopied
      PvDataRowsCopiedToPvFinal = $pvInfo.DataRowsCopied
      PvBlankRowsSkipped = $pvInfo.PvBlankRowsSkipped
      PvProposalFilterApplied = $pvInfo.PvProposalFilterApplied
      PvProposalBlankItemsHidden = $pvInfo.PvProposalBlankItemsHidden
      PvProposalFilterError = $pvInfo.PvProposalFilterError
      ReportPivotsRefreshed = $pvInfo.ReportPivotsRefreshed
      ReportStatusBlocksVisible = $pvInfo.ReportStatusBlocksVisible
      ReportStatusBlocksHidden = $pvInfo.ReportStatusBlocksHidden
      ReportStatusFilterErrors = $pvInfo.ReportStatusFilterErrors
    } | ConvertTo-Json -Depth 4
  }
  finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Restore-XlsxSheetProtection $resolvedWorkbookPath $reportProtection
  }

  exit 0
}

Write-ProgressLine 18 "กำลังตรวจไฟล์ที่รับเข้ามา"
$masterFile = Find-OneFile "*เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก*.xlsx" $true
$issueFile = Find-OneFile "*เช็คสถานะ ISSUE*.xlsx" $true
$reportFile = Find-OneFile "*รายงานงานประกันแรงงานต่างด้าว*.xlsx" $true
$m190File = Find-OneFile "M190027_PRD008_Premium_by_Policy*.xlsx" $true
$smFile = Find-OneFile "*ข้อมูลไม่สมบูรณ์*.xlsx" $false
$blFile = Find-OneFile "Blacklist*.xls*" $false
$etlFile = Find-OneFile "ETL*.txt" $false

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$workflowSucceeded = $false

Write-ProgressLine 20 "กำลังสร้างสำเนา Master ผลลัพธ์"
$outputPath = Join-Path $OutputDir $masterFile.Name
Copy-Item -LiteralPath $masterFile.FullName -Destination $outputPath -Force
[void](Repair-XlsxTableMetadata $outputPath)
[void](Ensure-XlsxNonIssueStatusFormula $outputPath)
$reportProtection = Suspend-XlsxSheetProtection $outputPath "Report"
$outputFile = Get-Item -LiteralPath $outputPath

Write-ProgressLine 22 "กำลังเปิด Microsoft Excel แบบ Turbo"
$excel = New-Object -ComObject Excel.Application
Save-ExcelProcessId $excel $OutputDir
Set-ExcelTurboMode $excel

try {
  $autoStartDate = ($StartDate -eq [datetime]::MinValue)
  $autoEndDate = ($EndDate -eq [datetime]::MinValue)
  $historicalMode = $false
  $dateMode = "Normal"

  Write-ProgressLine 25 "กำลังอ่านช่วงวันที่จาก Master"
  if ($autoStartDate) {
    $StartDate = Get-MasterStartDate $excel $outputFile
  }

  Write-ProgressLine 27 "กำลังตรวจช่วงวันที่และจำนวนแถวจาก Daily Report"
  $reportDateRange = Get-ReportDateRange $excel $reportFile
  $reportEarliestDate = $reportDateRange.MinDate
  $reportLatestDate = $reportDateRange.MaxDate

  if ($autoEndDate) {
    if ($reportLatestDate.Date -lt $TodayDate.Date) {
      $EndDate = $reportLatestDate.Date
    }
    else {
      $EndDate = $TodayDate.Date
    }
  }

  if ($EndDate.Date -lt $StartDate.Date) {
    if ($autoStartDate -and $autoEndDate) {
      # Historical Report Mode: ไฟล์ย้อนหลังเก่ากว่า Master ปัจจุบัน
      # ใช้วันล่าสุดของไฟล์เป็นรอบรายงาน เพื่อไม่ดึงประวัติทั้ง 5 แสนแถวเข้าผลลัพธ์
      $historicalMode = $true
      $dateMode = "HistoricalLatestDay"
      $StartDate = $reportLatestDate.Date
      $EndDate = $reportLatestDate.Date
    }
    else {
      throw "End date $($EndDate.ToString("yyyy-MM-dd")) is earlier than start date $($StartDate.ToString("yyyy-MM-dd"))"
    }
  }

  Write-ProgressLine 31 "ช่วงประมวลผล $($StartDate.ToString('yyyy-MM-dd')) ถึง $($EndDate.ToString('yyyy-MM-dd')) | Daily Report $($reportDateRange.WorksheetRows) แถว | Mode: $dateMode"
  Write-ProgressLine 34 "กำลังอ่าน Daily Report แบบ Large Data"
  $reportRows = Extract-ReportRows $excel $reportFile $StartDate $EndDate
  Write-ProgressLine 43 "กำลังอ่าน M190 แบบ Bulk"
  $m190Ids = Extract-M190PropIds $excel $m190File
  Write-ProgressLine 47 "กำลังตรวจ ETL"
  $etlRecords = Extract-EtlRecords $etlFile
  $etlIds = @($etlRecords | ForEach-Object { $_.PropId })
  Write-ProgressLine 49 "กำลังอ่านข้อมูลไม่สมบูรณ์"
  $smIds = Extract-CPropIds $excel $smFile
  Write-ProgressLine 51 "กำลังอ่าน Blacklist"
  $blIds = Extract-CPropIds $excel $blFile

  $checkSet = @{}
  foreach ($id in (@($m190Ids) + @($etlIds))) {
    if ($id -ne "") { $checkSet[$id] = $true }
  }

  Write-ProgressLine 54 "กำลังจับคู่ ProposalID และตัดรายการที่ออกแล้ว"
  $pendingRows = New-Object System.Collections.ArrayList
  $issuedCount = 0
  foreach ($row in $reportRows) {
    if ($checkSet.ContainsKey($row.ProposalId)) {
      $issuedCount++
    }
    else {
      [void]$pendingRows.Add($row)
    }
  }

  Write-ProgressLine 56 "กำลังอัปเดตเช็คสถานะ ISSUE"
  $issueInfo = Update-IssueWorkbook $excel $issueFile @($pendingRows.ToArray()) @($m190Ids) @($etlRecords)

  Write-ProgressLine 58 "กำลังเปิด Master เพื่ออัปเดต"
  $wbOut = Open-WorkbookSafe $excel $outputPath $false
  $saveWorkbook = $false
  try {
    Write-ProgressLine 61 "กำลังเขียน Pending ลง Master"
    Update-Table1 $wbOut @($pendingRows.ToArray())
    Write-ProgressLine 63 "กำลังอัปเดต SM และ Blacklist"
    Update-TwoColumnTable $wbOut "ข้อมูลไม่สมบูรณ์" "SM" "ข้อมูลไม่สมบูรณ์" @($smIds)
    Update-TwoColumnTable $wbOut "Black List" "BL" "Blacklist" @($blIds)
    Write-ProgressLine 65 "กำลังคำนวณสถานะ Menu E"
    $menuEInfo = Update-MenuEStatus $wbOut @($pendingRows.ToArray()) @($smIds) @($blIds) $EndDate
    Write-ProgressLine 67 "กำลัง Refresh Pivot และสร้าง Report"
    $pvInfo = Refresh-And-CopyPv $excel $wbOut
    Write-ProgressLine 71 "กำลังบันทึก Master และ ISSUE"
    $wbOut.Save()
    $saveWorkbook = $true
  }
  finally {
    Close-WorkbookSafe $wbOut $saveWorkbook
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbOut)
  }
  [void](Repair-XlsxTableMetadata $outputPath)

  $summary = [pscustomobject]@{
    OutputPath = $outputPath
    DateStart = $StartDate.ToString("yyyy-MM-dd")
    DateEnd = $EndDate.ToString("yyyy-MM-dd")
    ReportEarliestDate = $reportEarliestDate.ToString("yyyy-MM-dd")
    ReportLatestDate = $reportLatestDate.ToString("yyyy-MM-dd")
    ReportWorksheetRows = $reportDateRange.WorksheetRows
    ReportValidDateRows = $reportDateRange.ValidDateRows
    DateMode = $dateMode
    HistoricalMode = $historicalMode
    TodayDate = $TodayDate.ToString("yyyy-MM-dd")
    ReportRowsAfterDateStatusFilter = $reportRows.Count
    M190PropIdRows = $m190Ids.Count
    EtlPropIdRows = $etlIds.Count
    IssuedRowsRemoved = $issuedCount
    PendingRowsWrittenToData = $pendingRows.Count
    IssueDataRowsWritten = $issueInfo.IssueDataRows
    IssueCheckRowsWritten = $issueInfo.IssueCheckRows
    IssueM190RowsWritten = $issueInfo.IssueM190Rows
    IssueEtlRowsWritten = $issueInfo.IssueEtlRows
    MenuEStatusesWritten = $menuEInfo.MenuEWritten
    MenuEReferenceDate = $menuEInfo.MenuEReferenceDate
    SmPropIdsWritten = $smIds.Count
    BlPropIdsWritten = $blIds.Count
    PvRowsCopiedToPvFinal = $pvInfo.RowsCopied
    PvDataRowsCopiedToPvFinal = $pvInfo.DataRowsCopied
    PvBlankRowsSkipped = $pvInfo.PvBlankRowsSkipped
    PvProposalFilterApplied = $pvInfo.PvProposalFilterApplied
    PvProposalBlankItemsHidden = $pvInfo.PvProposalBlankItemsHidden
    PvProposalFilterError = $pvInfo.PvProposalFilterError
    ReportPivotsRefreshed = $pvInfo.ReportPivotsRefreshed
    ReportStatusBlocksVisible = $pvInfo.ReportStatusBlocksVisible
    ReportStatusBlocksHidden = $pvInfo.ReportStatusBlocksHidden
    ReportStatusFilterErrors = $pvInfo.ReportStatusFilterErrors
    EtlFile = if ($null -eq $etlFile) { "" } else { $etlFile.Name }
  }
  Write-ProgressLine 72 "Excel Workflow เสร็จสมบูรณ์"
  $workflowSucceeded = $true
  $summary | ConvertTo-Json -Depth 4
}
finally {
  $cleanupError = $null
  try {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Restore-XlsxSheetProtection $outputPath $reportProtection
  }
  catch {
    $workflowSucceeded = $false
    $cleanupError = $_
  }

  if ($null -ne $cleanupError) {
    throw $cleanupError
  }
}
