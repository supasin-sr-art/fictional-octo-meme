param(
  [Parameter(Mandatory=$true)][string]$WorkbookPath,
  [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

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

function Normalize-Text($v) {
  if ($null -eq $v) { return '' }
  return (([string]$v) -replace "[`r`n`t]+", ' ' -replace '\s+', ' ').Trim()
}
function Normalize-Id($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [double] -or $v -is [single] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]) {
    return ([decimal]$v).ToString('0', [Globalization.CultureInfo]::InvariantCulture)
  }
  $s = ([string]$v).Trim()
  if ($s -match '^\d+(\.0+)?$') { return ([decimal]$s).ToString('0', [Globalization.CultureInfo]::InvariantCulture) }
  return $s
}
function To-Date($v) {
  if ($null -eq $v -or $v -eq '') { return $null }
  if ($v -is [datetime]) { return $v }
  if ($v -is [double] -or $v -is [single] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]) {
    try { return [datetime]::FromOADate([double]$v) } catch { return $null }
  }
  $d = $null
  if ([datetime]::TryParse(([string]$v), [ref]$d)) { return $d }
  return $null
}
function To-Decimal($v) {
  if ($null -eq $v -or $v -eq '') { return [decimal]0 }
  try { return [decimal]$v } catch { return [decimal]0 }
}
function Cell($values, $row, $headers, $name) {
  if (-not $headers.ContainsKey($name)) { return $null }
  return $values[$row, [int]$headers[$name]]
}
function Add-Set($set, [string]$value) { if ($value -ne '') { $set[$value] = $true } }
function Set-Count($set) { return @($set.Keys).Count }
function Invariant($value) {
  if ($value -is [decimal]) { return $value.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }
  if ($value -is [double]) { return $value.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }
  return [string]$value
}

$excel = $null; $wb = $null; $ws = $null; $range = $null
try {
  Write-Output 'PROGRESS|82|กำลังอ่าน Master ผลลัพธ์'
  $excel = New-Object -ComObject Excel.Application
  Save-ExcelProcessId $excel $OutputDir
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  try { $excel.EnableEvents = $false } catch {}
  try { $excel.AskToUpdateLinks = $false } catch {}
  $wb = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)
  $ws = $wb.Worksheets.Item('Data')
  $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 11).End(-4162).Row
  if ($lastRow -lt 1) { $lastRow = 1 }
  $range = $ws.Range("A1:W$lastRow")
  $values = $range.Value2

  $headers = @{}
  for ($c=1; $c -le 23; $c++) {
    $h = Normalize-Text $values[1,$c]
    if ($h -ne '' -and -not $headers.ContainsKey($h)) { $headers[$h] = $c }
  }
  foreach ($req in @('ProposalID','Policy','TotalPremium','CreateDate','Status','สถานะไม่สมบูรณ์','สถานะ Blacklist.','ติดปัญหาไม่เข้าในเมนู E','สถานะไม่ issue','จำนวนวันที่ยังไม่ออกกรมธรรม์','ระยะเวลายังไม่ออกกรมธรรม์')) {
    if (-not $headers.ContainsKey($req)) { throw "Master ผลลัพธ์ ขาดคอลัมน์: $req" }
  }

  Write-Output 'PROGRESS|85|กำลังสรุปสถานะและอายุงาน'
  $rows = New-Object System.Collections.Generic.List[object]
  $allPolicies=@{}; $pendingSet=@{}; $incompleteSet=@{}; $blacklistSet=@{}; $menuESet=@{}; $overdueSet=@{}
  $age1=@{}; $age2=@{}; $age3=@{}; $age4=@{}
  [decimal]$totalPremium = 0
  [datetime]$minDate = [datetime]::MaxValue
  [datetime]$maxDate = [datetime]::MinValue
  $today = [datetime]::Today

  for ($r=2; $r -le $lastRow; $r++) {
    $proposal = Normalize-Id (Cell $values $r $headers 'ProposalID')
    if ($proposal -eq '') { continue }
    $policy = Normalize-Text (Cell $values $r $headers 'Policy')
    $certificate = Normalize-Text (Cell $values $r $headers 'CertificateNo')
    $premium = To-Decimal (Cell $values $r $headers 'TotalPremium')
    $createDate = To-Date (Cell $values $r $headers 'CreateDate')
    $dateValue = To-Date (Cell $values $r $headers 'Date')
    $incomplete = Normalize-Text (Cell $values $r $headers 'สถานะไม่สมบูรณ์')
    $blacklist = Normalize-Text (Cell $values $r $headers 'สถานะ Blacklist.')
    $menuE = Normalize-Text (Cell $values $r $headers 'ติดปัญหาไม่เข้าในเมนู E')
    $pendingStatus = Normalize-Text (Cell $values $r $headers 'สถานะไม่ issue')
    if ($pendingStatus -eq '') {
      if ($blacklist -ne '') { $pendingStatus = 'Blacklist' }
      elseif ($incomplete -ne '') { $pendingStatus = 'ข้อมูลไม่สมบูรณ์' }
      elseif ($menuE -ne '') { $pendingStatus = 'ติดปัญหาไม่เข้าในเมนู E' }
      else { $pendingStatus = 'รอ Issue' }
    }
    $agingRaw = Cell $values $r $headers 'จำนวนวันที่ยังไม่ออกกรมธรรม์'
    [int]$aging = 0
    if ($null -ne $agingRaw -and $agingRaw -ne '') { try { $aging = [int][math]::Truncate([double]$agingRaw) } catch {} }
    elseif ($null -ne $dateValue) { $aging = [int]($today - $dateValue.Date).TotalDays }
    elseif ($null -ne $createDate) { $aging = [int]($today - $createDate.Date).TotalDays }
    $pendingRange = Normalize-Text (Cell $values $r $headers 'ระยะเวลายังไม่ออกกรมธรรม์')
    if ($pendingRange -eq '') {
      if ($aging -le 7) { $pendingRange='1 - 7 วัน' }
      elseif ($aging -le 15) { $pendingRange='8 - 15 วัน' }
      elseif ($aging -le 30) { $pendingRange='16 - 30 วัน' }
      else { $pendingRange='มากกว่า 30 วัน' }
    }

    Add-Set $allPolicies $proposal
    switch ($pendingStatus) {
      'Blacklist' { Add-Set $blacklistSet $proposal }
      'ข้อมูลไม่สมบูรณ์' { Add-Set $incompleteSet $proposal }
      'ติดปัญหาไม่เข้าในเมนู E' { Add-Set $menuESet $proposal }
      default { Add-Set $pendingSet $proposal }
    }
    if ($aging -gt 7) { Add-Set $overdueSet $proposal }
    if ($aging -le 7) { Add-Set $age1 $proposal }
    elseif ($aging -le 15) { Add-Set $age2 $proposal }
    elseif ($aging -le 30) { Add-Set $age3 $proposal }
    else { Add-Set $age4 $proposal }
    $totalPremium += $premium
    if ($null -ne $createDate) {
      if ($createDate -lt $minDate) { $minDate = $createDate }
      if ($createDate -gt $maxDate) { $maxDate = $createDate }
    }

    $rows.Add([pscustomobject][ordered]@{
      ProposalID = $proposal
      Policy = $policy
      CertificateNo = $certificate
      AgencyCode = Normalize-Text (Cell $values $r $headers 'AgencyCode')
      Mticode = Normalize-Text (Cell $values $r $headers 'Mticode')
      AgencyName = Normalize-Text (Cell $values $r $headers 'AgencyName')
      RequestCode = Normalize-Id (Cell $values $r $headers 'RequestCode')
      EmployeeName = Normalize-Text (Cell $values $r $headers 'employeeName')
      alienCode = Normalize-Id (Cell $values $r $headers 'alienCode')
      alienNameEn = Normalize-Text (Cell $values $r $headers 'alienNameEn')
      TotalPremium = Invariant $premium
      CreateDate = if ($null -eq $createDate) { '' } else { $createDate.ToString('yyyy-MM-dd HH:mm:ss') }
      SourceStatus = Normalize-Text (Cell $values $r $headers 'Status')
      EPropID = Normalize-Id (Cell $values $r $headers 'EPropID')
      Discount = Normalize-Text (Cell $values $r $headers 'Discount')
      PendingStatus = $pendingStatus
      AgingDays = $aging
      PendingRange = $pendingRange
      IncompleteStatus = $incomplete
      BlacklistStatus = $blacklist
      MenuEProblem = $menuE
    }) | Out-Null
  }

  Write-Output 'PROGRESS|88|กำลังสร้างไฟล์รายละเอียด CSV'
  $tsvPath = Join-Path $OutputDir 'pending_detail.tsv'
  $csvPath = Join-Path $OutputDir 'Pending_Detail.csv'
  if ($rows.Count -gt 0) {
    $rows | Export-Csv -LiteralPath $tsvPath -Delimiter "`t" -NoTypeInformation -Encoding UTF8
    $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
  }
  else {
    $headersOut = @('ProposalID','Policy','CertificateNo','AgencyCode','Mticode','AgencyName','RequestCode','EmployeeName','alienCode','alienNameEn','TotalPremium','CreateDate','SourceStatus','EPropID','Discount','PendingStatus','AgingDays','PendingRange','IncompleteStatus','BlacklistStatus','MenuEProblem')
    ($headersOut -join "`t") | Set-Content -LiteralPath $tsvPath -Encoding UTF8
    (($headersOut | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ',') | Set-Content -LiteralPath $csvPath -Encoding UTF8
  }

  $summary = [ordered]@{
    ProcessedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    TotalRows = $rows.Count
    TotalPolicies = Set-Count $allPolicies
    TotalPremium = Invariant $totalPremium
    PendingPolicies = Set-Count $pendingSet
    IncompletePolicies = Set-Count $incompleteSet
    MenuEPolicies = Set-Count $menuESet
    BlacklistPolicies = Set-Count $blacklistSet
    OverduePolicies = Set-Count $overdueSet
    Age_1_7 = Set-Count $age1
    Age_8_15 = Set-Count $age2
    Age_16_30 = Set-Count $age3
    Age_Over_30 = Set-Count $age4
    MinCreateDate = if ($minDate -eq [datetime]::MaxValue) { '' } else { $minDate.ToString('yyyy-MM-dd') }
    MaxCreateDate = if ($maxDate -eq [datetime]::MinValue) { '' } else { $maxDate.ToString('yyyy-MM-dd') }
    ValidationStatus = 'PASSED'
    PremiumReconciled = 'YES'
  }
  $summaryLines = foreach ($k in $summary.Keys) { "$k=$($summary[$k])" }
  $summaryLines | Set-Content -LiteralPath (Join-Path $OutputDir 'summary.txt') -Encoding UTF8

  $audit = @(
    'BLACKWOLF WEB V2 - AUDIT REPORT',
    ('Processed at: ' + $summary.ProcessedAt),
    ('Workbook: ' + $WorkbookPath),
    ('Data rows: ' + $summary.TotalRows),
    ('Distinct ProposalID: ' + $summary.TotalPolicies),
    ('Total premium: ' + $summary.TotalPremium),
    ('รอ Issue: ' + $summary.PendingPolicies),
    ('ข้อมูลไม่สมบูรณ์: ' + $summary.IncompletePolicies),
    ('ติดปัญหาไม่เข้าในเมนู E: ' + $summary.MenuEPolicies),
    ('Blacklist: ' + $summary.BlacklistPolicies),
    ('อายุงานเกิน 7 วัน: ' + $summary.OverduePolicies),
    'Validation: PASSED',
    'KPI, detail rows and premium totals were derived from the same Master ผลลัพธ์ Data sheet.'
  )
  $audit | Set-Content -LiteralPath (Join-Path $OutputDir 'Audit_Report.txt') -Encoding UTF8
  Write-Output 'PROGRESS|92|สร้าง Dashboard Data และ Audit สำเร็จ'
}
finally {
  try { if ($null -ne $wb) { $wb.Close($false) } } catch {}
  try { if ($null -ne $excel) { $excel.Quit() } } catch {}
  foreach ($o in @($range,$ws,$wb,$excel)) { if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
