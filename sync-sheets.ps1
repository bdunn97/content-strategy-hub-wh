# Copies data from local Excel spreadsheets into site-data/*.json for the website.
# The spreadsheets stay where they are; this only reads them (Excel doesn't need to be open or even closed).
# Usage (from this folder):  powershell -ExecutionPolicy Bypass -File .\sync-sheets.ps1
#
# To add another spreadsheet, add an entry to $Sources below.
#   Source : full path to the .xlsx
#   Sheet  : worksheet (tab) name; leave '' for the first sheet
#   Out    : JSON file the page reads
#   Exclude: optional list of column headers to leave out of the JSON (never published)
# The header row is found automatically (the first row with 3+ filled cells). Rows above it
# (e.g. a title in A1) become "title". Every row below it becomes one record keyed by header text.

$Sources = @(
  @{
    Source = 'C:\Users\bdunn.NEWAGE\OneDrive - NewAge Products\Desktop\Wilder Home\Influencer\Wilder Home Potential Influencers.xlsx'
    Sheet  = ''
    Out    = 'site-data/potential-creators.json'
    # Columns kept off the site (the site and repo are public). Matched by header text.
    Exclude = @('Email')
  }
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ColIndex([string]$ref) {
  $letters = ($ref -replace '\d', '').ToUpper(); $n = 0
  foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
  $n - 1
}

function Read-Sheet([string]$path, [string]$sheetName) {
  # Read from a temp copy so an open/locked workbook is never touched
  $tmp = Join-Path $env:TEMP ('sync-' + [guid]::NewGuid().ToString('N') + '.xlsx')
  Copy-Item -LiteralPath $path $tmp
  $zip = [IO.Compression.ZipFile]::OpenRead($tmp)
  try {
    $readXml = {
      param($name)
      $e = $zip.GetEntry($name); if (-not $e) { return $null }
      $r = New-Object IO.StreamReader($e.Open()); [xml]$x = $r.ReadToEnd(); $r.Close(); $x
    }
    $shared = @()
    $sx = & $readXml 'xl/sharedStrings.xml'
    if ($sx) { foreach ($si in $sx.sst.si) { $shared += (($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '') } }

    $wb = & $readXml 'xl/workbook.xml'; $rels = & $readXml 'xl/_rels/workbook.xml.rels'
    $sheets = @($wb.workbook.sheets.sheet)
    $sheet = if ($sheetName) { $sheets | Where-Object { $_.name -eq $sheetName } | Select-Object -First 1 } else { $sheets[0] }
    if (-not $sheet) { throw "Sheet '$sheetName' not found in $path" }
    $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = (@($rels.Relationships.Relationship) | Where-Object { $_.Id -eq $rid }).Target -replace '^/?xl/', ''
    $sh = & $readXml ('xl/' + $target)

    # Hyperlinks attached to cells (used when a cell shows text but links somewhere)
    $links = @{}
    $sheetRelsName = 'xl/' + ($target -replace '([^/]+)$', '_rels/$1.rels')
    $srels = & $readXml $sheetRelsName
    foreach ($h in @($sh.worksheet.hyperlinks.hyperlink)) {
      if (-not $h) { continue }
      $hid = $h.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      if ($hid -and $srels) { $links[$h.ref] = (@($srels.Relationships.Relationship) | Where-Object { $_.Id -eq $hid }).Target }
    }

    $rows = @()
    foreach ($row in @($sh.worksheet.sheetData.row)) {
      if (-not $row) { continue }
      $cells = @{}
      foreach ($c in @($row.c)) {
        if (-not $c) { continue }
        $v = if ($c.t -eq 's') { $shared[[int]$c.v] } elseif ($c.t -eq 'inlineStr') { $c.is.InnerText } else { [string]$c.v }
        if ($null -eq $v) { $v = '' }
        $v = ([string]$v).Trim()
        if ($links.ContainsKey($c.r) -and $v -notmatch '^https?://') {
          $link = $links[$c.r] -replace '^mailto:', ''   # email links: keep just the address
          if (-not ($v -match '@' -and $link -match '@')) { $v = $link }
        }
        if ($v -ne '') { $cells[(ColIndex $c.r)] = $v }
      }
      $rows += , $cells
    }
    return $rows
  } finally { $zip.Dispose(); Remove-Item $tmp -ErrorAction SilentlyContinue }
}

foreach ($src in $Sources) {
  if (-not (Test-Path -LiteralPath $src.Source)) { throw "Spreadsheet not found: $($src.Source)" }
  $rows = Read-Sheet $src.Source $src.Sheet

  $hi = -1
  for ($i = 0; $i -lt $rows.Count; $i++) { if ($rows[$i].Count -ge 3) { $hi = $i; break } }
  if ($hi -lt 0) { throw "No header row found in $($src.Source)" }

  $title = (@($rows[0..([Math]::Max(0, $hi - 1))] | Where-Object { $_.Count -gt 0 } | ForEach-Object { $_.Values }) | Select-Object -First 1)
  $header = $rows[$hi]
  $exclude = @($src.Exclude | Where-Object { $_ } | ForEach-Object { $_.ToLower().Trim() })
  $cols = @($header.Keys | Sort-Object | Where-Object { $exclude -notcontains $header[$_].ToLower().Trim() })
  $columns = @($cols | ForEach-Object { $header[$_] })

  $records = @()
  for ($i = $hi + 1; $i -lt $rows.Count; $i++) {
    if ($rows[$i].Count -eq 0) { continue }
    $o = [ordered]@{}
    foreach ($ci in $cols) { $o[$header[$ci]] = $(if ($rows[$i].ContainsKey($ci)) { $rows[$i][$ci] } else { '' }) }
    $records += [pscustomobject]$o
  }

  $out = [ordered]@{
    title   = $(if ($hi -gt 0) { $title } else { '' })
    source  = Split-Path $src.Source -Leaf
    updated = (Get-Item -LiteralPath $src.Source).LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
    columns = $columns
    rows    = $records
  }
  $outPath = Join-Path $PSScriptRoot $src.Out
  New-Item -ItemType Directory -Force (Split-Path $outPath) | Out-Null
  $json = ConvertTo-Json -InputObject $out -Depth 5
  [IO.File]::WriteAllText($outPath, $json, (New-Object Text.UTF8Encoding $false))
  Write-Host ("{0} -> {1}: {2} rows, columns: {3}" -f $out.source, $src.Out, $records.Count, ($columns -join ', '))
}
