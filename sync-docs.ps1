# Copies the ad concepts Word doc into site-data/ad-concepts.json for the Creative > Scripts page.
# The doc stays where it is; this only reads a temp copy (it can be open in Word).
# Usage (from this folder):  powershell -ExecutionPolicy Bypass -File .\sync-docs.ps1
#
# How the doc is read (keep this pattern when adding concepts):
#   "Concept N. <title>"            starts a new concept (concepts are shown in number order)
#   "Vibe: ... Good for: ... Length: ..."  the concept's summary (one per line)
#   "<Label>" or "<Label>: <text>"  starts a section, e.g. Suggested Visual Moments, Talking Points,
#                                   Text Overlay, Content Example(s)
#   bullet / numbered lines         items in the current section
#   "X" or an empty value           shown as TBD on the site

$Source = Join-Path $PSScriptRoot 'Wilder Home Ad Concepts.docx'
$OutFile = 'site-data/ad-concepts.json'

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (-not (Test-Path -LiteralPath $Source)) { throw "Doc not found: $Source" }

$tmp = Join-Path $env:TEMP ('sync-' + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $Source $tmp
$zip = [IO.Compression.ZipFile]::OpenRead($tmp)
try {
  $e = $zip.GetEntry('word/document.xml'); $r = New-Object IO.StreamReader($e.Open()); [xml]$doc = $r.ReadToEnd(); $r.Close()
} finally { $zip.Dispose(); Remove-Item $tmp -ErrorAction SilentlyContinue }

$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ns = New-Object Xml.XmlNamespaceManager($doc.NameTable); $ns.AddNamespace('w', $W)
$quotes = [char[]]@('"', [char]0x201C, [char]0x201D, ' ')

function Get-ParaText($p) {
  $sb = New-Object Text.StringBuilder
  foreach ($n in $p.SelectNodes('.//w:t|.//w:tab|.//w:br|.//w:cr', $ns)) {
    switch ($n.LocalName) { 't' { [void]$sb.Append($n.InnerText) } 'tab' { [void]$sb.Append(' ') } default { [void]$sb.Append("`n") } }
  }
  ($sb.ToString() -replace '[  ]+', ' ').Trim()
}

$title = ''; $concepts = New-Object Collections.ArrayList; $cur = $null; $sec = $null
foreach ($p in $doc.SelectNodes('//w:body/w:p', $ns)) {
  $text = Get-ParaText $p
  if (-not $text) { continue }
  $style = $p.SelectSingleNode('w:pPr/w:pStyle/@w:val', $ns)
  $isItem = $p.SelectSingleNode('w:pPr/w:numPr', $ns) -or ($style -and $style.Value -match 'List')

  if (-not $isItem -and $text -match '^Concept\s*(\d+)\s*[.:)-]?\s*(.*)$') {
    $cur = [ordered]@{ n = [int]$Matches[1]; title = $Matches[2].Trim($quotes); meta = [ordered]@{}; sections = New-Object Collections.ArrayList }
    [void]$concepts.Add($cur); $sec = $null; continue
  }
  if (-not $cur) { if (-not $title) { $title = $text }; continue }

  if ($isItem) {
    if (-not $sec) { $sec = [ordered]@{ label = ''; text = ''; items = New-Object Collections.ArrayList }; [void]$cur.sections.Add($sec) }
    [void]$sec.items.Add($text); continue
  }
  if ($text -match '^\s*Vibe\s*:') {
    # "Vibe: ... / Good for: ... / Length: ..." (line breaks or all on one line)
    $parts = [regex]::Split($text, '\n|(?=\b(?:Good for|Length|Vibe)\s*:)')
    foreach ($part in $parts) {
      if ($part -match '^\s*([^:\n]{1,30}):\s*(.*)$') { $cur.meta[$Matches[1].Trim()] = $Matches[2].Trim() }
    }
    continue
  }
  $label = $text; $val = ''
  if ($text -match '^([^:\n]{1,45}):\s*(.*)$') { $label = $Matches[1].Trim(); $val = $Matches[2].Trim() }
  $sec = [ordered]@{ label = $label; text = $val; items = New-Object Collections.ArrayList }
  [void]$cur.sections.Add($sec)
}

$out = [ordered]@{
  title    = $title
  source   = Split-Path $Source -Leaf
  updated  = (Get-Item -LiteralPath $Source).LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
  concepts = @($concepts)
}
$outPath = Join-Path $PSScriptRoot $OutFile
New-Item -ItemType Directory -Force (Split-Path $outPath) | Out-Null
[IO.File]::WriteAllText($outPath, (ConvertTo-Json -InputObject $out -Depth 8), (New-Object Text.UTF8Encoding $false))
Write-Host ("{0} -> {1}: {2} concepts ({3})" -f $out.source, $OutFile, $concepts.Count, (($concepts | ForEach-Object { "$($_.n). $($_.title)" }) -join '; '))
