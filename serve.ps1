# Serves the Influencer Content Hub at http://localhost:4324
# Usage (from this folder):  powershell -ExecutionPolicy Bypass -File .\serve.ps1
# Stop it with Ctrl+C or by closing the window.
param([int]$Port = 4324)

$Root = $PSScriptRoot
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Influencer Content Hub running at http://localhost:$Port  (Ctrl+C to stop)"

$types = @{
  '.html' = 'text/html; charset=utf-8'; '.js' = 'application/javascript; charset=utf-8'
  '.css' = 'text/css'; '.json' = 'application/json'; '.png' = 'image/png'; '.jpg' = 'image/jpeg'
  '.svg' = 'image/svg+xml'; '.mp4' = 'video/mp4'; '.m4v' = 'video/mp4'; '.mov' = 'video/quicktime'; '.webm' = 'video/webm'
}

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $res = $ctx.Response
  try {
    $rel = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
    if (-not $rel) { $rel = 'index.html' }
    $full = [IO.Path]::GetFullPath((Join-Path $Root $rel))
    if (-not $full.StartsWith($Root) -or -not (Test-Path -LiteralPath $full -PathType Leaf)) {
      $res.StatusCode = 404; continue
    }
    $ext = [IO.Path]::GetExtension($full).ToLower()
    $res.ContentType = $(if ($types[$ext]) { $types[$ext] } else { 'application/octet-stream' })
    $res.AddHeader('Accept-Ranges', 'bytes')
    $fs = [IO.File]::OpenRead($full)
    try {
      $len = $fs.Length; $start = 0; $end = $len - 1
      if ($ctx.Request.Headers['Range'] -match 'bytes=(\d*)-(\d*)') {
        if ($matches[1]) { $start = [long]$matches[1] }
        if ($matches[2]) { $end = [Math]::Min([long]$matches[2], $len - 1) }
        $res.StatusCode = 206
        $res.AddHeader('Content-Range', "bytes $start-$end/$len")
      }
      $res.ContentLength64 = $end - $start + 1
      $fs.Seek($start, 'Begin') | Out-Null
      $buf = New-Object byte[] 65536; $left = $end - $start + 1
      while ($left -gt 0) {
        $n = $fs.Read($buf, 0, [int][Math]::Min($buf.Length, $left))
        if ($n -le 0) { break }
        $res.OutputStream.Write($buf, 0, $n); $left -= $n
      }
    } finally { $fs.Close() }
  } catch {
    # client closed the connection mid-stream (normal when seeking video)
  } finally {
    try { $res.Close() } catch {}
  }
}
