# docs 폴더를 http://localhost:8765/ 로 보여 주는 아주 작은 웹서버 (미리보기용)
param([int]$Port = 8765)
$root = Join-Path $PSScriptRoot "..\docs" | Resolve-Path
$types = @{ ".html"="text/html; charset=utf-8"; ".js"="text/javascript"; ".css"="text/css"; ".json"="application/json"; ".png"="image/png"; ".jpg"="image/jpeg"; ".svg"="image/svg+xml" }
$l = New-Object System.Net.HttpListener
$l.Prefixes.Add("http://localhost:$Port/")
$l.Start()
Write-Host "serving $root on http://localhost:$Port/"
while ($l.IsListening) {
  $c = $l.GetContext()
  $p = [Uri]::UnescapeDataString($c.Request.Url.AbsolutePath).TrimStart("/")
  $f = Join-Path $root $p
  if (Test-Path $f -PathType Container) { $f = Join-Path $f "index.html" }
  $full = [IO.Path]::GetFullPath($f)
  if ($full.StartsWith($root.Path) -and (Test-Path $full -PathType Leaf)) {
    $b = [IO.File]::ReadAllBytes($full)
    $t = $types[[IO.Path]::GetExtension($full).ToLower()]; if (-not $t) { $t = "application/octet-stream" }
    $c.Response.ContentType = $t
    $c.Response.OutputStream.Write($b, 0, $b.Length)
  } else { $c.Response.StatusCode = 404 }
  $c.Response.Close()
}
