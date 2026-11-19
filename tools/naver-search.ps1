# 네이버 검색 API (지역·블로그·카페·뉴스)
# 예: & .\tools\naver-search.ps1 -Query "마니반찬 선학" -Type local
param(
  [Parameter(Mandatory = $true)][string]$Query,
  [ValidateSet("local", "blog", "cafearticle", "news", "webkr")][string]$Type = "local",
  [int]$Display = 5,
  [ValidateSet("sim", "date", "random", "comment")][string]$Sort = "sim"
)
$path = Join-Path $env:USERPROFILE ".naver-api.json"
if (-not (Test-Path $path)) { Write-Error "키 파일이 없어요. 먼저 tools\naver-key-save.ps1 을 실행하세요."; exit 1 }
$k = Get-Content -Raw -Encoding UTF8 $path | ConvertFrom-Json
$max = if ($Type -eq "local") { [Math]::Min($Display, 5) } else { [Math]::Min($Display, 100) }
$url = "https://openapi.naver.com/v1/search/$Type.json?query=$([Uri]::EscapeDataString($Query))&display=$max&sort=$Sort"
$res = Invoke-RestMethod -Uri $url -Headers @{ "X-Naver-Client-Id" = $k.id; "X-Naver-Client-Secret" = $k.secret }
$strip = { param($s) ($s -replace '<[^>]+>', '') -replace '&amp;', '&' -replace '&quot;', '"' }
"총 $($res.total)건"
foreach ($it in $res.items) {
  if ($Type -eq "local") {
    [pscustomobject]@{ 이름 = & $strip $it.title; 분류 = $it.category; 주소 = $it.roadAddress; 전화 = $it.telephone; 링크 = $it.link }
  } else {
    [pscustomobject]@{ 제목 = & $strip $it.title; 날짜 = $it.postdate; 요약 = (& $strip $it.description); 링크 = $it.link }
  }
}
