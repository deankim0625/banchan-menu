# 네이버 검색 API 키를 이 PC에만 저장합니다 (GitHub에는 올라가지 않음).
# 실행: 터미널에서  & .\tools\naver-key-save.ps1
$id = Read-Host "Client ID를 붙여 넣고 Enter"
$secure = Read-Host "Client Secret을 붙여 넣고 Enter (화면에 안 보여요)" -AsSecureString
$secret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
$path = Join-Path $env:USERPROFILE ".naver-api.json"
@{ id = $id.Trim(); secret = $secret.Trim() } | ConvertTo-Json | Set-Content -Encoding UTF8 $path
Write-Host "저장했어요: $path"
