$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"
$body = '{"action":"prewarmCache","args":[]}'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "text/plain;charset=utf-8"
$sw.Stop()
Write-Host "Prewarm Time: $($sw.ElapsedMilliseconds) ms"
Write-Host ($res | ConvertTo-Json -Depth 5)
