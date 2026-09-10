$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# 1. Test POST getEmployeeAndRegistration
$postBody = '{"action":"getEmployeeAndRegistration","args":["000001"]}'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$resPost = Invoke-RestMethod -Uri $url -Method Post -Body $postBody -ContentType "text/plain;charset=utf-8"
$sw.Stop()
Write-Host "POST getEmployeeAndRegistration took: $($sw.ElapsedMilliseconds) ms"
Write-Host "POST Result: $($resPost | ConvertTo-Json -Compress)"

# 2. Test GET getEmployeeAndRegistration
$sw2 = [System.Diagnostics.Stopwatch]::StartNew()
$resGet = Invoke-RestMethod -Uri "$url`?action=getEmployeeAndRegistration&args=%5B%22000001%22%5D" -Method Get
$sw2.Stop()
Write-Host "GET getEmployeeAndRegistration took: $($sw2.ElapsedMilliseconds) ms"
Write-Host "GET Result: $($resGet | ConvertTo-Json -Compress)"
