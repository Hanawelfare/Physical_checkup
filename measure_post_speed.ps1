$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$ids = @("003049", "084390", "012055", "004148", "082977")

Write-Host "Measuring API response times for getEmployeeAndRegistration via POST:"
foreach ($id in $ids) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $body = @{ action = "getEmployeeAndRegistration"; args = @($id) } | ConvertTo-Json
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    $sw.Stop()
    Write-Host "ID $id : $($sw.ElapsedMilliseconds) ms | Success: $($res.success)"
}
