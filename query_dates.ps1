$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"
$body = @{
    action = "getConfigAndSlots"
    args = @()
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "DATES CONFIG:" -ForegroundColor Cyan
    $res.data.dates | ConvertTo-Json
    
    Write-Host "`nTIME SLOTS CONFIG:" -ForegroundColor Cyan
    $res.data.timeSlots | ConvertTo-Json
} catch {
    Write-Error $_
}
