$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Test 1: Get Config and Slots
$body1 = @{
    action = "getConfigAndSlots"
    args = @()
} | ConvertTo-Json

Write-Host "Sending getConfigAndSlots request..."
try {
    $res1 = Invoke-RestMethod -Uri $url -Method Post -Body $body1 -ContentType "application/json"
    Write-Host "Success!"
    $res1 | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}

# Test 2: Search Employee '084843'
$body2 = @{
    action = "getEmployeeData"
    args = @("084843")
} | ConvertTo-Json

Write-Host "`nSending getEmployeeData for 084843 request..."
try {
    $res2 = Invoke-RestMethod -Uri $url -Method Post -Body $body2 -ContentType "application/json"
    Write-Host "Success!"
    $res2 | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}
