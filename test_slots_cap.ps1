$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Let's inspect the config slots and available remaining capacity
$bodyConfig = @{
    action = "getConfigAndSlots"
    args = @()
} | ConvertTo-Json

try {
    $resConfig = Invoke-RestMethod -Uri $url -Method Post -Body $bodyConfig -ContentType "application/json"
    Write-Host "Config Dates:"
    $resConfig.data.dates | ConvertTo-Json
    
    Write-Host "`nTime Slots count:" $resConfig.data.timeSlots.Count
    
    Write-Host "`nRegistration counts per slot:"
    $resConfig.data.registrationCounts | ConvertTo-Json
} catch {
    Write-Error $_
}
