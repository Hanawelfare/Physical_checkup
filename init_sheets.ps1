$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"
$body = @{
    action = "initializeSheets"
    args = @()
} | ConvertTo-Json

Write-Host "Sending initializeSheets request to backend..."
try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "Response received:"
    $res | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}
