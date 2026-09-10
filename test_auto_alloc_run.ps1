$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Call autoAllocateRemainingEmployees to see the actual response or error
$body = @{
    action = "autoAllocateRemainingEmployees"
    args = @()
} | ConvertTo-Json

try {
    Write-Host "Calling autoAllocateRemainingEmployees API..."
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "API Result:"
    $res | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}
