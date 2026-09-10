$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Check getAdminDashboardData to see what headers or employee IDs exist
$body = @{ action = "getAdminDashboardData"; args = @() } | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "Total employees returned by getAdminDashboardData:" $res.data.employees.Count
    Write-Host "Total registrations returned by getAdminDashboardData:" $res.data.registrations.Count
    
    # Check sample employee
    Write-Host "Sample employee:"
    $res.data.employees[0] | ConvertTo-Json
    
    # Check sample registration
    Write-Host "Sample registration:"
    $res.data.registrations[0] | ConvertTo-Json
} catch {
    Write-Error $_
}
