$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Query getAdminDashboardData
$body = @{
    action = "getAdminDashboardData"
    args = @()
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "Admin Dashboard Summary:"
    Write-Host "Total registered:" $res.data.registrations.Count
    if ($res.data.registrations.Count -gt 0) {
        Write-Host "Sample registration records:"
        $res.data.registrations | Select-Object -First 5 | ConvertTo-Json
        
        $firstId = $res.data.registrations[0].employeeId
        Write-Host "Testing getRegistrationByEmpId with ID: '$firstId'"
        $body2 = @{
            action = "getRegistrationByEmpId"
            args = @($firstId)
        } | ConvertTo-Json
        $res2 = Invoke-RestMethod -Uri $url -Method Post -Body $body2 -ContentType "application/json"
        Write-Host "Result for $firstId :"
        $res2 | ConvertTo-Json -Depth 5
    }
} catch {
    Write-Error $_
}
