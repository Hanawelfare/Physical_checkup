$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$body = @{ action = "getAdminDashboardData"; args = @() } | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    $emps = $res.data.employees
    $regs = $res.data.registrations

    $regSet = @{}
    $regs | ForEach-Object { $regSet[$_.employeeId] = $true }

    $unregistered = $emps | Where-Object { -not $regSet.ContainsKey($_.employeeId) }

    Write-Host "Total Eligible Employees:" $emps.Count
    Write-Host "Total Registered Records:" $regs.Count
    Write-Host "Remaining Unregistered Employees:" $unregistered.Count

    if ($unregistered.Count -gt 0) {
        Write-Host "Remaining Unregistered IDs:"
        $unregistered | ForEach-Object { Write-Host " - $($_.employeeId)" }
    } else {
        Write-Host "`n>>> AMAZING! ALL EMPLOYEES ARE NOW 100% REGISTERED! 0 UNREGISTERED REMAINING! <<<"
    }
} catch {
    Write-Error $_
}
