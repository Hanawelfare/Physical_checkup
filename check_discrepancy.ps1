$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$body = @{
    action = "getAdminDashboardData"
    args = @()
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    $emps = $res.data.employees
    $regs = $res.data.registrations

    Write-Host "Total eligible employees in Name sheet (excluding ไม่มีสิทธิ์ / มี remark):" $emps.Count
    Write-Host "Total registration records in Registration sheet:" $regs.Count

    # Find unique registration IDs
    $regIds = $regs | ForEach-Object { $_.employeeId }
    $uniqueRegIds = $regIds | Select-Object -Unique
    Write-Host "Unique registered employee IDs:" $uniqueRegIds.Count

    # Check for duplicate registrations in Registration sheet
    $duplicateRegs = $regIds | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicateRegs) {
        Write-Host "`nFound duplicate registration IDs in Registration sheet:"
        $duplicateRegs | ForEach-Object { Write-Host " - ID: $($_.Name) count: $($_.Count)" }
    } else {
        Write-Host "`nNo duplicate registrations found."
    }

    # Find registered IDs that are NOT in eligible employees list
    $regEmpIdSet = @{}
    $emps | ForEach-Object { $regEmpIdSet[$_.employeeId] = $true }

    $regNotInEmps = $uniqueRegIds | Where-Object { -not $regEmpIdSet.ContainsKey($_) }
    if ($regNotInEmps) {
        Write-Host "`nFound registered IDs that are NOT in eligible employees list (e.g. have remark or marked ไม่มีสิทธิ์ or not in Name sheet):"
        $regNotInEmps | ForEach-Object { Write-Host " - Registered ID not in eligible: $_" }
    }

    # Find unregistered employees
    $registeredSet = @{}
    $uniqueRegIds | ForEach-Object { $registeredSet[$_] = $true }

    $unregistered = $emps | Where-Object { -not $registeredSet.ContainsKey($_.employeeId) }
    Write-Host "`nTotal Unregistered Employees in system:" $unregistered.Count
    Write-Host "Unregistered list:"
    $unregistered | ForEach-Object {
        Write-Host " - $($_.employeeId) : $($_.firstName) $($_.lastName) ($($_.department))"
    }

} catch {
    Write-Error $_
}
