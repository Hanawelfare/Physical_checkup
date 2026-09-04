$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$body = @{
    action = "getAdminDashboardData"
    args = @()
} | ConvertTo-Json

try {
    $adminRes = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    $registrations = @($adminRes.data.registrations)
    $total = $registrations.Length
    Write-Host "Total registered employees in sheet: $total"
    
    $sampleIndices = @(0, 1, 2, 5, 10, 50, 100, 500, 1000, 1400, ($total - 1))
    $successCount = 0
    $failCount = 0
    
    foreach ($idx in $sampleIndices) {
        if ($idx -ge 0 -and $idx -lt $total) {
            $emp = $registrations[$idx]
            $empId = $emp.employeeId
            $lookupBody = @{
                action = "getRegistrationByEmpId"
                args = @($empId)
            } | ConvertTo-Json
            
            $regRes = Invoke-RestMethod -Uri $url -Method Post -Body $lookupBody -ContentType "application/json"
            if ($regRes.success -and $regRes.data -ne $null) {
                $d = $regRes.data
                Write-Host " [PASS] ID: $($d.employeeId) | Name: $($d.firstName) $($d.lastName) | Program: $($d.programGroup) | Date: $($d.dateString)"
                $successCount++
            } else {
                Write-Host " [FAIL] ID: $empId | Error: $($regRes.error)"
                $failCount++
            }
        }
    }
    
    $checkedTotal = $successCount + $failCount
    Write-Host "`nTest Summary: Passed $successCount / $checkedTotal checks."
} catch {
    Write-Error $_
}
