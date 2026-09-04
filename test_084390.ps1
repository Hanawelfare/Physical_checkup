$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$id = "084390"
$body = @{
    action = "getRegistrationByEmpId"
    args = @($id)
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "getRegistrationByEmpId result:"
    $res | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}

$bodyEmp = @{
    action = "getEmployeeData"
    args = @($id)
} | ConvertTo-Json

try {
    $resEmp = Invoke-RestMethod -Uri $url -Method Post -Body $bodyEmp -ContentType "application/json"
    Write-Host "getEmployeeData result:"
    $resEmp | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}
