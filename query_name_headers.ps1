$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"
$id = "084390"

$bodyEmp = @{
    action = "getEmployeeData"
    args = @($id)
} | ConvertTo-Json

try {
    $resEmp = Invoke-RestMethod -Uri $url -Method Post -Body $bodyEmp -ContentType "application/json"
    $resEmp | ConvertTo-Json -Depth 5
} catch {
    Write-Error $_
}
