$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Query an employee to see what fields are returned
$body = @{
    action = "getEmployeeData"
    args = @("053625")
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "Employee 053625 data:"
    $res.data | ConvertTo-Json
} catch {
    Write-Error $_
}
