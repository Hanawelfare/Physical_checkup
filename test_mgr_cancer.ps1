$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Let's test getRegistrationByEmpId for some MGR employees to see what is returned
$testIds = @("003049", "006125", "009892", "010268", "011382", "084390")

foreach ($id in $testIds) {
    $b = @{ action = "getRegistrationByEmpId"; args = @($id) } | ConvertTo-Json
    $r = Invoke-RestMethod -Uri $url -Method Post -Body $b -ContentType "application/json"
    Write-Host "ID $id Registration:"
    $r.data | ConvertTo-Json
    Write-Host "-------------------------------------------"
}
