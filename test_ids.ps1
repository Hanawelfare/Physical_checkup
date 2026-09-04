$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$testIds = @("82977", "082977", "082977 ", " 082977", "084843")
foreach ($id in $testIds) {
    $body = @{
        action = "getRegistrationByEmpId"
        args = @($id)
    } | ConvertTo-Json
    $res = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
    Write-Host "Search for '$id':" ($res.data -ne $null)
    if ($res.data -ne $null) {
        Write-Host "  Name: " $res.data.firstName $res.data.lastName "Date: " $res.data.dateString
    } else {
        Write-Host "  Result is NULL"
    }
}
