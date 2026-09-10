$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

$unregistered19 = @(
    "003049", "053625", "066525", "074011", "080095", "081473", "081586", "082821", 
    "086046", "086752", "087563", "087693", "089281", "089352", "089375", "089468", 
    "089704", "089799", "089890"
)

Write-Host "Inspecting defaultLocation, shift, and match rules for the 19 employees:"
foreach ($id in $unregistered19) {
    $b = @{ action = "getEmployeeData"; args = @($id) } | ConvertTo-Json
    $r = Invoke-RestMethod -Uri $url -Method Post -Body $b -ContentType "application/json"
    if ($r.success) {
        $d = $r.data
        Write-Host "$id : $($d.firstName) $($d.lastName) | Dept: $($d.department) | Location: '$($d.defaultLocation)'"
    }
}
