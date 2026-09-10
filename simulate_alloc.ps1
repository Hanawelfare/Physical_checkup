$url = "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec"

# Let's inspect the 18 employees in Name sheet:
# Why did autoAllocate not match them?
# Let's write a script that simulates autoAllocate step-by-step for the 18 unregistered employees!

$unregistered19 = @(
    "003049", "053625", "066525", "074011", "080095", "081473", "081586", "082821", 
    "086046", "086752", "087563", "087693", "089281", "089352", "089375", "089468", 
    "089704", "089799", "089890"
)

# Fetch config
$bodyConfig = @{ action = "getConfigAndSlots"; args = @() } | ConvertTo-Json
$resConfig = Invoke-RestMethod -Uri $url -Method Post -Body $bodyConfig -ContentType "application/json"
$dates = $resConfig.data.dates
$timeSlots = $resConfig.data.timeSlots
$counts = $resConfig.data.registrationCounts

Write-Host "Simulating autoAllocate step-by-step for each of the 19 employees:`n"

foreach ($id in $unregistered19) {
    $b = @{ action = "getEmployeeData"; args = @($id) } | ConvertTo-Json
    $r = Invoke-RestMethod -Uri $url -Method Post -Body $b -ContentType "application/json"
    $emp = $r.data
    
    $loc = $emp.defaultLocation
    $shift = "คร่อมกะ" # default
    
    $matched = $null
    
    foreach ($d in $dates) {
        if ($d.location -ne $loc) { continue }
        
        # Shift matching
        $shiftMatch = ($shift -eq "คร่อมกะ" -or $shift -eq "เช้าตลอด" -or $d.team -eq $shift)
        if (-not $shiftMatch) { continue }
        
        foreach ($t in $timeSlots) {
            # isFirstOrFifth check
            $is1or5 = ($d.dateString -like "*1 ตุลาคม*" -or $d.dateString -like "*5 ตุลาคม*")
            if ($is1or5) {
                # slots >= 14:00
                $matchHour = [regex]::Match($t.slotTime, "^(\d{2})[.:](\d{2})")
                if ($matchHour.Success) {
                    $h = [int]$matchHour.Groups[1].Value
                    $m = [int]$matchHour.Groups[2].Value
                    if (($h * 60 + $m) -ge 840) { continue }
                }
            }
            
            # Check key format!
            # In google_script.js line 1613, it used: location + "_" + dateObj.dateString + "_" + timeObj.slotTime
            # But counts dictionary has keys like: location + "|" + dateObj.dateString + "|" + timeObj.slotTime
            $keyPipe = "$($loc)|$($d.dateString)|$($t.slotTime)"
            $keyUnderscore = "$($loc)_$($d.dateString)_$($t.slotTime)"
            
            $curPipe = if ($counts.PSObject.Properties[$keyPipe]) { $counts.$keyPipe } else { 0 }
            
            if ($curPipe -lt $t.limit) {
                $matched = "$($d.dateString) ($($t.slotTime)) [Current: $curPipe / $($t.limit)]"
                break
            }
        }
        if ($matched) { break }
    }
    
    Write-Host "$id ($($emp.department)): Loc='$loc', Remark='$($emp.remark)', MatchedSlot: $matched"
}
