# =========================
# CONFIG
# =========================
$policyName1 = "CIS Intune Baseline (L2)  4.0.0"
$policyName2 = "Windows 11 Baseline"
$outputExcel = "C:\Temp\IntunePolicyDiff_Live.xlsx"

# =========================
# HELPER: SHORT ID
# =========================
function Get-ShortSettingId($fullId) {
    #if (-not $fullId) { return $null }

    #$parts = $fullId -split "_"

    #if ($parts.Count -ge 2) {
    #    Write-Host "   → Kürze '$fullId' zu '$($parts[-2..-1] -join "_")'" -ForegroundColor Gray
    #    return ($parts[-2..-1] -join "_")
    #}

    return $fullId
}

# =========================
# CONNECT GRAPH
# =========================
Write-Host "🔐 Verbinde zu Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" #-UseDeviceCode

$apiBase = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"

# =========================
# HELPER: GET POLICY BY NAME
# =========================
function Get-PolicyByName($name) {
    try{
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies`?`$filter=name eq '$name'"
        write-host $uri
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        return $result.value
        }
    catch {
        Write-Host "❌ Fehler beim Abrufen der Policies: $_" -ForegroundColor Red
        return $null
        }
}

# =========================
# LOAD POLICIES
# =========================
Write-Host "🔎 Suche Policies..." -ForegroundColor Cyan

$policy1 = Get-PolicyByName $policyName1
$policy2 = Get-PolicyByName $policyName2

if (-not $policy1 -or -not $policy2) {
    Write-Host "❌ Eine oder beide Policies wurden nicht gefunden!" -ForegroundColor Red
    return
}

Write-Host "✔ Policy 1 gefunden: $($policy1.name)" -ForegroundColor Green
Write-Host "✔ Policy 2 gefunden: $($policy2.name)" -ForegroundColor Green

# =========================
# GET FULL POLICIES
# =========================
Write-Host "📥 Lade vollständige Policies inkl. Settings..." -ForegroundColor Cyan

$policy1Full = Invoke-MgGraphRequest -Method GET `
-Uri "$apiBase/$($policy1.id)?`$expand=settings"

$policy2Full = Invoke-MgGraphRequest -Method GET `
-Uri "$apiBase/$($policy2.id)?`$expand=settings"

# =========================
# NORMALIZE SETTINGS
# =========================
function Get-NormalizedSettings($policy) {
    return $policy.settings | ForEach-Object {
        $inst = $_.settingInstance
        $fullId = $inst.settingDefinitionId

        [PSCustomObject]@{
            Id      = $fullId
            ShortId = Get-ShortSettingId $fullId
            Value   = (
                $inst.simpleSettingValue.value ??
                $inst.choiceSettingValue.value ??
                $inst.value
            )
        }
    }
}

Write-Host "🔧 Normalisiere Settings..." -ForegroundColor Cyan

$old = Get-NormalizedSettings $policy1Full
$new = Get-NormalizedSettings $policy2Full

# =========================
# FILTER (Device only)
# =========================
Write-Host "🎯 Filtere Device Settings..." -ForegroundColor Cyan

$old = $old | Where-Object { $_.Id -like "device_*" }
$new = $new | Where-Object { $_.Id -like "device_*" }

# =========================
# COMPARE
# =========================
Write-Host "🔍 Vergleiche Policies..." -ForegroundColor Cyan

# ADDED
$added = foreach ($n in $new) {
    if (-not ($old | Where-Object { $_.ShortId -eq $n.ShortId })) {
        Write-Host "🟢 NUR in $($policy2.name): $($n.ShortId)" -ForegroundColor Green
        [PSCustomObject]@{
            SettingId = $n.ShortId
            Value     = $n.Value
        }
    }
}

# REMOVED
$removed = foreach ($o in $old) {
    if (-not ($new | Where-Object { $_.ShortId -eq $o.ShortId })) {
        Write-Host "🔴 NUR in $($policy1.name): $($o.ShortId)" -ForegroundColor Red
        [PSCustomObject]@{
            SettingId = $o.ShortId
            Value     = $o.Value
        }
    }
}

# CHANGED
$changed = foreach ($n in $new) {
    $o = $old | Where-Object { $_.ShortId -eq $n.ShortId }

    if ($o -and $o.Value -ne $n.Value) {
        Write-Host "🟡 GEÄNDERT: $($n.ShortId) | ALT=$($o.Value) → NEU=$($n.Value)" -ForegroundColor Yellow
        [PSCustomObject]@{
            SettingId = $n.ShortId
            OldValue  = $o.Value
            NewValue  = $n.Value
        }
    }
}

# =========================
# SUMMARY
# =========================
Write-Host "`n📊 ZUSAMMENFASSUNG:" -ForegroundColor Cyan
Write-Host "🟢 Added:   $($added.Count)" -ForegroundColor Green
Write-Host "🔴 Removed: $($removed.Count)" -ForegroundColor Red
Write-Host "🟡 Changed: $($changed.Count)" -ForegroundColor Yellow

# =========================
# EXCEL EXPORT
# =========================
Write-Host "`n📤 Exportiere nach Excel..." -ForegroundColor Cyan

if (-not (Get-Module -ListAvailable ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

if (Test-Path $outputExcel) {
    Remove-Item $outputExcel -Force
}

$added   | Export-Excel $outputExcel -WorksheetName "Added"   -AutoSize
$removed | Export-Excel $outputExcel -WorksheetName "Removed" -AutoSize
$changed | Export-Excel $outputExcel -WorksheetName "Changed" -AutoSize

Write-Host "✅ Fertig: $outputExcel" -ForegroundColor Green