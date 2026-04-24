# =========================
# CONFIG
# =========================
$oldJsonPath = "C:\Users\thomas.hoins\Downloads\IntuneWindows11v4.0.0\Settings Catalog\mergedPolicy.json"
$newJsonPath = "C:\Users\thomas.hoins\Downloads\IntuneWindows11v4.0.0\Settings Catalog\MS Baseline.json"
$outputExcel  = "C:\Users\thomas.hoins\Downloads\IntuneWindows11v4.0.0\Settings Catalog\IntunePolicyDiff.xlsx"

# =========================
# LOAD JSON
# =========================
$oldPolicy = Get-Content $oldJsonPath -Raw | ConvertFrom-Json
$newPolicy = Get-Content $newJsonPath -Raw | ConvertFrom-Json

# =========================
# EXTRACT SETTINGS
# =========================
function Get-NormalizedSettings($policy) {
    return $policy.settings.settingInstance | ForEach-Object {
        [PSCustomObject]@{
            Id    = $_.settingDefinitionId
            Value = (
                $_.simpleSettingValue.value ??
                $_.choiceSettingValue.value ??
                $_.value
            )
        }
    }
}

$old = Get-NormalizedSettings $oldPolicy
$new = Get-NormalizedSettings $newPolicy

# =========================
# COMPARE
# =========================
# =========================
# ADDED
# =========================
$added = foreach ($n in $new) {
    if (-not ($old | Where-Object { $_.Id -eq $n.Id })) {
        Write-Host "🟢 Nicht in $([System.IO.Path]::GetFileName($newJsonPath)) : $($n.Id)" -ForegroundColor Green
        [PSCustomObject]@{
            SettingId = $n.Id
            Value     = $n.Value
        }
    }
}

# =========================
# REMOVED
# =========================
$removed = foreach ($o in $old) {
    if (-not ($new | Where-Object { $_.Id -eq $o.Id })) {
        Write-Host "� Nicht in $([System.IO.Path]::GetFileName($oldJsonPath)) : $($o.Id)" -ForegroundColor Red
        [PSCustomObject]@{
            SettingId = $o.Id
            Value     = $o.Value
        }
    }
}

# =========================
# CHANGED
# =========================
$changed = foreach ($n in $new) {
    $o = $old | Where-Object { $_.Id -eq $n.Id }

    if ($o -and $o.Value -ne $n.Value) {
        Write-Host "🟢 GEÄNDERT: $($n.Id)" -ForegroundColor Yellow
        [PSCustomObject]@{
            SettingId = $n.Id
            OldValue  = $o.Value
            NewValue  = $n.Value
        }
    }
}

# =========================
# EXCEL EXPORT
# =========================

# Install module if needed
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

# Remove old file
if (Test-Path $outputExcel) {
    Remove-Item $outputExcel -Force
}

# Write sheets
$added   | Export-Excel $outputExcel -WorksheetName "Added"   -AutoSize
$removed | Export-Excel $outputExcel -WorksheetName "Removed" -AutoSize
$changed | Export-Excel $outputExcel -WorksheetName "Changed" -AutoSize

Write-Host "✅ Excel Report erstellt: $outputExcel" -ForegroundColor Green