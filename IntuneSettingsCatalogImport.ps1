# =========================
# CONFIG
# =========================
$policyName = "Merged Settings Catalog Policy"
$apiBase = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
$newPolicyPath = "C:\Users\thomas.hoins\Downloads\IntuneWindows11v4.0.0\Settings Catalog\mergedPolicy.json"

Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All" # -UseDeviceCode

# =========================
# LOAD NEW POLICY
# =========================
$newPolicy = Get-Content $newPolicyPath -Raw | ConvertFrom-Json

# =========================
# FIND EXISTING POLICY
# =========================
$existing = (Invoke-MgGraphRequest -Method GET -Uri $apiBase).value |
    Where-Object { $_.name -eq $policyName -or $_.displayName -eq $policyName }

if (-not $existing) {
    Write-Host "🆕 Keine bestehende Policy gefunden – wird neu erstellt" -ForegroundColor Green

    Invoke-MgGraphRequest -Method POST -Uri $apiBase `
        -Body ($newPolicy | ConvertTo-Json -Depth 40)

    return
}

# =========================
# GET FULL EXISTING POLICY (inkl. settings)
# =========================
$existingFull = Invoke-MgGraphRequest -Method GET -Uri "$apiBase/$($existing.id)"

# =========================
# EXTRACT SETTINGS
# =========================
$oldSettings = $existingFull.settings
$newSettings = $newPolicy.settings

# =========================
# CREATE DIFF
# =========================
Write-Host "`n📊 SETTINGS DIFFERENCES:" -ForegroundColor Cyan

$diff = Compare-Object `
    -ReferenceObject $oldSettings.settingInstance.settingDefinitionId `
    -DifferenceObject $newSettings.settingInstance.settingDefinitionId

foreach ($d in $diff) {
    if ($d.SideIndicator -eq "=>") {
        Write-Host "🟢 NEU: $($d.InputObject)"
    }
    elseif ($d.SideIndicator -eq "<=") {
        Write-Host "🔴 ENTFERNT: $($d.InputObject)"
    }
}

# =========================
# OPTIONAL: REBUILD POLICY (because PATCH not allowed for settings)
# =========================
Write-Host "`n🔄 Recreate Policy (Settings Catalog limitation)" -ForegroundColor Yellow

Invoke-MgGraphRequest -Method DELETE -Uri "$apiBase/$($existing.id)"

# Cleanup before POST
$removeProps = @(
    "id",
    "createdDateTime",
    "lastModifiedDateTime",
    "version",
    "@odata.context",
    "@odata.type"
)

foreach ($p in $removeProps) {
    if ($newPolicy.PSObject.Properties[$p]) {
        $newPolicy.PSObject.Properties.Remove($p)
    }
}

Invoke-MgGraphRequest -Method POST -Uri $apiBase `
    -Body ($newPolicy | ConvertTo-Json -Depth 40)

Write-Host "✅ Policy erfolgreich neu erstellt"