# =========================
# CONFIG
# =========================
$policyName1 = "CIS Intune Baseline (L2)  4.0.0"
$policyName2 = "Windows 11 Baseline"
$outputExcel = "C:\Temp\IntunePolicyDiff_Live.xlsx"

# =========================
# CONNECT GRAPH
# =========================
Write-Host "🔐 Verbinde zu Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"

$apiBase = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"

# =========================
# HELPER: POLICY LOOKUP
# =========================
function Get-PolicyByName($name) {
    try {
        $uri = "$apiBase`?`$filter=name eq '$name'"
        (Invoke-MgGraphRequest -Method GET -Uri $uri).value
    }
    catch {
        Write-Host "❌ Fehler beim Abrufen der Policy '$name'" -ForegroundColor Red
        $null
    }
}

# =========================
# HELPER: TYPE ERKENNEN
# =========================
function Get-SettingType($id) {
    switch -Regex ($id) {
        "^device_vendor_msft" { "SettingsCatalog" }
        "^admxtemplate"       { "ADMX" }
        default               { "Unknown" }
    }
}

# =========================
# HELPER: POLICY AXIS
# =========================
function Get-PolicyAxis($id) {
    switch -Regex ($id.ToLower()) {
        "defender|antivirus|asr" { "MicrosoftDefender" }
        "bitlocker"              { "BitLocker" }
        "credentialguard|lsa"    { "CredentialGuard" }
        "smart|smartscreen"      { "SmartScreen" }
        "firewall"               { "Firewall" }
        "uac"                    { "UAC" }
        "windowsupdate|wu"       { "WindowsUpdate" }
        default                  { "Other" }
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

Write-Host "✔ Policy 1: $($policy1.name)" -ForegroundColor Green
Write-Host "✔ Policy 2: $($policy2.name)" -ForegroundColor Green

# =========================
# LOAD FULL POLICIES
# =========================
Write-Host "📥 Lade vollständige Policies..." -ForegroundColor Cyan

$policy1Full = Invoke-MgGraphRequest -Method GET -Uri "$apiBase/$($policy1.id)?`$expand=settings"
$policy2Full = Invoke-MgGraphRequest -Method GET -Uri "$apiBase/$($policy2.id)?`$expand=settings"

# =========================
# NORMALIZE SETTINGS
# =========================
function Get-NormalizedSettings($policy) {
    foreach ($s in $policy.settings) {

        $inst = $s.settingInstance
        $id   = $inst.settingDefinitionId
        $type = Get-SettingType $id

        [PSCustomObject]@{
            PolicyName = $policy.name
            Id         = $id
            Type       = $type
            PolicyAxis = Get-PolicyAxis $id
            Value      = (
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
# FILTER DEVICE SETTINGS
# =========================
$old = $old | Where-Object { $_.Id -like "device_*" -or $_.Id -like "admxtemplate_*" }
$new = $new | Where-Object { $_.Id -like "device_*" -or $_.Id -like "admxtemplate_*" }

# =========================
# COMPARE
# =========================

# ADDED
$added = foreach ($n in $new) {
    if (-not ($old | Where-Object { $_.Id -eq $n.Id })) {
        [PSCustomObject]@{
            Setting  = $n.Id
            Type     = $n.Type
            Axis     = $n.PolicyAxis
            Value    = $n.Value
        }
    }
}

# REMOVED
$removed = foreach ($o in $old) {
    if (-not ($new | Where-Object { $_.Id -eq $o.Id })) {
        [PSCustomObject]@{
            Setting  = $o.Id
            Type     = $o.Type
            Axis     = $o.PolicyAxis
            Value    = $o.Value
        }
    }
}

# CHANGED
$changed = foreach ($n in $new) {
    $o = $old | Where-Object { $_.Id -eq $n.Id }
    if ($o -and $o.Value -ne $n.Value) {
        [PSCustomObject]@{
            Setting   = $n.Id
            Type      = $n.Type
            Axis      = $n.PolicyAxis
            OldValue  = $o.Value
            NewValue  = $n.Value
        }
    }
}

# =========================
# MODERN ↔ ADMX CONFLICTS
# =========================
Write-Host "⚠ Suche Modern ↔ ADMX Konflikte..." -ForegroundColor Magenta

$conflicts = foreach ($m in $new | Where-Object Type -eq "SettingsCatalog") {

    $matches = $old | Where-Object {
        $_.PolicyAxis -eq $m.PolicyAxis -and
        $_.Type -eq "ADMX" -and
        $_.Value -ne $m.Value
    }

    foreach ($a in $matches) {
        [PSCustomObject]@{
            PolicyAxis      = $m.PolicyAxis
            ModernSetting   = $m.Id
            ModernValue     = $m.Value
            ADMXSetting     = $a.Id
            ADMXValue       = $a.Value
            Recommendation  = "Remove ADMX, keep Settings Catalog"
        }
    }
}

# =========================
# SUMMARY
# =========================
Write-Host "`n📊 ZUSAMMENFASSUNG:" -ForegroundColor Cyan
Write-Host "🟢 Added:     $($added.Count)" -ForegroundColor Green
Write-Host "🔴 Removed:   $($removed.Count)" -ForegroundColor Red
Write-Host "🟡 Changed:   $($changed.Count)" -ForegroundColor Yellow
Write-Host "⚠ Conflicts: $($conflicts.Count)" -ForegroundColor Magenta

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

$added     | Export-Excel $outputExcel -WorksheetName "Added"     -AutoSize
$removed   | Export-Excel $outputExcel -WorksheetName "Removed"   -AutoSize
$changed   | Export-Excel $outputExcel -WorksheetName "Changed"   -AutoSize
$conflicts | Export-Excel $outputExcel -WorksheetName "Conflicts" -AutoSize -FreezeTopRow

Write-Host "✅ Fertig: $outputExcel" -ForegroundColor Green
