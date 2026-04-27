<#
.SYNOPSIS
Merge multiple CIS Intune Settings Catalog JSON policies into a single policy.
.DESCRIPTION
Reads all JSON policy files from a configured folder, detects duplicates and conflicts
by settingDefinitionId, merges settings into one output policy, and writes
mergedPolicy.json for later import.
.NOTES
Adjust $folderPath to the local CIS policy JSON folder before running.
#>
$folderPath = "C:\temp\IntuneWindows11v4.0.0\Settings Catalog\Level 1&2"

$files = Get-ChildItem -Path $folderPath -Filter *.json

$allSettings = @()
$basePolicy = $null

foreach ($file in $files) {
    Write-Host "Lade $($file.Name)..."

    $json = Get-Content $file.FullName -Raw | ConvertFrom-Json

    if (-not $basePolicy) {
        $basePolicy = $json
    }

    if ($json.settings) {
        foreach ($setting in $json.settings) {
            # Quelle merken (für Konfliktanzeige)
            $setting | Add-Member -NotePropertyName SourceFile -NotePropertyValue $file.Name
            $allSettings += $setting
        }
    }
}


# 🔍 Konflikte erkennen (detailliert)
$conflicts = $allSettings | Group-Object `
    { $_.settingInstance.settingDefinitionId } |
    Where-Object { $_.Count -gt 1 }

if ($conflicts) {
    Write-Host "`n⚠️ Konflikte gefunden:" -ForegroundColor Yellow

    foreach ($c in $conflicts) {
        Write-Host "`n====================================" -ForegroundColor DarkGray
        Write-Host "SettingDefinitionId:" $c.Name -ForegroundColor Cyan

        foreach ($item in $c.Group) {

            # Wert extrahieren (funktioniert für die meisten Typen)
            $value = $null

            if ($item.settingInstance.simpleSettingValue) {
                $value = $item.settingInstance.simpleSettingValue.value
            }
            elseif ($item.settingInstance.choiceSettingValue) {
                $value = $item.settingInstance.choiceSettingValue.value
            }
            elseif ($item.settingInstance.collectionSettingValue) {
                $value = ($item.settingInstance.collectionSettingValue.values) -join ", "
            }
            else {
                $value = "[komplex / nicht direkt lesbar]"
            }

            Write-Host "  Datei : $($item.SourceFile)" -ForegroundColor Gray
            Write-Host "  Wert  : $value" -ForegroundColor White
            Write-Host ""
        }
    }
}
else {
    Write-Host "✅ Keine Konflikte gefunden"
}

# 🧹 Duplikate entfernen (erste gewinnt)
$dedupedSettings = $allSettings | Sort-Object `
    @{Expression = { $_.settingInstance.settingDefinitionId }} -Unique
$dedupedSettings | ForEach-Object {
    if ($_.PSObject.Properties["SourceFile"]) {
        $_.PSObject.Properties.Remove("SourceFile")
    }
}

# 🧱 Neue Policy bauen
$basePolicy.settings = $dedupedSettings

# =========================
# REMOVE INTUNE READ-ONLY FIELDS
# =========================
$removeProps = @(
    "id",
    "createdDateTime",
    "lastModifiedDateTime",
    "version",
    "@odata.context",
    "@odata.type",
    "status",
    "roleScopeTagIds"
)

$basePolicy.PSObject.Properties.Remove("id")
$basePolicy.PSObject.Properties.Remove("createdDateTime")
$basePolicy.PSObject.Properties.Remove("lastModifiedDateTime")
$basePolicy.PSObject.Properties.Remove("version")
$basePolicy.PSObject.Properties.Remove("@odata.context")
$basePolicy.PSObject.Properties.Remove("@odata.type")

# Settings-Array bereinigen (wichtig!)
foreach ($setting in $basePolicy.settings) {
    foreach ($prop in $removeProps) {
        if ($setting.PSObject.Properties[$prop]) {
            $setting.PSObject.Properties.Remove($prop)
        }
    }

    # Sicherheit: SourceFile entfernen (falls noch drin)
    if ($setting.PSObject.Properties["SourceFile"]) {
        $setting.PSObject.Properties.Remove("SourceFile")
    }
}

if ($basePolicy.PSObject.Properties.Name -contains "displayName") {
    $basePolicy.displayName = "Merged Settings Catalog Policy"
}
elseif ($basePolicy.PSObject.Properties.Name -contains "name") {
    $basePolicy.name = "Merged Settings Catalog Policy"
}
else {
    $basePolicy | Add-Member -NotePropertyName displayName -NotePropertyValue "Merged Settings Catalog Policy"
}
$basePolicy.description = "Auto-merged from multiple JSON files"
If ($basePolicy.id ) {$basePolicy.id = $null}

$basePolicy.PSObject.Properties.Remove("@odata.context")

# 💾 Speichern
$outputPath = Join-Path $folderPath "mergedPolicy.json"
$basePolicy | ConvertTo-Json -Depth 25 | Out-File $outputPath -Encoding utf8

Write-Host "`n✅ Fertig: $outputPath"