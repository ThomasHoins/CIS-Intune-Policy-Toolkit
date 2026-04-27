# =========================
# INTUNE POLICY IMPORT TOOL
# =========================
# Importiert eine Intune Configuration Policy aus einer JSON-Datei

# =========================
# CONNECT TO GRAPH
Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All" -ErrorAction SilentlyContinue
} catch {
    Write-Host "⚠ Already connected to Microsoft Graph" -ForegroundColor Gray
}

$apiBase = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"

# =========================
# FILE SELECTION
$defaultPath = Join-Path $env:TEMP "policy.json"
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = 'Select policy JSON file to import'
$openFileDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
$openFileDialog.FileName = [System.IO.Path]::GetFileName($defaultPath)
$openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($defaultPath)

if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "❌ Import cancelled." -ForegroundColor Yellow
    exit 0
}

$policyPath = $openFileDialog.FileName
Write-Host "📄 Selected file: $policyPath" -ForegroundColor Cyan

# =========================
# LOAD POLICY JSON
try {
    $policyJson = Get-Content $policyPath -Raw | ConvertFrom-Json
} catch {
    Write-Host "❌ Fehler beim Lesen der Datei: $_" -ForegroundColor Red
    exit 1
}

$policyName = $policyJson.name
if (-not $policyName) {
    $policyName = $policyJson.displayName
}

if (-not $policyName) {
    Write-Host "❌ Die JSON-Datei enthält keinen Namen oder DisplayName." -ForegroundColor Red
    exit 1
}

Write-Host "🔎 Policy name: $policyName" -ForegroundColor Cyan

# =========================
# CHECK EXISTING POLICY
try {
    $existing = (Invoke-MgGraphRequest -Method GET -Uri $apiBase).value |
        Where-Object { $_.name -eq $policyName -or $_.displayName -eq $policyName }
} catch {
    Write-Host "❌ Fehler beim Abrufen bestehender Policies: $_" -ForegroundColor Red
    exit 1
}

if ($existing) {
    Write-Host "⚠ Es gibt bereits eine Policy mit diesem Namen: $policyName" -ForegroundColor Yellow
    $choice = Read-Host "Soll die vorhandene Policy gelöscht und neu importiert werden? (J/N)"
    if ($choice -notin 'J','j','Y','y') {
        Write-Host "ℹ️ Import abgebrochen." -ForegroundColor Yellow
        exit 0
    }

    try {
        Invoke-MgGraphRequest -Method DELETE -Uri "$apiBase/$($existing.id)"
        Write-Host "🗑️ Bestehende Policy gelöscht." -ForegroundColor Green
    } catch {
        Write-Host "❌ Fehler beim Löschen der bestehenden Policy: $_" -ForegroundColor Red
        exit 1
    }
}

# =========================
# CLEANUP READ-ONLY PROPERTIES
$removeProps = @(
    "id",
    "createdDateTime",
    "lastModifiedDateTime",
    "version",
    "@odata.context",
    "@odata.type"
)

foreach ($p in $removeProps) {
    if ($policyJson.PSObject.Properties[$p]) {
        $policyJson.PSObject.Properties.Remove($p)
    }
}

# =========================
# IMPORT POLICY
Write-Host "`n📤 Importing policy..." -ForegroundColor Cyan
try {
    Invoke-MgGraphRequest -Method POST -Uri $apiBase `
        -Body ($policyJson | ConvertTo-Json -Depth 50)
    Write-Host "✔ Policy erfolgreich importiert." -ForegroundColor Green
} catch {
    Write-Host "❌ Fehler beim Importieren der Policy: $_" -ForegroundColor Red
    exit 1
}
