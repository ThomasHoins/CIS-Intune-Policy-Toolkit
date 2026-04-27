# =========================
# INTUNE POLICY EXPORT TOOL
# =========================
# Exportiert eine Intune Configuration Policy als JSON-Datei

# =========================
# CONNECT TO GRAPH
# =========================
Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" -ErrorAction SilentlyContinue
} catch {
    Write-Host "⚠ Already connected to Microsoft Graph" -ForegroundColor Gray
}

$apiBase = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"

# =========================
# GET ALL POLICIES
# =========================
Write-Host "`n🔎 Fetching all policies..." -ForegroundColor Cyan
try {
    $allPolicies = @()
    $uri = $apiBase
    do {
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allPolicies += $result.value
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    Write-Host "✔ Found $($allPolicies.Count) policies" -ForegroundColor Green
} catch {
    Write-Host "❌ Error fetching policies: $_" -ForegroundColor Red
    exit 1
}

# =========================
# POLICY SELECTION
# =========================
Write-Host "`n📋 Available Policies:" -ForegroundColor Cyan
for ($i = 0; $i -lt $allPolicies.Count; $i++) {
    Write-Host "  $($i + 1). $($allPolicies[$i].name)" -ForegroundColor White
}

do {
    $selection = Read-Host "`nSelect policy number (1-$($allPolicies.Count))"
    $policyIndex = [int]$selection - 1
} while ($policyIndex -lt 0 -or $policyIndex -ge $allPolicies.Count)

$selectedPolicy = $allPolicies[$policyIndex]
Write-Host "✔ Selected: $($selectedPolicy.name)" -ForegroundColor Green

# =========================
# FILE PATH SELECTION
# =========================
$defaultPath = "C:\Temp\$($selectedPolicy.name -replace '[^a-zA-Z0-9]', '_').json"

Add-Type -AssemblyName System.Windows.Forms
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Title = 'Select export file'
$saveFileDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
$saveFileDialog.FileName = [System.IO.Path]::GetFileName($defaultPath)
$saveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($defaultPath)

if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $exportPath = $saveFileDialog.FileName
} else {
    Write-Host "❌ Export cancelled." -ForegroundColor Yellow
    exit 0
}

# Ensure directory exists
$exportDir = Split-Path $exportPath -Parent
if (-not (Test-Path $exportDir)) {
    New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
}

# =========================
# EXPORT POLICY
# =========================
Write-Host "`n📤 Exporting policy..." -ForegroundColor Cyan

try {
    $policy = Invoke-MgGraphRequest -Method GET `
        -Uri "$apiBase/$($selectedPolicy.id)?`$expand=settings"

    $policy | ConvertTo-Json -Depth 50 | Out-File $exportPath -Encoding UTF8

    Write-Host "✔ Policy exported to: $exportPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Error exporting policy: $_" -ForegroundColor Red
}