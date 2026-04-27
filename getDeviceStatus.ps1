param(
    [Parameter(Mandatory=$true)]
    [string]$DeviceName
)

# ----------------------------
# 1. Device holen
# ----------------------------

$uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$deviceName'"
$response = Invoke-MgGraphRequest -Method GET -Uri $uri
$device = $response.value[0]
$device | Select-Object id, deviceName, lastSyncDateTime
$deviceId = $device.Id
Write-Host "Device: $($device.DeviceName)" -ForegroundColor Green

# ----------------------------
# 2. Device Configuration States und Baseline States
# ----------------------------
$response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceId/deviceConfigurationStates"

$baselineStates = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceId/securityBaselineStates"

Write-Host "Baseline States gefunden: $($baselineStates.value.Count)" -ForegroundColor Cyan

# ----------------------------
# 3. Output normalisieren
# ----------------------------
$result = foreach ($p in $response.value) {

    [PSCustomObject]@{
        PolicyType = "Device Configuration"
        PolicyName = $p.displayName
        PolicyId   = $p.id
        State      = $p.state   # succeeded / error / conflict / pending
        LastSync   = $device.lastSyncDateTime
    }
}

foreach ($b in $baselineStates.value) {
    $result += [PSCustomObject]@{
        PolicyType = "Security Baseline"
        PolicyName = $b.displayName
        PolicyId   = $b.id
        State      = $b.state
        LastSync   = $device.lastSyncDateTime
    }
}

if ($baselineStates.value.Count -eq 0) {
    Write-Host "Keine Baseline Policies gefunden für dieses Gerät." -ForegroundColor Yellow
}

# ----------------------------
# 4. Ausgabe
# ----------------------------
$result |
    Sort-Object State, PolicyName |
    Format-Table -AutoSize