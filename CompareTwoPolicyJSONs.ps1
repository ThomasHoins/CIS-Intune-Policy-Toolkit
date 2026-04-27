<#
.SYNOPSIS
Compare two Intune policy JSON files directly.
.DESCRIPTION
Loads two local JSON policy files, extracts their settings, and creates a comparison
matrix showing differences between the policies. Outputs a summary to the console
and optionally exports the result to Excel. Can also create a "slim" version of one
policy by removing settings that already exist in the other policy.
.PARAMETER Policy1Path
Path to the first JSON policy file.
.PARAMETER Policy2Path
Path to the second JSON policy file.
.PARAMETER OutputExcel
Path to the Excel file that will be written (optional).
.PARAMETER CreateSlimPolicy
Specify which policy to slim ("Policy1" or "Policy2"). Creates a new JSON file
with "_slim" suffix containing only settings not present in the other policy.
.NOTES
No Microsoft Graph connection required. Requires the ImportExcel module for Excel export.
#>
# =========================
# COMPARE TWO POLICY JSONs
# =========================
# Vergleicht zwei lokale JSON-Policy-Dateien direkt
# Ausgabe: Text (Kurzmeldungen) + optionale Excel-Matrix
#
# Usage:
# .\CompareTwoPolicyJSONs.ps1 -Policy1Path "C:\path\to\policy1.json" -Policy2Path "C:\path\to\policy2.json"
# .\CompareTwoPolicyJSONs.ps1 -Policy1Path "C:\path\to\policy1.json" -Policy2Path "C:\path\to\policy2.json" -OutputExcel "C:\temp\Comparison.xlsx"
# .\CompareTwoPolicyJSONs.ps1 -Policy1Path "C:\path\to\policy1.json" -Policy2Path "C:\path\to\policy2.json" -CreateSlimPolicy "Policy1"

param(
    [Parameter(Mandatory=$true)]
    [string]$Policy1Path,
    [Parameter(Mandatory=$true)]
    [string]$Policy2Path,
    [string]$OutputExcel = $null,
    [ValidateSet("Policy1", "Policy2")]
    [string]$CreateSlimPolicy = $null
)

# =========================
# CONFIGURATION
# =========================
$ErrorActionPreference = "Stop"

# =========================
# HELPER: LOAD POLICY FROM FILE
# =========================
function Get-PolicyFromFile([string]$path) {
    Write-Host "📂 Loading policy from file: $path" -ForegroundColor Cyan
    try {
        $policy = Get-Content -Path $path -Raw | ConvertFrom-Json
        $policyName = $policy.name ?? $policy.displayName ?? "Unknown Policy"
        Write-Host "✔ Policy loaded: $policyName" -ForegroundColor Green
        return $policy
    }
    catch {
        Write-Host "❌ Error loading policy file: $_" -ForegroundColor Red
        return $null
    }
}

# ===========================
# HELPER: GET POLICY SETTINGS
# ===========================
function Get-PolicySettings([object]$policy) {
    $settings = @{}

    if ($policy.settings) {
        foreach ($setting in $policy.settings) {
            $inst = $setting.settingInstance
            $id = $inst.settingDefinitionId

            $value = (
                $inst.simpleSettingValue.value ??
                $inst.choiceSettingValue.value ??
                $inst.groupSettingCollectionValue ??
                $inst.value ??
                "N/A"
            )

            $settings[$id] = Get-CleanValue -value $value -settingId $id
        }
    }

    return $settings
}

# ===========================
# HELPER: GET CLEAN VALUE
# ===========================
function Get-CleanValue([string]$value, [string]$settingId) {
    if (-not $value -or $value -eq "N/A") { return $value }

    if ($settingId) {
        $prefix = $settingId + "_"
        if ($value.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $value.Substring($prefix.Length)
        }

        if ($value -eq $settingId) {
            return ""
        }
    }

    return $value
}

# =========================
# MAIN LOGIC
# =========================

# Validate input files
if (-not (Test-Path $Policy1Path)) {
    Write-Host "❌ Policy1 file not found: $Policy1Path" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $Policy2Path)) {
    Write-Host "❌ Policy2 file not found: $Policy2Path" -ForegroundColor Red
    exit 1
}

# Load both policies
Write-Host "`n🔎 Loading policies..." -ForegroundColor Cyan
$policy1 = Get-PolicyFromFile $Policy1Path
$policy2 = Get-PolicyFromFile $Policy2Path

if (-not $policy1 -or -not $policy2) {
    Write-Host "❌ Failed to load one or both policies" -ForegroundColor Red
    exit 1
}

$policy1Name = $policy1.name ?? $policy1.displayName ?? "Policy 1"
$policy2Name = $policy2.name ?? $policy2.displayName ?? "Policy 2"

# Extract settings
Write-Host "`n🔧 Extracting policy settings..." -ForegroundColor Cyan
$policy1Settings = Get-PolicySettings $policy1
$policy2Settings = Get-PolicySettings $policy2

Write-Host "✔ Policy 1: $($policy1Settings.Count) settings" -ForegroundColor Green
Write-Host "✔ Policy 2: $($policy2Settings.Count) settings" -ForegroundColor Green

# =========================
# CREATE SLIM POLICY (Optional)
# =========================
$removedSettings = @()
if ($CreateSlimPolicy) {
    Write-Host "`n🔧 Creating slim policy..." -ForegroundColor Cyan

    # Determine which policy to slim and which to use as reference
    if ($CreateSlimPolicy -eq "Policy1") {
        $sourcePolicy = $policy1
        $sourceSettings = $policy1Settings
        $referenceSettings = $policy2Settings
        $sourceName = $policy1Name
        $sourcePath = $Policy1Path
    } else {
        $sourcePolicy = $policy2
        $sourceSettings = $policy2Settings
        $referenceSettings = $policy1Settings
        $sourceName = $policy2Name
        $sourcePath = $Policy2Path
    }

    # Clone the source policy
    $slimPolicy = $sourcePolicy | ConvertTo-Json -Depth 50 | ConvertFrom-Json

    # Filter settings to keep only those not in reference policy or with different values
    $settingsToKeep = @()
    foreach ($setting in $slimPolicy.settings) {
        $settingId = $setting.settingInstance.settingDefinitionId
        $currentValue = Get-CleanValue -value (
            $setting.settingInstance.simpleSettingValue.value ??
            $setting.settingInstance.choiceSettingValue.value ??
            $setting.settingInstance.groupSettingCollectionValue ??
            $setting.settingInstance.value ??
            "N/A"
        ) -settingId $settingId

        $referenceValue = $referenceSettings[$settingId]

        # Keep the setting if it's not in reference policy, or if it has a different value
        if (-not $referenceSettings.ContainsKey($settingId) -or $currentValue -ne $referenceValue) {
            $settingsToKeep += $setting
        } else {
            $removedSettings += $settingId
        }
    }

    $slimPolicy.settings = $settingsToKeep

    # Update policy name
    if ($slimPolicy.PSObject.Properties.Name -contains "name") {
        $slimPolicy.name = "$sourceName`_slim"
    }
    if ($slimPolicy.PSObject.Properties.Name -contains "displayName") {
        $slimPolicy.displayName = "$sourceName`_slim"
    }

    # Generate output path
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $directory = [System.IO.Path]::GetDirectoryName($sourcePath)
    $slimPath = Join-Path $directory ($baseName + "_slim.json")

    # Save slim policy
    try {
        $slimPolicy | ConvertTo-Json -Depth 50 | Out-File $slimPath -Encoding UTF8
        Write-Host "✔ Slim policy created: $slimPath" -ForegroundColor Green
        Write-Host "  Original settings: $($sourceSettings.Count)" -ForegroundColor Gray
        Write-Host "  Slim settings: $($settingsToKeep.Count)" -ForegroundColor Gray
        Write-Host "  Removed settings: $($removedSettings.Count)" -ForegroundColor Gray
    } catch {
        Write-Host "❌ Error creating slim policy: $_" -ForegroundColor Red
    }
}

# Build comparison matrix
Write-Host "`n📊 Building comparison matrix..." -ForegroundColor Cyan

$matrixData = @()
$allSettingIds = @($policy1Settings.Keys) + @($policy2Settings.Keys) | Sort-Object -Unique

Write-Host "✔ Total unique settings: $($allSettingIds.Count)" -ForegroundColor Green

# Create matrix rows
foreach ($settingId in $allSettingIds) {
    $row = [ordered]@{
        "Setting ID" = $settingId
        $policy1Name = $policy1Settings[$settingId] ?? ""
        $policy2Name = $policy2Settings[$settingId] ?? ""
    }

    # Add "Removed" column if slim policy was created
    if ($CreateSlimPolicy) {
        $row["Removed"] = if ($removedSettings -contains $settingId) { "Yes" } else { "No" }
    }

    $matrixData += [PSCustomObject]$row
}

# =========================
# TEXT OUTPUT (Status Messages)
# =========================
Write-Host "`n" -ForegroundColor Cyan
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "📋 COMPARISON SUMMARY" -ForegroundColor Cyan
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan

Write-Host "`nPolicy 1: $policy1Name" -ForegroundColor Yellow
Write-Host "Policy 2: $policy2Name" -ForegroundColor Yellow
Write-Host "Total Settings: $($allSettingIds.Count)" -ForegroundColor Yellow

# Analyze differences
$onlyInPolicy1 = @()
$onlyInPolicy2 = @()
$differentValues = @()
$sameValues = @()

foreach ($row in $matrixData) {
    $settingId = $row."Setting ID"
    $value1 = $row[$policy1Name]
    $value2 = $row[$policy2Name]

    if ($value1 -and -not $value2) {
        $onlyInPolicy1 += $settingId
    }
    elseif (-not $value1 -and $value2) {
        $onlyInPolicy2 += $settingId
    }
    elseif ($value1 -and $value2) {
        if ($value1 -eq $value2) {
            $sameValues += $settingId
        } else {
            $differentValues += $settingId
        }
    }
}

Write-Host "`n📈 Comparison Results:" -ForegroundColor Cyan
Write-Host "  ⚠ Settings only in Policy 1: $($onlyInPolicy1.Count)" -ForegroundColor Yellow
Write-Host "  ⚠ Settings only in Policy 2: $($onlyInPolicy2.Count)" -ForegroundColor Yellow
Write-Host "  ⚠ Settings with different values: $($differentValues.Count)" -ForegroundColor Yellow
Write-Host "  ✅ Settings with same values: $($sameValues.Count)" -ForegroundColor Green

# =========================
# EXCEL EXPORT (Optional)
# =========================
if ($OutputExcel) {
    Write-Host "`n📤 Exporting to Excel..." -ForegroundColor Cyan

    # Ensure ImportExcel module
    if (-not (Get-Module -ListAvailable ImportExcel)) {
        Write-Host "Installing ImportExcel module..." -ForegroundColor Gray
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    }

    # Clean data for Excel export
    $cleanMatrixData = $matrixData | ForEach-Object {
        $row = [ordered]@{}
        foreach ($prop in $_.PSObject.Properties) {
            $cleanValue = $prop.Value

            # Remove null bytes and other problematic characters
            if ($cleanValue -is [string]) {
                $cleanValue = $cleanValue -replace '\x00', '' -replace '[\x00-\x1F\x7F-\x9F]', ''
            }

            # Limit string length to prevent Excel issues
            if ($cleanValue -is [string] -and $cleanValue.Length -gt 32000) {
                $cleanValue = $cleanValue.Substring(0, 32000) + "..."
            }

            $row[$prop.Name] = $cleanValue
        }
        [PSCustomObject]$row
    }

    try {
        # Export to Excel with bold headers
        $cleanMatrixData | Export-Excel -Path $OutputExcel -WorksheetName "Comparison" -AutoSize -BoldTopRow -ErrorAction Stop

        Write-Host "✔ Excel file created: $OutputExcel" -ForegroundColor Green
    } catch {
        Write-Host "❌ Error creating Excel file: $_" -ForegroundColor Red
        Write-Host "Trying alternative export method..." -ForegroundColor Yellow

        try {
            # Fallback: Export as CSV
            $csvPath = $OutputExcel -replace '\.xlsx$', '.csv'
            $cleanMatrixData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "✔ CSV file created as fallback: $csvPath" -ForegroundColor Green
        } catch {
            Write-Host "❌ Even CSV export failed: $_" -ForegroundColor Red
        }
    }
}

# =========================
# FINAL SUMMARY
# =========================
Write-Host "`n" -ForegroundColor Cyan
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "✅ COMPARISON COMPLETE" -ForegroundColor Green
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Policy 1: $policy1Name ($($policy1Settings.Count) settings)" -ForegroundColor Cyan
Write-Host "Policy 2: $policy2Name ($($policy2Settings.Count) settings)" -ForegroundColor Cyan
Write-Host "Total Unique Settings: $($allSettingIds.Count)" -ForegroundColor Cyan
if ($OutputExcel) {
    Write-Host "Excel Output: $OutputExcel" -ForegroundColor Cyan
}
Write-Host ""