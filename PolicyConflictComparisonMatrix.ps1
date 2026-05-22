<#
.SYNOPSIS
Compare a base Intune policy against other policies and generate a matrix.
.DESCRIPTION
Loads a base policy from file or from Intune, finds other policies with overlapping
settings, and creates a comparison matrix of setting values.
Outputs a summary to the console and exports the result to Excel.
.PARAMETER BasePolicyPath
Path to a local JSON policy file to use as the base policy.
.PARAMETER BasePolicyName
Name of a policy in Intune to use as the base policy.
.PARAMETER OutputExcel
Path to the Excel file that will be written.
.NOTES
Requires Microsoft Graph PowerShell SDK and optionally the ImportExcel module.
#>
# =========================
# POLICY COMPARISON MATRIX
# =========================
# Vergleicht eine Basis-Policy gegen alle Policies des gleichen Typs
# Ausgabe: Text (Kurzmeldungen) + Excel-Matrix
#
# Usage:
# .\PolicyConflictComparisonMatrix.ps1 -BasePolicyPath "C:\path\to\policy.json"
# or
# .\PolicyConflictComparisonMatrix.ps1 -BasePolicyName "Windows 11 Baseline"

param(
    [string]$BasePolicyPath = "C:\Users\$env:USERNAME\Downloads\IntuneWindows11v4.0.0\Settings Catalog\MS Baseline_slim.json",
    [string]$BasePolicyName =  $null,
    [string]$OutputExcel = "C:\Users\$env:USERNAME\Downloads\IntuneWindows11v4.0.0\Settings Catalog\MS Baseline_Comparison.xlsx" 
)

# =========================
# CONFIGURATION
# =========================
$ErrorActionPreference = "Stop"

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
# HELPER: LOAD POLICY FROM FILE
# =========================
function Get-PolicyFromFile([string]$path) {
    Write-Host "📂 Loading policy from file: $path" -ForegroundColor Cyan
    try {
        $policy = Get-Content -Path $path -Raw | ConvertFrom-Json
        Write-Host "✔ Policy loaded: $($policy.name)" -ForegroundColor Green
        return $policy
    }
    catch {
        Write-Host "❌ Error loading policy file: $_" -ForegroundColor Red
        return $null
    }
}

# =========================
# HELPER: GET POLICY FROM ONLINE
# =========================
function Get-PolicyByName([string]$name) {
    Write-Host "🔎 Searching for policy: $name" -ForegroundColor Cyan
    try {
        $uri = "$apiBase`?`$filter=name eq '$name'"
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        $policy = $result.value | Select-Object -First 1
        
        if ($policy) {
            Write-Host "✔ Found policy: $($policy.name)" -ForegroundColor Green
            # Load full policy with settings
            $fullPolicy = Invoke-MgGraphRequest -Method GET -Uri "$apiBase/$($policy.id)`?`$expand=settings"
            return $fullPolicy
        } else {
            Write-Host "❌ Policy not found: $name" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "❌ Error fetching policy: $_" -ForegroundColor Red
        return $null
    }
}

# =========================
# HELPER: GET POLICY TEMPLATES
# =========================
function Get-PolicyTemplates() {
    Write-Host "🔎 Fetching all policy templates..." -ForegroundColor Cyan
    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        $allPolicies = @()
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        
        $allPolicies += $result.value
        
        # Handle pagination
        while ($result.'@odata.nextLink') {
            $result = Invoke-MgGraphRequest -Method GET -Uri $result.'@odata.nextLink'
            $allPolicies += $result.value
        }
        
        Write-Host "✔ Found $($allPolicies.Count) policies" -ForegroundColor Green
        return $allPolicies
    }
    catch {
        Write-Host "❌ Error fetching templates: $_" -ForegroundColor Red
        return @()
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

# Load Base Policy
$basePolicy = $null

if ($BasePolicyPath -and (Test-Path $BasePolicyPath)) {
    $basePolicy = Get-PolicyFromFile $BasePolicyPath
} elseif ($BasePolicyName) {
    $basePolicy = Get-PolicyByName $BasePolicyName
} else {
    Write-Host "❌ Neither -BasePolicyPath nor -BasePolicyName provided" -ForegroundColor Red
    Write-Host "Usage: .\PolicyComparisonMatrix.ps1 -BasePolicyName 'PolicyName'" -ForegroundColor Yellow
    Write-Host "  or  : .\PolicyComparisonMatrix.ps1 -BasePolicyPath 'C:\path\policy.json'" -ForegroundColor Yellow
    exit 1
}

if (-not $basePolicy) {
    Write-Host "❌ Failed to load base policy" -ForegroundColor Red
    exit 1
}

# Extract Base Policy Settings
Write-Host "`n🔧 Extracting base policy settings..." -ForegroundColor Cyan
$baseSettings = Get-PolicySettings $basePolicy
Write-Host "✔ Found $($baseSettings.Count) settings" -ForegroundColor Green

# Get All Policies
Write-Host "`n🔎 Fetching all policies..." -ForegroundColor Cyan
$allPolicies = Get-PolicyTemplates

# Filter policies by overlapping settingDefinitionIds
$comparePolicies = @()
$baseSettingIds = $baseSettings.Keys

Write-Host "🔍 Finding policies with overlapping settings..." -ForegroundColor Cyan

foreach ($p in $allPolicies) {
    if ($p.id -eq $basePolicy.id) { continue }  # Skip base policy itself

    try {
        # Load full policy with settings
        $fullPolicy = Invoke-MgGraphRequest -Method GET -Uri "$apiBase/$($p.id)`?`$expand=settings"
        $policySettings = Get-PolicySettings $fullPolicy
        $policySettingIds = $policySettings.Keys

        # Check for overlapping settingDefinitionIds
        $overlap = $baseSettingIds | Where-Object { $policySettingIds -contains $_ }

        if ($overlap.Count -gt 0) {
            $comparePolicies += $fullPolicy
            Write-Host "  ✓ Added: $($p.name) ($($overlap.Count) overlapping settings)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  ⚠ Skipped: $($p.name) (error loading)" -ForegroundColor DarkGray
    }
}

Write-Host "✔ Found $($comparePolicies.Count) other policies to compare" -ForegroundColor Green

# Build Comparison Matrix
Write-Host "`n📊 Building comparison matrix..." -ForegroundColor Cyan

$matrixData = @()
$policyNames = @($basePolicy.name) + $($comparePolicies | ForEach-Object { $_.name })

# Collect all setting IDs
$allSettingIds = @($baseSettings.Keys)
foreach ($p in $comparePolicies) {
    $pSettings = Get-PolicySettings $p
    $allSettingIds += $pSettings.Keys
}
$allSettingIds = $allSettingIds | Sort-Object -Unique

Write-Host "✔ Total unique settings: $($allSettingIds.Count)" -ForegroundColor Green

# Create Matrix
foreach ($settingId in $allSettingIds) {
    $row = [ordered]@{
        "Setting ID" = $settingId
    }
    
    # Base policy value
    $row[$basePolicy.name] = $baseSettings[$settingId] ?? ""
    
    # Compare policies values
    foreach ($p in $comparePolicies) {
        $pSettings = Get-PolicySettings $p
        $row[$p.name] = $pSettings[$settingId] ?? ""
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

Write-Host "`nBase Policy: $($basePolicy.name)" -ForegroundColor Yellow
Write-Host "Compared Policies: $($comparePolicies.Count)" -ForegroundColor Yellow
Write-Host "Total Settings: $($allSettingIds.Count)" -ForegroundColor Yellow

# Show comparison results
$differentCount = 0
foreach ($row in $matrixData) {
    $settingId = $row."Setting ID"
    $baseValue = $row[$basePolicy.name]
    
    $hasDifference = $false
    foreach ($policyName in $policyNames | Select-Object -Skip 1) {
        if ($row[$policyName] -ne $baseValue) {
            $hasDifference = $true
            break
        }
    }
    
    if ($hasDifference) {
        $differentCount++
        Write-Host "  ⚠ Different: $settingId" -ForegroundColor Yellow
    }
}

Write-Host "`n✔ Settings with differences: $differentCount" -ForegroundColor Green
Write-Host "✔ Settings with same values: $($allSettingIds.Count - $differentCount)" -ForegroundColor Green

# =========================
# EXCEL EXPORT
# =========================
Write-Host "`n📤 Exporting to Excel..." -ForegroundColor Cyan

# Ensure ImportExcel module
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Gray
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction SilentlyContinue
}

# Clean data for Excel export (remove problematic characters)
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
        # Fallback: Export as CSV first, then convert
        $csvPath = $OutputExcel -replace '\.xlsx$', '.csv'
        $cleanMatrixData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "✔ CSV file created as fallback: $csvPath" -ForegroundColor Green
    } catch {
        Write-Host "❌ Even CSV export failed: $_" -ForegroundColor Red
    }
}

# =========================
# FINAL SUMMARY
# =========================
Write-Host "`n" -ForegroundColor Cyan
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "✅ COMPARISON COMPLETE" -ForegroundColor Green
Write-Host "═════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Base Policy: $($basePolicy.name)" -ForegroundColor Cyan
Write-Host "Policies Compared: $($comparePolicies.Count)" -ForegroundColor Cyan
Write-Host "Settings Analyzed: $($allSettingIds.Count)" -ForegroundColor Cyan
Write-Host "Excel Output: $OutputExcel" -ForegroundColor Cyan
Write-Host ""
