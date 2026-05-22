# CIS-Intune-Policy-Toolkit

This repository contains PowerShell tools for working with CIS Intune policy JSON exports and Intune device config status. The toolkit is designed for Intune administrators who want to:

- download CIS Intune policies from the CIS Workbench
- extract and unpack the policy JSON files
- merge multiple Settings Catalog policies into one JSON payload
- compare local policy JSON files
- detect conflicts between merged policies and existing Intune policies
- export and import policies to Microsoft Intune using Microsoft Graph
- inspect device configuration and security baseline state for managed devices

## Intended workflow

1. Download CIS Intune policies from the CIS Workbench:
   `https://workbench.cisecurity.org/files?q=intune&tags=3`
2. Unpack the downloaded packages and identify the individual policy JSON files.
3. Use `MergeIntunePolicies.ps1` to merge Level 1 / Level 2 Settings Catalog policies from a folder.
4. Use `CompareTwoPolicyJSONs.ps1` or `PolicyConflictComparisonMatrix.ps1` to compare policy settings and find differences.
5. Export or import policies between JSON and Intune using `ExportPolicyFromIntune.ps1` and `ImportPolicyToIntune.ps1`.
6. Optionally use `getDeviceStatus.ps1` to inspect device configuration and baseline status for a managed device.

## Scripts

### `ExportPolicyFromIntune.ps1`
- Connects to Microsoft Graph using `Connect-MgGraph`.
- Lists available Intune configuration policies in the tenant.
- Lets you choose a policy and export it to a JSON file with a Windows Save File dialog.
- Exports the selected policy with expanded settings and writes UTF-8 JSON output.

### `ImportPolicyToIntune.ps1`
- Opens a Windows Open File dialog to select a local JSON policy file.
- Loads the policy JSON and determines the policy name.
- Checks Intune for an existing policy with the same `name` or `displayName`.
- Optionally deletes the existing policy and re-imports the JSON payload to Intune.
- Removes read-only Graph metadata properties before upload.

### `MergeIntunePolicies.ps1`
- Reads all policy JSON files from a configured folder.
- Aggregates settings from multiple policies into a single policy object.
- Detects duplicate `settingDefinitionId` values and reports conflicts with source file details.
- Keeps the first occurrence for duplicate settings and removes later duplicates.
- Writes the merged output to `mergedPolicy.json` inside the configured folder.
- Useful for building a consolidated Settings Catalog policy from CIS Level 1 and Level 2 exports.

### `CompareTwoPolicyJSONs.ps1`
- Compares two local Intune policy JSON files directly.
- Extracts settings from both files and builds a comparison matrix.
- Outputs a console summary and can optionally export results to Excel.
- Can create a slim policy JSON by removing settings that already exist in the other policy.
- No Microsoft Graph connection is required for this script.

### `PolicyConflictComparisonMatrix.ps1`
- Loads a base policy from a local file or from Intune by policy name.
- Fetches all Intune policies and compares overlapping setting definitions.
- Builds a comparison matrix of setting values across the base policy and other policies.
- Writes a console summary and exports the matrix to an Excel file.
- Helpful for discovering policy overlap, differences, and conflict candidates.

### `getDeviceStatus.ps1`
- Retrieves a managed device by name from Microsoft Graph.
- Gets device configuration state and security baseline state for that device.
- Normalizes output and displays policy status, policy name, policy ID, and last sync time.
- Useful for checking whether a device has applied configuration and baseline policies.

## Requirements

- PowerShell with the Microsoft Graph PowerShell SDK installed.
- `ImportExcel` PowerShell module for optional Excel export in `CompareTwoPolicyJSONs.ps1` and `PolicyConflictComparisonMatrix.ps1`.
- Graph permissions:
  - `DeviceManagementConfiguration.Read.All` for export, comparison, and device status reads.
  - `DeviceManagementConfiguration.ReadWrite.All` for importing policies.

## Notes

- Scripts that interact with Intune require an authenticated Microsoft Graph session.
- Update local folder paths inside `MergeIntunePolicies.ps1` and `PolicyConflictComparisonMatrix.ps1` to match your environment.
- `MergeIntunePolicies.ps1` resolves duplicate settings by keeping the first occurrence and removing duplicates.
- Review merged or slimmed policy output before importing into production tenants.
- `CompareTwoPolicyJSONs.ps1` is ideal when you want to compare policies without querying Intune.
