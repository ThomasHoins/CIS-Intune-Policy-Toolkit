# CIS-Intune-Policy-Toolkit

This repository contains PowerShell tools for working with CIS Intune policy JSON exports. The toolkit is designed for Intune administrators who want to:

- download CIS Intune policies from the CIS Workbench
- extract and unpack the policy JSON files
- merge Level 1 and Level 2 Settings Catalog policies
- detect conflicts between merged policies and existing Intune policies
- export and import policies to Microsoft Intune using Microsoft Graph

## Intended workflow

1. Download CIS Intune policies from the CIS Workbench:
   `https://workbench.cisecurity.org/files?q=intune&tags=3`
2. Unpack the downloaded packages and identify the individual policy JSON files.
3. Use this toolkit to merge selected Level 1 or Level 2 policies into a single JSON payload.
4. Analyze conflicts between the merged policy and any existing Intune policies.
5. Upload the merged policy into Intune or export/import policies as needed.

## Scripts

### `ExportPolicyFromIntune.ps1`
- Connects to Microsoft Graph using `Connect-MgGraph`.
- Lists available Intune configuration policies from the tenant.
- Lets you choose a policy and export it to a JSON file via a Windows Save File dialog.
- Includes the policy settings using `$expand=settings` and writes UTF-8 JSON output.

### `ImportPolicyToIntune.ps1`
- Opens a Windows Open File dialog to select a local JSON policy file.
- Loads the selected JSON and determines the policy name.
- Checks Intune for an existing policy with the same `name` or `displayName`.
- Optionally deletes the existing policy and re-imports the JSON to Intune.
- Cleans up read-only Graph metadata properties before uploading.

### `MergeIntunePolicies.ps1`
- Reads all policy JSON files from a configured folder.
- Merges settings from multiple policies into a single policy object.
- Detects duplicate `settingDefinitionId` values and reports conflicts.
- Keeps the first occurrence for duplicate settings and removes duplicates.
- Writes the merged output to `mergedPolicy.json`.
- Intended for combining Level 1 and Level 2 Settings Catalog policies.

### `PolicyConflictComparisonMatrix.ps1`
- Loads a base policy either from a file or by name from Intune.
- Fetches all Intune policies and compares overlapping settings.
- Builds a comparison matrix of setting values across policies.
- Writes a summary to the console and exports results to an Excel file.
- Useful for finding policy overlap, differences, and potential conflicts.

## Requirements

- PowerShell with the Microsoft Graph PowerShell SDK installed.
- Permissions:
  - `DeviceManagementConfiguration.Read.All` for export and comparison.
  - `DeviceManagementConfiguration.ReadWrite.All` for import.
- `ImportExcel` PowerShell module for Excel export in `PolicyConflictComparisonMatrix.ps1`.

## Notes

- The scripts are intended for administrators already familiar with Intune policy structure and the CIS Settings Catalog format.
- Adjust the folder paths inside the scripts to match your local download and unpack locations.
- The merge logic currently resolves duplicates by taking the first encountered setting and removing later duplicates.
- Review merged output before importing into production tenants.
