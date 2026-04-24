$policyId = "5d31808c-fcb8-457b-a82e-1977236a9420"

$policy = Invoke-MgGraphRequest -Method GET `
-Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId`?`$expand=settings"

$policy | ConvertTo-Json -Depth 50 | Out-File "C:\Users\thomas.hoins\Downloads\IntuneWindows11v4.0.0\Settings Catalog\MS Baseline.json"