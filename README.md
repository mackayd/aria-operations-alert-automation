# Aria Operations alert automation

PowerShell tooling for validating, creating, updating, and enabling a defined
set of VMware Aria Operations alert definitions from a CSV condition matrix.

The repository contains a curated set of Aria Operations alert rules in a CSV
condition matrix. It also includes a read-only utility for listing Aria
Operations policy names and GUIDs.

## Author

Drew Mackay

## Repository structure

```text
aria-operations-alerts/
|-- README.md
|-- .gitignore
|-- config/
|   `-- alert-rules.csv
|-- output/
|   `-- .gitkeep
`-- scripts/
    |-- Get-AriaOpsPolicies.ps1
    `-- New-AriaOpsAlertsFromCsv.ps1
```

- `scripts/New-AriaOpsAlertsFromCsv.ps1` is the main deployment tool.
- `scripts/Get-AriaOpsPolicies.ps1` retrieves policy names and GUIDs without
  changing Aria Operations.
- `config/alert-rules.csv` is the deployable condition matrix.
- `output/` receives timestamped execution reports. Generated reports are
  excluded from Git.

## Requirements

- PowerShell 7 or later.
- Network access to the Aria Operations REST API.
- An Aria Operations account with permission to read metadata and manage alert
  definitions.
- Additional permission to change policy alert-definition settings when using
  `-PolicyId`.
- A trusted Aria Operations certificate, or `-SkipCertificateCheck` for an
  isolated lab environment.

No server names, usernames, or passwords are stored in the scripts.

## Quick start

Open PowerShell 7 in the repository root and create a credential object:

```powershell
$credential = Get-Credential
```

If `-Credential` is omitted, each script prompts securely for credentials.

### List policies and GUIDs

```powershell
.\scripts\Get-AriaOpsPolicies.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential
```

For a lab appliance with an untrusted certificate:

```powershell
.\scripts\Get-AriaOpsPolicies.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck
```

Export the policy list to CSV:

```powershell
.\scripts\Get-AriaOpsPolicies.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -OutputPath '.\output\ariaops-policies.csv'
```

The returned columns are `Name`, `PolicyGuid`, and `DefaultPolicy`.

## Validate before deployment

`-ValidateOnly` validates the complete CSV and performs live API preflight
checks without creating, updating, or enabling alert definitions:

```powershell
.\scripts\New-AriaOpsAlertsFromCsv.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -ValidateOnly
```

If `-NamePrefix` is omitted, the script prompts for it:

```text
Enter the alert definition name prefix (for example, ABC.inc Alerts):
```

The prefix must not be blank. For unattended execution, provide it explicitly:

```powershell
-NamePrefix 'ABC.inc Alerts'
```

Use `-WhatIf` to preview the intended create or update actions after live
validation:

```powershell
.\scripts\New-AriaOpsAlertsFromCsv.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -NamePrefix 'ABC.inc Alerts' `
    -WhatIf
```

## Deploy alert definitions

Create or update all deployable definitions without changing policy state:

```powershell
.\scripts\New-AriaOpsAlertsFromCsv.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -NamePrefix 'ABC.inc Alerts'
```

Create or update the definitions and enable them in a policy:

```powershell
.\scripts\New-AriaOpsAlertsFromCsv.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -NamePrefix 'ABC.inc Alerts' `
    -PolicyId '<policy-guid>'
```

Multiple policies can be supplied:

```powershell
-PolicyId @('<policy-guid-1>', '<policy-guid-2>')
```

Only CSV rules with `EnableInPolicy=true` are enabled. `MONITOR_ONLY` rules are
never created and are not enabled in policies.

### Policy assignment is separate

`-PolicyId` enables alert definitions inside an existing policy. It does not:

- make that policy the default policy;
- assign the policy to a custom group or resource;
- change policy priority.

For the alerts to evaluate, the selected policy must already apply to the
intended resources through Aria Operations policy assignment and priority.
Avoid making a test policy globally applicable unless that scope is intended.

## Deploy a subset

Use the stable CSV `RuleId` values to validate or deploy a small subset:

```powershell
.\scripts\New-AriaOpsAlertsFromCsv.ps1 `
    -Server 'ariaops.example.com' `
    -Credential $credential `
    -SkipCertificateCheck `
    -NamePrefix 'ABC.inc Alerts' `
    -IncludeRuleId 'Y02-VSPHERE-COMMON-R16' `
    -ValidateOnly
```

Multiple rule identifiers can be supplied as an array.

## Existing-definition behaviour

The default `-ExistingDefinitionAction Update` behaviour updates an existing
definition when its generated name exactly matches:

```text
<NamePrefix> - <RuleName>
```

Available behaviours are:

- `Update` updates the matching definition and preserves its ID.
- `Skip` reuses the matching definition without modifying its condition data.
- `Fail` stops that rule when a matching definition already exists.

Important: the current lookup uses the exact generated name. Changing
`-NamePrefix` does not rename existing definitions; it creates another set with
new names. Keep the prefix stable after initial deployment, or remove/rename
the previous definitions through an approved change process before deploying a
different prefix.

## Rule model

The CSV contains one row per condition. The script groups rows as follows:

1. `RuleId` identifies one logical monitoring rule.
2. `StateSeverity` separates warning and critical evaluation states.
3. `Branch` identifies alternative paths within a severity.
4. Conditions within the same branch are combined with `AND`.
5. Separate branches are combined with `OR`.

This permits simple thresholds and correlated conditions across related
resources while retaining the intended monitoring behaviour.

Supported condition types are:

- `STATIC` for fixed metric thresholds;
- `DYNAMIC` for Aria dynamic-threshold comparisons;
- `PROPERTY_NUMERIC` for numeric property comparisons;
- `COLLECTION_RATIO` for thresholds derived from the collection interval.

Supported relationships include self and related-resource evaluation, such as
descendant datastore conditions associated with a parent vSphere object.

## CSV overview

The committed configuration represents:

- 44 logical rules;
- 40 deployable alert definitions;
- 4 `MONITOR_ONLY` rules;
- 135 CSV rows in total;
- 131 persisted alert-condition rows.

Key column groups are:

| Purpose | Columns |
| --- | --- |
| Identity and source | `RuleId`, `SourceSheet`, `SourceRow`, `RuleName` |
| Deployment control | `DeploymentMode`, `EnableInPolicy`, `Fidelity` |
| Root resource | `AdapterKindKey`, `ResourceKindKey`, `MetricKey` |
| Alert classification | `AlertType`, `AlertSubType`, `ImpactType`, `ImpactDetail` |
| Logic | `StateSeverity`, `Branch`, `ConditionId` |
| Condition target | `ConditionAdapterKindKey`, `ConditionResourceKindKey`, `Relation`, `Aggregation` |
| Threshold | `ConditionType`, `ConditionKey`, `Operator`, `Value`, `ValueMultiplier`, `ThresholdType` |
| Timing | `AlertWaitCycles`, `AlertCancelCycles`, `ConditionWaitCycles`, `ConditionCancelCycles` |
| Traceability | `SourceDecision`, `Notes` |

Do not edit the CSV header names. The script validates required columns,
booleans, integers, numeric values, operators, condition identities, consistent
rule metadata, and live Aria resource/metric metadata before making changes.

## Collection interval

`COLLECTION_RATIO` conditions derive their threshold from the adapter collection
interval. The default is 300 seconds. Override it when the environment uses a
different interval:

```powershell
-CollectionIntervalSeconds 600
```

Using an incorrect value changes the effective threshold for those rules.

## Metric preflight

Live metadata validation verifies adapters, resource kinds, metrics, and
properties before deployment. It can be bypassed with
`-SkipMetricValidation`, but this should normally be used only for controlled
diagnostics because it removes an important deployment safeguard.

## Results and exit behaviour

Each execution writes a timestamped CSV report beneath `output/`. Important
columns include:

- `RuleId` and resource identity;
- `Status`;
- `AlertAction` (`CREATED`, `UPDATED`, `SKIPPED_EXISTING`, or validation/WhatIf
  equivalents);
- `PolicyAction` (`ENABLED`, `NOT_REQUESTED`, or another explicit outcome);
- `AlertDefinitionId`;
- `ConditionCount`;
- `Error`.

The script throws a terminating error if any processed rule fails. Review the
timestamped result file before treating a deployment as successful.

## Security guidance

- Use `Get-Credential` or a securely sourced `PSCredential`.
- Do not commit exported credential objects, passwords, API tokens, or `.env`
  files. Common secret-file patterns are excluded by `.gitignore`.
- Use a least-privilege service account for production automation.
- Use a trusted certificate in production.
- `-SkipCertificateCheck` is intended only for controlled lab environments.
- Keep execution reports out of Git because they contain environment-specific
  definition identifiers.

## Troubleshooting

### A different prefix creates duplicates

The existing-definition lookup is name-based. Reuse the original prefix, or
remove/rename the old definitions before deploying with a new prefix.

### Definitions exist but alerts do not evaluate

Confirm all of the following:

1. The execution report shows `PolicyAction` as `ENABLED`.
2. The selected policy is assigned to the relevant custom group or resources.
3. The policy has the required priority.
4. The resources are collecting the referenced metrics or properties.
5. The configured wait cycles have elapsed while the conditions remain true.

### Certificate validation fails

Install and trust the appliance certificate. In a lab only, rerun with
`-SkipCertificateCheck`.

### A metric or resource kind fails preflight

Check that the relevant management pack or adapter is installed and collecting
data. Confirm the exact metric key in Aria Operations before editing the CSV.

### Authentication fails

Create a new credential object and verify that the account can access the Aria
Operations API:

```powershell
$credential = Get-Credential
```

Credentials are not read from the CSV or stored by either script.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for the
full terms.
