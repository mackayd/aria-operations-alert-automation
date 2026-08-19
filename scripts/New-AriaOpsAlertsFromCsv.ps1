#Requires -Version 7.0

<#
.SYNOPSIS
    Creates or updates Aria Operations alert definitions from the extended
    condition-matrix CSV generated from the highlighted workbook rows.

.DESCRIPTION
    The input CSV stores one row per condition. Conditions are grouped by RuleId,
    StateSeverity, and Branch. Conditions inside a branch use AND; branches use
    OR. Warning and Critical condition severities are placed in one AUTO alert
    state, allowing the most severe contributing condition to determine the
    alert severity.

    Supported condition types are STATIC, DYNAMIC, PROPERTY_NUMERIC, and
    COLLECTION_RATIO. MONITOR_ONLY rules are validated but deliberately do not
    create alerts. This preserves the workbook guidance for metrics that are
    useful as context but unsafe as standalone alerts.

    Server and credentials are never stored in this script. CsvPath defaults to
    config/alert-rules.csv in the repository root. Credentials can be supplied
    as a PSCredential or entered securely when prompted.

.PARAMETER Server
    Aria Operations FQDN, hostname, or base URL.

.PARAMETER CsvPath
    Extended condition-matrix CSV. Defaults to config/alert-rules.csv in the
    repository root.

.PARAMETER Credential
    Aria Operations credential. If omitted, the script prompts securely.

.PARAMETER CollectionIntervalSeconds
    Collection interval used to resolve COLLECTION_RATIO condition values.
    Defaults to 300 seconds (five minutes).

.PARAMETER IncludeRuleId
    Optional RuleId filter. The complete CSV is still validated locally, but
    only matching rules are preflighted and processed.

.PARAMETER ExistingDefinitionAction
    Update (default) converges existing same-name definitions on the CSV. Skip
    reuses them unchanged. Fail rejects a same-name definition.

.PARAMETER NamePrefix
    Prefix applied to every generated alert definition name. If omitted, the
    script prompts for a non-empty value before processing any rules.

.PARAMETER PolicyId
    Optional policy identifiers. Alerts whose EnableInPolicy column is True are
    enabled in these policies after they are saved. With no PolicyId, policy
    state is left unchanged.

.PARAMETER SkipCertificateCheck
    Allows an untrusted HTTPS certificate. Intended for lab environments.

.PARAMETER SkipMetricValidation
    Skips live metric/property metadata validation. Resource kinds and alert
    taxonomy are still validated.

.PARAMETER ValidateOnly
    Performs local validation and live API preflight without changing Aria
    Operations.

.EXAMPLE
    $cred = Get-Credential
    .\New-AriaOpsAlertsFromCsv-Extended.ps1 -Server ariaops.example.com `
        -Credential $cred -SkipCertificateCheck -ValidateOnly

.EXAMPLE
    .\New-AriaOpsAlertsFromCsv-Extended.ps1 -Server ariaops.example.com `
        -SkipCertificateCheck -PolicyId 'policy-guid'

.NOTES
    Author: Drew Mackay
    Requires PowerShell 7 or later. Use -WhatIf to preview API writes after live
    validation. No alert is enabled in a policy unless -PolicyId is supplied.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath = (Join-Path -Path $PSScriptRoot -ChildPath '../config/alert-rules.csv'),

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [string]$ResultsPath,

    [Parameter()]
    [ValidateSet('Update', 'Skip', 'Fail')]
    [string]$ExistingDefinitionAction = 'Update',

    [Parameter()]
    [AllowEmptyString()]
    [string]$NamePrefix,

    [Parameter()]
    [ValidateRange(1, 86400)]
    [int]$CollectionIntervalSeconds = 300,

    [Parameter()]
    [string[]]$IncludeRuleId,

    [Parameter()]
    [string[]]$PolicyId,

    [Parameter()]
    [switch]$SkipCertificateCheck,

    [Parameter()]
    [switch]$SkipMetricValidation,

    [Parameter()]
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($NamePrefix)) {
    $NamePrefix = Read-Host 'Enter the alert definition name prefix (for example, ABC.inc Alerts)'
}
if ([string]::IsNullOrWhiteSpace($NamePrefix)) {
    throw 'NamePrefix is required. Supply -NamePrefix or enter a non-empty value at the prompt.'
}
$NamePrefix = $NamePrefix.Trim()

$script:ApiHeaders = $null
$script:ApiBaseUri = $null
$script:SkipCertificateCheck = [bool]$SkipCertificateCheck
$script:CallerPSCmdlet = $PSCmdlet
$script:AlertCache = @{}
$script:AdapterCache = @{}
$script:StatMetadataCache = @{}
$script:PropertyMetadataCache = @{}

function ConvertTo-EncodedValue {
    param([Parameter(Mandatory = $true)][string]$Value)

    return [System.Uri]::EscapeDataString($Value)
}

function Get-OptionalCsvValue {
    param(
        [Parameter(Mandatory = $true)][psobject]$Row,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        return $null
    }

    return ([string]$property.Value).Trim()
}

function ConvertTo-CsvBoolean {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string]$ColumnName
    )

    $parsed = $false
    if ([bool]::TryParse($Value.Trim(), [ref]$parsed)) { return $parsed }
    if ($Value.Trim() -eq '1') { return $true }
    if ($Value.Trim() -eq '0') { return $false }
    throw "$ColumnName '$Value' must be True, False, 1, or 0."
}

function ConvertTo-PositiveCsvInteger {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$ColumnName
    )

    $parsed = 0
    if (-not [int]::TryParse($Value.Trim(), [ref]$parsed) -or $parsed -lt 1) {
        throw "$ColumnName '$Value' must be a positive integer."
    }
    return $parsed
}

function ConvertTo-InvariantCsvNumber {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$ColumnName
    )

    $parsed = 0.0
    $style = [System.Globalization.NumberStyles]::Float
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if (-not [double]::TryParse($Value.Trim(), $style, $culture, [ref]$parsed)) {
        throw "$ColumnName '$Value' must be an invariant-culture number."
    }
    return $parsed
}

function Format-InvariantNumber {
    param([Parameter(Mandatory = $true)][double]$Value)

    return $Value.ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-ApiErrorMessage {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $messages = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($ErrorRecord.Exception.Message)) {
        $messages.Add($ErrorRecord.Exception.Message)
    }
    if ($null -ne $ErrorRecord.ErrorDetails -and
        -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        $messages.Add($ErrorRecord.ErrorDetails.Message)
    }

    return ($messages | Select-Object -Unique) -join ' | '
}

function Invoke-AriaOpsApi {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Get', 'Post', 'Put', 'Delete')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter()]
        [AllowNull()]
        [object]$Body
    )

    $invokeParameters = @{
        Uri         = $Uri
        Method      = $Method
        Headers     = $script:ApiHeaders
        ErrorAction = 'Stop'
    }
    if ($script:SkipCertificateCheck) {
        $invokeParameters.SkipCertificateCheck = $true
    }
    if ($PSBoundParameters.ContainsKey('Body')) {
        $invokeParameters.Body = $Body | ConvertTo-Json -Depth 30
        $invokeParameters.ContentType = 'application/json'
    }

    try {
        return Invoke-RestMethod @invokeParameters
    }
    catch {
        $safePath = try { ([uri]$Uri).PathAndQuery } catch { $Uri }
        $detail = Get-ApiErrorMessage -ErrorRecord $_
        throw "Aria Operations API $Method $safePath failed: $detail"
    }
}

function Get-NormalizedServerUri {
    param([Parameter(Mandatory = $true)][string]$Value)

    $candidate = $Value.Trim().TrimEnd('/')
    if ($candidate -notmatch '^https?://') {
        $candidate = "https://$candidate"
    }
    $candidate = $candidate -replace '/suite-api/api$', ''

    $parsed = $null
    if (-not [uri]::TryCreate($candidate, [UriKind]::Absolute, [ref]$parsed)) {
        throw "Server '$Value' is not a valid hostname or absolute HTTP(S) URL."
    }
    if ($parsed.Scheme -notin @('http', 'https')) {
        throw "Server '$Value' must use HTTP or HTTPS."
    }

    return $candidate
}

function Get-ConsistentRuleValue {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)][string]$RuleId
    )

    $values = @($Rows | ForEach-Object { [string]$_.$PropertyName } | Select-Object -Unique)
    if ($values.Count -ne 1) {
        throw "Rule '$RuleId' has inconsistent $PropertyName values: $($values -join ', ')."
    }
    return $Rows[0].$PropertyName
}

function Get-AdapterMetadata {
    param([Parameter(Mandatory = $true)][string]$AdapterKindKey)

    if (-not $script:AdapterCache.ContainsKey($AdapterKindKey)) {
        $encodedAdapter = ConvertTo-EncodedValue -Value $AdapterKindKey
        $script:AdapterCache[$AdapterKindKey] = Invoke-AriaOpsApi -Method Get `
            -Uri "$script:ApiBaseUri/adapterkinds/$encodedAdapter"
    }
    return $script:AdapterCache[$AdapterKindKey]
}

function Assert-ResourceKindExists {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey
    )

    $adapter = Get-AdapterMetadata -AdapterKindKey $AdapterKindKey
    if ($ResourceKindKey -cnotin @($adapter.resourceKinds)) {
        throw "Resource kind '$AdapterKindKey/$ResourceKindKey' does not exist on the target server."
    }
}

function Get-StatMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey,
        [Parameter(Mandatory = $true)][string]$Key
    )

    $cacheKey = "$AdapterKindKey`n$ResourceKindKey`n$Key"
    if (-not $script:StatMetadataCache.ContainsKey($cacheKey)) {
        $adapter = ConvertTo-EncodedValue -Value $AdapterKindKey
        $resource = ConvertTo-EncodedValue -Value $ResourceKindKey
        $stat = ConvertTo-EncodedValue -Value $Key
        $uri = "$script:ApiBaseUri/adapterkinds/$adapter/resourcekinds/$resource/statkeys" +
            "?attributeKeys=$stat&retrieveAttributeKeys=true"
        $response = Invoke-AriaOpsApi -Method Get -Uri $uri
        $matches = @($response.resourceTypeAttributes | Where-Object { $_.key -ceq $Key })
        $script:StatMetadataCache[$cacheKey] = if ($matches.Count -eq 1) { $matches[0] } else { $null }
    }
    return $script:StatMetadataCache[$cacheKey]
}

function Get-PropertyMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey,
        [Parameter(Mandatory = $true)][string]$Key
    )

    $resourceCacheKey = "$AdapterKindKey`n$ResourceKindKey"
    if (-not $script:PropertyMetadataCache.ContainsKey($resourceCacheKey)) {
        $adapter = ConvertTo-EncodedValue -Value $AdapterKindKey
        $resource = ConvertTo-EncodedValue -Value $ResourceKindKey
        $uri = "$script:ApiBaseUri/adapterkinds/$adapter/resourcekinds/$resource/properties"
        $response = Invoke-AriaOpsApi -Method Get -Uri $uri
        $script:PropertyMetadataCache[$resourceCacheKey] = @($response.resourceTypeAttributes)
    }
    $matches = @($script:PropertyMetadataCache[$resourceCacheKey] |
            Where-Object { $_.key -ceq $Key })
    if ($matches.Count -eq 1) { return $matches[0] }
    return $null
}

function Get-PagedAlertDefinitions {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $page = 0
    do {
        $adapter = ConvertTo-EncodedValue -Value $AdapterKindKey
        $resource = ConvertTo-EncodedValue -Value $ResourceKindKey
        $uri = "$script:ApiBaseUri/alertdefinitions?adapterKind=$adapter&resourceKind=$resource&page=$page&pageSize=1000"
        $response = Invoke-AriaOpsApi -Method Get -Uri $uri
        foreach ($item in @($response.alertDefinitions)) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $totalCount = if ($null -ne $response.pageInfo) { [int]$response.pageInfo.totalCount } else { $items.Count }
        $page++
    } while ($items.Count -lt $totalCount)

    return $items.ToArray()
}

function Get-CachedAlertDefinitions {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey
    )

    $cacheKey = "$AdapterKindKey`n$ResourceKindKey"
    if (-not $script:AlertCache.ContainsKey($cacheKey)) {
        $script:AlertCache[$cacheKey] = @(Get-PagedAlertDefinitions `
                -AdapterKindKey $AdapterKindKey -ResourceKindKey $ResourceKindKey)
    }
    return @($script:AlertCache[$cacheKey])
}

function Update-AlertCache {
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition
    )

    $cacheKey = "$($Definition.adapterKindKey)`n$($Definition.resourceKindKey)"
    $current = @(Get-CachedAlertDefinitions -AdapterKindKey ([string]$Definition.adapterKindKey) `
            -ResourceKindKey ([string]$Definition.resourceKindKey))
    $script:AlertCache[$cacheKey] = @($current | Where-Object { $_.id -ne $Definition.id }) + @($Definition)
}

function Save-AriaOpsAlertDefinition {
    param([Parameter(Mandatory = $true)][hashtable]$Desired)

    $matches = @(Get-CachedAlertDefinitions -AdapterKindKey ([string]$Desired.adapterKindKey) `
            -ResourceKindKey ([string]$Desired.resourceKindKey) |
            Where-Object { $_.name -ceq $Desired.name })
    if ($matches.Count -gt 1) {
        throw "More than one alert definition named '$($Desired.name)' exists for '$($Desired.resourceKindKey)'."
    }
    $existing = if ($matches.Count -eq 1) { $matches[0] } else { $null }

    if ($null -ne $existing) {
        switch ($ExistingDefinitionAction) {
            'Fail' { throw "Alert definition '$($Desired.name)' already exists (ID $($existing.id))." }
            'Skip' {
                return [pscustomobject]@{ Definition = $existing; Outcome = 'SKIPPED_EXISTING' }
            }
            'Update' {
                if ($ValidateOnly) {
                    return [pscustomobject]@{ Definition = $existing; Outcome = 'VALIDATED_UPDATE' }
                }
                if (-not $script:CallerPSCmdlet.ShouldProcess("alert definition '$($Desired.name)'", 'Update')) {
                    return [pscustomobject]@{ Definition = $existing; Outcome = 'WHATIF_UPDATE' }
                }
                $payload = [ordered]@{}
                foreach ($key in $Desired.Keys) { $payload[$key] = $Desired[$key] }
                $payload['id'] = $existing.id
                $saved = Invoke-AriaOpsApi -Method Put -Uri "$script:ApiBaseUri/alertdefinitions" -Body $payload
                if ([string]::IsNullOrWhiteSpace([string]$saved.id)) {
                    throw "The API updated '$($Desired.name)' but returned no alert definition ID."
                }
                Update-AlertCache -Definition $saved
                return [pscustomobject]@{ Definition = $saved; Outcome = 'UPDATED' }
            }
        }
    }

    if ($ValidateOnly) {
        return [pscustomobject]@{
            Definition = [pscustomobject]@{ id = "validate-only-$([guid]::NewGuid())"; name = $Desired.name }
            Outcome    = 'VALIDATED_CREATE'
        }
    }
    if (-not $script:CallerPSCmdlet.ShouldProcess("alert definition '$($Desired.name)'", 'Create')) {
        return [pscustomobject]@{
            Definition = [pscustomobject]@{ id = "whatif-$([guid]::NewGuid())"; name = $Desired.name }
            Outcome    = 'WHATIF_CREATE'
        }
    }

    $created = Invoke-AriaOpsApi -Method Post -Uri "$script:ApiBaseUri/alertdefinitions" -Body $Desired
    if ([string]::IsNullOrWhiteSpace([string]$created.id)) {
        throw "The API created '$($Desired.name)' but returned no alert definition ID."
    }
    Update-AlertCache -Definition $created
    return [pscustomobject]@{ Definition = $created; Outcome = 'CREATED' }
}

function Enable-AlertDefinition {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][bool]$EnabledByCsv
    )

    if (-not $EnabledByCsv) { return 'DISABLED_BY_CSV' }
    $policyIds = @($PolicyId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } | Select-Object -Unique)
    if ($policyIds.Count -eq 0) { return 'NOT_REQUESTED' }
    if ($ValidateOnly) { return 'VALIDATED' }
    if (-not $script:CallerPSCmdlet.ShouldProcess("alert definition '$Id'", "Enable in $($policyIds.Count) policy/policies")) {
        return 'WHATIF'
    }

    $encodedId = ConvertTo-EncodedValue -Value $Id
    $query = ($policyIds | ForEach-Object { 'policyId=' + (ConvertTo-EncodedValue -Value $_) }) -join '&'
    Invoke-AriaOpsApi -Method Put -Uri "$script:ApiBaseUri/alertdefinitions/$encodedId/enable?$query" | Out-Null
    return 'ENABLED'
}

function New-ConditionPayload {
    param([Parameter(Mandatory = $true)][psobject]$ConditionRow)

    switch ($ConditionRow.ConditionType) {
        'STATIC' {
            return [ordered]@{
                type          = 'CONDITION_HT'
                key           = $ConditionRow.ConditionKey
                operator      = $ConditionRow.Operator
                value         = Format-InvariantNumber -Value ([double]$ConditionRow.NumericValue)
                valueType     = 'NUMERIC'
                instanced     = [bool]$ConditionRow.Instanced
                thresholdType = 'STATIC'
            }
        }
        'COLLECTION_RATIO' {
            $computedValue = ([double]$CollectionIntervalSeconds * 1000.0) * [double]$ConditionRow.ValueMultiplier
            return [ordered]@{
                type          = 'CONDITION_HT'
                key           = $ConditionRow.ConditionKey
                operator      = $ConditionRow.Operator
                value         = Format-InvariantNumber -Value $computedValue
                valueType     = 'NUMERIC'
                instanced     = [bool]$ConditionRow.Instanced
                thresholdType = 'STATIC'
            }
        }
        'DYNAMIC' {
            return [ordered]@{
                type      = 'CONDITION_DT'
                key       = $ConditionRow.ConditionKey
                operator  = $ConditionRow.Operator
                instanced = [bool]$ConditionRow.Instanced
            }
        }
        'PROPERTY_NUMERIC' {
            return [ordered]@{
                type          = 'CONDITION_PROPERTY_NUMERIC'
                key           = $ConditionRow.ConditionKey
                operator      = $ConditionRow.Operator
                value         = [double]$ConditionRow.NumericValue
                instanced     = [bool]$ConditionRow.Instanced
                thresholdType = 'STATIC'
            }
        }
        default { throw "Unsupported ConditionType '$($ConditionRow.ConditionType)'." }
    }
}

function New-BranchSymptomSet {
    param(
        [Parameter(Mandatory = $true)][object[]]$BranchRows,
        [Parameter(Mandatory = $true)][psobject]$Rule
    )

    $simpleSets = @(
        foreach ($targetGroup in @($BranchRows | Group-Object {
                    "$($_.ConditionAdapterKindKey)`n$($_.ConditionResourceKindKey)`n$($_.Relation)`n$($_.Aggregation)"
                })) {
            $targetRows = @($targetGroup.Group)
            $alertConditions = @(
                foreach ($row in $targetRows) {
                    [ordered]@{
                        waitCycles   = [int]$row.ConditionWaitCycles
                        cancelCycles = [int]$row.ConditionCancelCycles
                        severity     = $row.StateSeverity
                        condition    = New-ConditionPayload -ConditionRow $row
                    }
                }
            )
            $set = [ordered]@{
                type                 = 'SYMPTOM_SET'
                relation             = $targetRows[0].Relation
                aggregation          = $targetRows[0].Aggregation
                symptomSetOperator   = 'AND'
                symptomDefinitionIds = @()
                alertConditions      = $alertConditions
            }
            if ($targetRows[0].Relation -ne 'SELF' -or
                $targetRows[0].ConditionAdapterKindKey -cne $Rule.AdapterKindKey -or
                $targetRows[0].ConditionResourceKindKey -cne $Rule.ResourceKindKey) {
                $set['adapterKindKey'] = $targetRows[0].ConditionAdapterKindKey
                $set['resourceKindKey'] = $targetRows[0].ConditionResourceKindKey
            }
            $set
        }
    )
    if ($simpleSets.Count -eq 1) { return $simpleSets[0] }

    return [ordered]@{
        type           = 'SYMPTOM_SET_COMPOSITE'
        operator       = 'AND'
        'symptom-sets' = $simpleSets
    }
}

function New-BaseSymptomSet {
    param(
        [Parameter(Mandatory = $true)][object[]]$ConditionRows,
        [Parameter(Mandatory = $true)][psobject]$Rule
    )

    $branchSets = @(
        foreach ($severityGroup in @($ConditionRows | Group-Object StateSeverity)) {
            foreach ($branchGroup in @($severityGroup.Group | Group-Object Branch)) {
                New-BranchSymptomSet -BranchRows @($branchGroup.Group) -Rule $Rule
            }
        }
    )
    if ($branchSets.Count -eq 1) { return $branchSets[0] }

    return [ordered]@{
        type           = 'SYMPTOM_SET_COMPOSITE'
        operator       = 'OR'
        'symptom-sets' = $branchSets
    }
}

function Get-PersistedAlertConditions {
    param(
        [Parameter(Mandatory = $true)][psobject]$SymptomSet,
        [Parameter(Mandatory = $true)][string]$RootAdapterKindKey,
        [Parameter(Mandatory = $true)][string]$RootResourceKindKey
    )

    if ($SymptomSet.type -eq 'SYMPTOM_SET_COMPOSITE') {
        foreach ($child in @($SymptomSet.'symptom-sets')) {
            Get-PersistedAlertConditions -SymptomSet $child `
                -RootAdapterKindKey $RootAdapterKindKey -RootResourceKindKey $RootResourceKindKey
        }
        return
    }

    $adapterProperty = $SymptomSet.PSObject.Properties['adapterKindKey']
    $resourceProperty = $SymptomSet.PSObject.Properties['resourceKindKey']
    $adapter = if ($null -eq $adapterProperty -or [string]::IsNullOrWhiteSpace([string]$adapterProperty.Value)) {
        $RootAdapterKindKey
    } else { [string]$adapterProperty.Value }
    $resource = if ($null -eq $resourceProperty -or [string]::IsNullOrWhiteSpace([string]$resourceProperty.Value)) {
        $RootResourceKindKey
    } else { [string]$resourceProperty.Value }
    foreach ($condition in @($SymptomSet.alertConditions)) {
        [pscustomobject]@{
            AdapterKindKey  = $adapter
            ResourceKindKey = $resource
            Relation        = [string]$SymptomSet.relation
            Severity        = [string]$condition.severity
            WaitCycles      = [int]$condition.waitCycles
            CancelCycles    = [int]$condition.cancelCycles
            Condition       = $condition.condition
        }
    }
}

function Get-ConditionSignature {
    param(
        [Parameter(Mandatory = $true)][string]$AdapterKindKey,
        [Parameter(Mandatory = $true)][string]$ResourceKindKey,
        [Parameter(Mandatory = $true)][string]$Relation,
        [Parameter(Mandatory = $true)][string]$Severity,
        [Parameter(Mandatory = $true)][int]$WaitCycles,
        [Parameter(Mandatory = $true)][int]$CancelCycles,
        [Parameter(Mandatory = $true)][psobject]$Condition
    )

    $value = ''
    if ($null -ne $Condition.PSObject.Properties['value'] -and $null -ne $Condition.value) {
        $number = 0.0
        if ([double]::TryParse([string]$Condition.value,
                [System.Globalization.NumberStyles]::Float,
                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            $value = Format-InvariantNumber -Value $number
        }
        else { $value = [string]$Condition.value }
    }
    return @(
        $AdapterKindKey, $ResourceKindKey, $Relation, $Severity,
        [string]$WaitCycles, [string]$CancelCycles,
        [string]$Condition.type, [string]$Condition.key,
        [string]$Condition.operator, $value, [string][bool]$Condition.instanced
    ) -join '|'
}

function Assert-PersistedConditions {
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [Parameter(Mandatory = $true)][object[]]$ExpectedRows
    )

    $actualRecords = @(
        foreach ($state in @($Definition.states)) {
            Get-PersistedAlertConditions -SymptomSet $state.'base-symptom-set' `
                -RootAdapterKindKey ([string]$Definition.adapterKindKey) `
                -RootResourceKindKey ([string]$Definition.resourceKindKey)
        }
    )
    $expectedSignatures = @(
        foreach ($row in $ExpectedRows) {
            $condition = [pscustomobject](New-ConditionPayload -ConditionRow $row)
            Get-ConditionSignature -AdapterKindKey $row.ConditionAdapterKindKey `
                -ResourceKindKey $row.ConditionResourceKindKey -Relation $row.Relation `
                -Severity $row.StateSeverity -WaitCycles $row.ConditionWaitCycles `
                -CancelCycles $row.ConditionCancelCycles -Condition $condition
        }
    )
    $actualSignatures = @(
        foreach ($record in $actualRecords) {
            Get-ConditionSignature -AdapterKindKey $record.AdapterKindKey `
                -ResourceKindKey $record.ResourceKindKey -Relation $record.Relation `
                -Severity $record.Severity -WaitCycles $record.WaitCycles `
                -CancelCycles $record.CancelCycles -Condition $record.Condition
        }
    )

    if ($actualSignatures.Count -ne $expectedSignatures.Count) {
        throw "Saved alert '$($Definition.name)' contains $($actualSignatures.Count) conditions; expected $($expectedSignatures.Count)."
    }
    $expectedGroups = @($expectedSignatures | Group-Object | Sort-Object Name)
    $actualGroups = @($actualSignatures | Group-Object | Sort-Object Name)
    if (($expectedGroups | ForEach-Object { "$($_.Count)x$($_.Name)" }) -join "`n" -cne
        (($actualGroups | ForEach-Object { "$($_.Count)x$($_.Name)" }) -join "`n")) {
        throw "Saved alert '$($Definition.name)' does not preserve the CSV condition matrix."
    }
}

function Export-Results {
    param(
        [Parameter(Mandatory = $true)][object[]]$InputObject,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force -WhatIf:$false | Out-Null
    }
    $InputObject | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 -WhatIf:$false
}

# Resolve and validate the complete CSV before opening an API session.
$resolvedCsvPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CsvPath)
if (-not (Test-Path -LiteralPath $resolvedCsvPath -PathType Leaf)) {
    throw "CSV file not found: $resolvedCsvPath"
}
if ([string]::IsNullOrWhiteSpace($ResultsPath)) {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $ResultsPath = Join-Path -Path $PSScriptRoot -ChildPath "../output/extended-alert-results-$timestamp.csv"
}
$resolvedResultsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsPath)
if ($resolvedResultsPath -eq $resolvedCsvPath) {
    throw 'ResultsPath must not overwrite CsvPath.'
}

$rawRows = @(Import-Csv -LiteralPath $resolvedCsvPath)
if ($rawRows.Count -eq 0) { throw "CSV '$resolvedCsvPath' contains no data rows." }

$requiredColumns = @(
    'RuleId', 'SourceSheet', 'SourceRow', 'DeploymentMode', 'EnableInPolicy', 'Fidelity', 'RuleName',
    'Domain', 'ObjectType', 'AdapterKindKey', 'ResourceKindKey', 'MetricName', 'MetricKey', 'ThresholdModel',
    'AlertType', 'AlertSubType', 'ImpactType', 'ImpactDetail', 'AlertWaitCycles', 'AlertCancelCycles',
    'StateSeverity', 'Branch', 'ConditionId', 'ConditionName', 'ConditionAdapterKindKey', 'ConditionResourceKindKey',
    'Relation', 'Aggregation', 'ConditionType', 'ConditionKey', 'Operator', 'ThresholdType', 'Value',
    'ValueMultiplier', 'ValueType', 'Instanced', 'ConditionWaitCycles', 'ConditionCancelCycles', 'SourceDecision',
    'SourceAlertMonitor', 'SourceSignal', 'SourceGreen', 'SourceAmber', 'SourceRed', 'Correlation',
    'ImplementationNotes', 'Source', 'Confidence'
)
$actualColumns = @($rawRows[0].PSObject.Properties.Name)
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })
if ($missingColumns.Count -gt 0) {
    throw "CSV is missing required column(s): $($missingColumns -join ', ')"
}

$commonRequiredValues = @(
    'RuleId', 'SourceSheet', 'SourceRow', 'DeploymentMode', 'EnableInPolicy', 'Fidelity', 'RuleName',
    'Domain', 'ObjectType', 'AdapterKindKey', 'ResourceKindKey', 'MetricName', 'MetricKey', 'ThresholdModel',
    'AlertType', 'AlertSubType', 'ImpactType', 'ImpactDetail', 'AlertWaitCycles', 'AlertCancelCycles'
)
$conditionRequiredValues = @(
    'StateSeverity', 'Branch', 'ConditionId', 'ConditionName', 'ConditionAdapterKindKey',
    'ConditionResourceKindKey', 'Relation', 'Aggregation', 'ConditionType', 'ConditionKey',
    'Operator', 'ValueType', 'Instanced', 'ConditionWaitCycles', 'ConditionCancelCycles'
)
$staticOperators = @('GT', 'GT_EQ', 'LT', 'LT_EQ', 'EQ', 'NOT_EQ')
$dynamicOperators = @('DT_ABOVE', 'DT_BELOW', 'DT_ABNORMAL')
$allowedRelations = @('SELF', 'PARENT', 'CHILD', 'ANCESTOR', 'DESCENDANT')
$allowedAggregations = @('ALL', 'ANY')
$conditionIdentitySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$normalizedRows = [System.Collections.Generic.List[object]]::new()

for ($index = 0; $index -lt $rawRows.Count; $index++) {
    $row = $rawRows[$index]
    $csvLine = $index + 2
    try {
        foreach ($column in $commonRequiredValues) {
            if ([string]::IsNullOrWhiteSpace([string]$row.$column)) {
                throw "Required value '$column' is empty."
            }
        }
        $mode = ([string]$row.DeploymentMode).Trim().ToUpperInvariant()
        if ($mode -notin @('ALERT', 'MONITOR_ONLY')) {
            throw "DeploymentMode '$mode' must be ALERT or MONITOR_ONLY."
        }
        $sourceRow = ConvertTo-PositiveCsvInteger -Value ([string]$row.SourceRow) -ColumnName SourceRow
        $alertType = ConvertTo-PositiveCsvInteger -Value ([string]$row.AlertType) -ColumnName AlertType
        $alertSubType = ConvertTo-PositiveCsvInteger -Value ([string]$row.AlertSubType) -ColumnName AlertSubType
        $alertWaitCycles = ConvertTo-PositiveCsvInteger -Value ([string]$row.AlertWaitCycles) -ColumnName AlertWaitCycles
        $alertCancelCycles = ConvertTo-PositiveCsvInteger -Value ([string]$row.AlertCancelCycles) -ColumnName AlertCancelCycles
        $enableInPolicy = ConvertTo-CsvBoolean -Value ([string]$row.EnableInPolicy) -ColumnName EnableInPolicy

        $normalized = [ordered]@{
            CsvLine                   = $csvLine
            RuleId                    = ([string]$row.RuleId).Trim()
            SourceSheet               = ([string]$row.SourceSheet).Trim()
            SourceRow                 = $sourceRow
            DeploymentMode            = $mode
            EnableInPolicy            = $enableInPolicy
            Fidelity                  = ([string]$row.Fidelity).Trim()
            RuleName                  = ([string]$row.RuleName).Trim()
            Domain                    = ([string]$row.Domain).Trim()
            ObjectType                = ([string]$row.ObjectType).Trim()
            AdapterKindKey            = ([string]$row.AdapterKindKey).Trim()
            ResourceKindKey           = ([string]$row.ResourceKindKey).Trim()
            MetricName                = ([string]$row.MetricName).Trim()
            MetricKey                 = ([string]$row.MetricKey).Trim()
            ThresholdModel            = ([string]$row.ThresholdModel).Trim()
            AlertType                 = $alertType
            AlertSubType              = $alertSubType
            ImpactType                = ([string]$row.ImpactType).Trim()
            ImpactDetail              = ([string]$row.ImpactDetail).Trim()
            AlertWaitCycles           = $alertWaitCycles
            AlertCancelCycles         = $alertCancelCycles
            StateSeverity             = $null
            Branch                    = $null
            ConditionId               = $null
            ConditionName             = $null
            ConditionAdapterKindKey   = $null
            ConditionResourceKindKey  = $null
            Relation                  = $null
            Aggregation               = $null
            ConditionType             = $null
            ConditionKey              = $null
            Operator                  = $null
            ThresholdType             = $null
            NumericValue              = $null
            ValueMultiplier           = $null
            ValueType                 = $null
            Instanced                 = $false
            ConditionWaitCycles       = $null
            ConditionCancelCycles     = $null
            SourceDecision            = Get-OptionalCsvValue -Row $row -Name SourceDecision
            SourceAlertMonitor        = Get-OptionalCsvValue -Row $row -Name SourceAlertMonitor
            SourceSignal              = Get-OptionalCsvValue -Row $row -Name SourceSignal
            SourceGreen               = Get-OptionalCsvValue -Row $row -Name SourceGreen
            SourceAmber               = Get-OptionalCsvValue -Row $row -Name SourceAmber
            SourceRed                 = Get-OptionalCsvValue -Row $row -Name SourceRed
            Correlation               = Get-OptionalCsvValue -Row $row -Name Correlation
            ImplementationNotes       = Get-OptionalCsvValue -Row $row -Name ImplementationNotes
            Source                    = Get-OptionalCsvValue -Row $row -Name Source
            Confidence                = Get-OptionalCsvValue -Row $row -Name Confidence
        }

        if ($mode -eq 'ALERT') {
            foreach ($column in $conditionRequiredValues) {
                if ([string]::IsNullOrWhiteSpace([string]$row.$column)) {
                    throw "Required ALERT value '$column' is empty."
                }
            }
            $severity = ([string]$row.StateSeverity).Trim().ToUpperInvariant()
            if ($severity -notin @('WARNING', 'CRITICAL')) {
                throw "StateSeverity '$severity' must be WARNING or CRITICAL."
            }
            $branch = ConvertTo-PositiveCsvInteger -Value ([string]$row.Branch) -ColumnName Branch
            $conditionType = ([string]$row.ConditionType).Trim().ToUpperInvariant()
            if ($conditionType -notin @('STATIC', 'DYNAMIC', 'PROPERTY_NUMERIC', 'COLLECTION_RATIO')) {
                throw "Unsupported ConditionType '$conditionType'."
            }
            $operator = ([string]$row.Operator).Trim().ToUpperInvariant()
            if ($conditionType -eq 'DYNAMIC') {
                if ($operator -notin $dynamicOperators) { throw "Dynamic operator '$operator' is unsupported." }
            }
            elseif ($operator -notin $staticOperators) {
                throw "Operator '$operator' is unsupported for $conditionType."
            }
            $relation = ([string]$row.Relation).Trim().ToUpperInvariant()
            if ($relation -notin $allowedRelations) { throw "Relation '$relation' is unsupported." }
            $aggregation = ([string]$row.Aggregation).Trim().ToUpperInvariant()
            if ($aggregation -notin $allowedAggregations) { throw "Aggregation '$aggregation' is unsupported." }

            $numericValue = $null
            $valueMultiplier = $null
            if ($conditionType -in @('STATIC', 'PROPERTY_NUMERIC')) {
                $valueText = Get-OptionalCsvValue -Row $row -Name Value
                if ($null -eq $valueText) { throw "Value is required for $conditionType." }
                $numericValue = ConvertTo-InvariantCsvNumber -Value $valueText -ColumnName Value
            }
            elseif ($conditionType -eq 'COLLECTION_RATIO') {
                $multiplierText = Get-OptionalCsvValue -Row $row -Name ValueMultiplier
                if ($null -eq $multiplierText) { throw 'ValueMultiplier is required for COLLECTION_RATIO.' }
                $valueMultiplier = ConvertTo-InvariantCsvNumber -Value $multiplierText -ColumnName ValueMultiplier
                if ($valueMultiplier -le 0) { throw 'ValueMultiplier must be greater than zero.' }
            }
            $instanced = ConvertTo-CsvBoolean -Value ([string]$row.Instanced) -ColumnName Instanced
            $conditionWaitCycles = ConvertTo-PositiveCsvInteger `
                -Value ([string]$row.ConditionWaitCycles) -ColumnName ConditionWaitCycles
            $conditionCancelCycles = ConvertTo-PositiveCsvInteger `
                -Value ([string]$row.ConditionCancelCycles) -ColumnName ConditionCancelCycles
            $identity = "$($normalized.RuleId)`n$severity`n$branch`n$(([string]$row.ConditionId).Trim())"
            if (-not $conditionIdentitySet.Add($identity)) {
                throw "Duplicate condition identity '$identity'."
            }

            $normalized.StateSeverity = $severity
            $normalized.Branch = $branch
            $normalized.ConditionId = ([string]$row.ConditionId).Trim()
            $normalized.ConditionName = ([string]$row.ConditionName).Trim()
            $normalized.ConditionAdapterKindKey = ([string]$row.ConditionAdapterKindKey).Trim()
            $normalized.ConditionResourceKindKey = ([string]$row.ConditionResourceKindKey).Trim()
            $normalized.Relation = $relation
            $normalized.Aggregation = $aggregation
            $normalized.ConditionType = $conditionType
            $normalized.ConditionKey = ([string]$row.ConditionKey).Trim()
            $normalized.Operator = $operator
            $normalized.ThresholdType = Get-OptionalCsvValue -Row $row -Name ThresholdType
            $normalized.NumericValue = $numericValue
            $normalized.ValueMultiplier = $valueMultiplier
            $normalized.ValueType = ([string]$row.ValueType).Trim().ToUpperInvariant()
            $normalized.Instanced = $instanced
            $normalized.ConditionWaitCycles = $conditionWaitCycles
            $normalized.ConditionCancelCycles = $conditionCancelCycles
        }
        else {
            $conditionValues = @($conditionRequiredValues | ForEach-Object { Get-OptionalCsvValue -Row $row -Name $_ } |
                    Where-Object { $null -ne $_ })
            if ($conditionValues.Count -gt 0) {
                throw 'MONITOR_ONLY rows must not contain condition values.'
            }
            if ($enableInPolicy) {
                throw 'MONITOR_ONLY rows must set EnableInPolicy to False.'
            }
        }

        $normalizedRows.Add([pscustomobject]$normalized)
    }
    catch {
        throw "CSV validation failed at line ${csvLine}: $($_.Exception.Message)"
    }
}

$rules = [System.Collections.Generic.List[object]]::new()
foreach ($group in @($normalizedRows | Group-Object RuleId)) {
    $rows = @($group.Group)
    $ruleId = [string]$group.Name
    $consistentProperties = @(
        'SourceSheet', 'SourceRow', 'DeploymentMode', 'EnableInPolicy', 'Fidelity', 'RuleName',
        'Domain', 'ObjectType', 'AdapterKindKey', 'ResourceKindKey', 'MetricName', 'MetricKey',
        'ThresholdModel', 'AlertType', 'AlertSubType', 'ImpactType', 'ImpactDetail',
        'AlertWaitCycles', 'AlertCancelCycles'
    )
    foreach ($property in $consistentProperties) {
        Get-ConsistentRuleValue -Rows $rows -PropertyName $property -RuleId $ruleId | Out-Null
    }
    $rule = [pscustomobject]@{
        RuleId            = $ruleId
        SourceSheet       = $rows[0].SourceSheet
        SourceRow         = $rows[0].SourceRow
        DeploymentMode    = $rows[0].DeploymentMode
        EnableInPolicy    = $rows[0].EnableInPolicy
        Fidelity          = $rows[0].Fidelity
        RuleName          = $rows[0].RuleName
        Domain            = $rows[0].Domain
        ObjectType        = $rows[0].ObjectType
        AdapterKindKey    = $rows[0].AdapterKindKey
        ResourceKindKey   = $rows[0].ResourceKindKey
        MetricName        = $rows[0].MetricName
        MetricKey         = $rows[0].MetricKey
        ThresholdModel    = $rows[0].ThresholdModel
        AlertType         = $rows[0].AlertType
        AlertSubType      = $rows[0].AlertSubType
        ImpactType        = $rows[0].ImpactType
        ImpactDetail      = $rows[0].ImpactDetail
        AlertWaitCycles   = $rows[0].AlertWaitCycles
        AlertCancelCycles = $rows[0].AlertCancelCycles
        Conditions        = if ($rows[0].DeploymentMode -eq 'ALERT') { $rows } else { @() }
        Source            = $rows[0].Source
        ImplementationNotes = $rows[0].ImplementationNotes
    }
    if ($rule.DeploymentMode -eq 'ALERT') {
        $severities = @($rule.Conditions.StateSeverity | Select-Object -Unique)
        foreach ($requiredSeverity in @('WARNING', 'CRITICAL')) {
            if ($requiredSeverity -notin $severities) {
                throw "Rule '$ruleId' has no $requiredSeverity condition."
            }
        }
    }
    $rules.Add($rule)
}

if ($null -ne $IncludeRuleId -and @($IncludeRuleId).Count -gt 0) {
    $requestedRuleIds = @($IncludeRuleId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } | Select-Object -Unique)
    $knownRuleIds = @($rules.RuleId)
    $unknownRuleIds = @($requestedRuleIds | Where-Object { $_ -notin $knownRuleIds })
    if ($unknownRuleIds.Count -gt 0) {
        throw "IncludeRuleId contains unknown rule(s): $($unknownRuleIds -join ', ')."
    }
    $rules = @($rules | Where-Object { $_.RuleId -in $requestedRuleIds })
    $normalizedRows = @($normalizedRows | Where-Object { $_.RuleId -in $requestedRuleIds })
}

$serverRoot = Get-NormalizedServerUri -Value $Server
$script:ApiBaseUri = "$serverRoot/suite-api/api"
if ($null -eq $Credential) {
    $Credential = Get-Credential -Message "Aria Operations credentials for $serverRoot"
}
if ($null -eq $Credential) { throw 'No Aria Operations credential was supplied.' }

$tokenAcquired = $false
$results = [System.Collections.Generic.List[object]]::new()
try {
    $authBody = @{
        username = $Credential.UserName
        password = $Credential.GetNetworkCredential().Password
    }
    $authParameters = @{
        Uri         = "$script:ApiBaseUri/auth/token/acquire"
        Method      = 'Post'
        Body        = ($authBody | ConvertTo-Json)
        ContentType = 'application/json'
        Headers     = @{ Accept = 'application/json' }
        ErrorAction = 'Stop'
    }
    if ($script:SkipCertificateCheck) { $authParameters.SkipCertificateCheck = $true }
    try {
        $authResponse = Invoke-RestMethod @authParameters
    }
    catch {
        $detail = Get-ApiErrorMessage -ErrorRecord $_
        throw "Authentication to '$serverRoot' failed: $detail"
    }
    finally {
        $authBody = $null
        $authParameters.Body = $null
    }
    if ([string]::IsNullOrWhiteSpace([string]$authResponse.token)) {
        throw "Authentication to '$serverRoot' returned no token."
    }

    $script:ApiHeaders = @{
        Authorization = "vRealizeOpsToken $($authResponse.token)"
        Accept        = 'application/json'
    }
    $tokenAcquired = $true
    Write-Host "Authenticated to $serverRoot." -ForegroundColor Green

    # Complete live preflight before the first write.
    foreach ($identity in @(
            $rules | ForEach-Object { "$($_.AdapterKindKey)`n$($_.ResourceKindKey)" }
            $normalizedRows | Where-Object DeploymentMode -eq 'ALERT' |
                ForEach-Object { "$($_.ConditionAdapterKindKey)`n$($_.ConditionResourceKindKey)" }
        ) | Select-Object -Unique) {
        $parts = $identity -split "`n", 2
        Assert-ResourceKindExists -AdapterKindKey $parts[0] -ResourceKindKey $parts[1]
    }

    if (-not $SkipMetricValidation) {
        foreach ($rule in $rules) {
            if ($null -eq (Get-StatMetadata -AdapterKindKey $rule.AdapterKindKey `
                        -ResourceKindKey $rule.ResourceKindKey -Key $rule.MetricKey)) {
                throw "Source metric '$($rule.MetricKey)' is not defined for '$($rule.AdapterKindKey)/$($rule.ResourceKindKey)' (rule $($rule.RuleId))."
            }
        }
        foreach ($row in @($normalizedRows | Where-Object DeploymentMode -eq 'ALERT')) {
            if ($row.ConditionType -eq 'PROPERTY_NUMERIC') {
                $metadata = Get-PropertyMetadata -AdapterKindKey $row.ConditionAdapterKindKey `
                    -ResourceKindKey $row.ConditionResourceKindKey -Key $row.ConditionKey
                if ($null -eq $metadata) {
                    throw "Property '$($row.ConditionKey)' is not defined for '$($row.ConditionAdapterKindKey)/$($row.ConditionResourceKindKey)' (CSV line $($row.CsvLine))."
                }
                if ([string]$metadata.dataType2 -notmatch '^(INTEGER|LONG|FLOAT|DOUBLE|NUMBER)$') {
                    throw "Property '$($row.ConditionKey)' is not numeric (CSV line $($row.CsvLine))."
                }
            }
            elseif ($null -eq (Get-StatMetadata -AdapterKindKey $row.ConditionAdapterKindKey `
                        -ResourceKindKey $row.ConditionResourceKindKey -Key $row.ConditionKey)) {
                throw "Metric '$($row.ConditionKey)' is not defined for '$($row.ConditionAdapterKindKey)/$($row.ConditionResourceKindKey)' (CSV line $($row.CsvLine))."
            }
        }
    }

    $alertTypes = Invoke-AriaOpsApi -Method Get -Uri "$script:ApiBaseUri/alerts/types"
    foreach ($taxonomy in @($rules | Select-Object AlertType, AlertSubType -Unique)) {
        $type = @($alertTypes.alertTypes | Where-Object { [int]$_.id -eq [int]$taxonomy.AlertType })
        if ($type.Count -ne 1) { throw "Alert type '$($taxonomy.AlertType)' is unavailable on '$serverRoot'." }
        $subtype = @($type[0].subTypes | Where-Object { [int]$_.id -eq [int]$taxonomy.AlertSubType })
        if ($subtype.Count -ne 1) {
            throw "Alert subtype '$($taxonomy.AlertSubType)' is invalid for type '$($taxonomy.AlertType)' on '$serverRoot'."
        }
    }
    Write-Host "Preflight passed for $($rules.Count) rule(s) and $($normalizedRows.Count) CSV row(s)." -ForegroundColor Green

    foreach ($rule in $rules) {
        Write-Host "Processing $($rule.RuleId): $($rule.MetricName) [$($rule.ResourceKindKey)]" -ForegroundColor Cyan
        if ($rule.DeploymentMode -eq 'MONITOR_ONLY') {
            $results.Add([pscustomobject]@{
                    RuleId            = $rule.RuleId
                    SourceSheet       = $rule.SourceSheet
                    SourceRow         = $rule.SourceRow
                    DeploymentMode    = $rule.DeploymentMode
                    Fidelity          = $rule.Fidelity
                    ResourceKindKey   = $rule.ResourceKindKey
                    MetricKey         = $rule.MetricKey
                    Status            = if ($ValidateOnly) { 'VALIDATED_MONITOR_ONLY' } else { 'MONITOR_ONLY' }
                    AlertAction       = 'NOT_CREATED_BY_DESIGN'
                    PolicyAction      = 'NOT_APPLICABLE'
                    AlertDefinitionId = $null
                    ConditionCount    = 0
                    Error             = $null
                })
            continue
        }

        try {
            $alertName = "$NamePrefix - $($rule.RuleName)"
            $baseSet = New-BaseSymptomSet -ConditionRows @($rule.Conditions) -Rule $rule
            $description = "Generated from '$([IO.Path]::GetFileName($resolvedCsvPath))'; $($rule.RuleId); " +
                "$($rule.MetricKey); $($rule.ThresholdModel); fidelity=$($rule.Fidelity)."
            $payload = [ordered]@{
                name            = $alertName
                description     = $description
                adapterKindKey  = $rule.AdapterKindKey
                resourceKindKey = $rule.ResourceKindKey
                waitCycles      = [int]$rule.AlertWaitCycles
                cancelCycles    = [int]$rule.AlertCancelCycles
                type            = [int]$rule.AlertType
                subType         = [int]$rule.AlertSubType
                forVCDTenants   = $false
                states          = @(
                    [ordered]@{
                        severity           = 'AUTO'
                        'base-symptom-set' = $baseSet
                        impact             = [ordered]@{
                            impactType = $rule.ImpactType
                            detail     = $rule.ImpactDetail
                        }
                    }
                )
            }
            $save = Save-AriaOpsAlertDefinition -Desired $payload
            if ($save.Outcome -in @('CREATED', 'UPDATED', 'SKIPPED_EXISTING')) {
                Assert-PersistedConditions -Definition $save.Definition -ExpectedRows @($rule.Conditions)
            }
            $policyAction = Enable-AlertDefinition -Id ([string]$save.Definition.id) `
                -EnabledByCsv ([bool]$rule.EnableInPolicy
                )
            $status = if ($ValidateOnly) { 'VALIDATED' } elseif ($WhatIfPreference) { 'WHATIF' } else { 'SUCCEEDED' }
            $results.Add([pscustomobject]@{
                    RuleId            = $rule.RuleId
                    SourceSheet       = $rule.SourceSheet
                    SourceRow         = $rule.SourceRow
                    DeploymentMode    = $rule.DeploymentMode
                    Fidelity          = $rule.Fidelity
                    ResourceKindKey   = $rule.ResourceKindKey
                    MetricKey         = $rule.MetricKey
                    Status            = $status
                    AlertAction       = $save.Outcome
                    PolicyAction      = $policyAction
                    AlertDefinitionId = $save.Definition.id
                    ConditionCount    = @($rule.Conditions).Count
                    Error             = $null
                })
        }
        catch {
            $failureMessage = $_.Exception.Message
            if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
                $failureMessage += " | $($_.ScriptStackTrace -replace "`r?`n", ' <- ')"
            }
            Write-Warning "Rule '$($rule.RuleId)' failed: $failureMessage"
            $results.Add([pscustomobject]@{
                    RuleId            = $rule.RuleId
                    SourceSheet       = $rule.SourceSheet
                    SourceRow         = $rule.SourceRow
                    DeploymentMode    = $rule.DeploymentMode
                    Fidelity          = $rule.Fidelity
                    ResourceKindKey   = $rule.ResourceKindKey
                    MetricKey         = $rule.MetricKey
                    Status            = 'FAILED'
                    AlertAction       = $null
                    PolicyAction      = $null
                    AlertDefinitionId = $null
                    ConditionCount    = @($rule.Conditions).Count
                    Error             = $failureMessage
                })
        }
    }
}
finally {
    if ($tokenAcquired) {
        try {
            Invoke-AriaOpsApi -Method Post -Uri "$script:ApiBaseUri/auth/token/release" | Out-Null
        }
        catch {
            Write-Warning "Could not release the Aria Operations API token: $($_.Exception.Message)"
        }
        $script:ApiHeaders.Authorization = $null
    }
}

if ($results.Count -gt 0) {
    Export-Results -InputObject $results.ToArray() -Path $resolvedResultsPath
    Write-Host "Results written to $resolvedResultsPath" -ForegroundColor Green
    $results | Format-Table RuleId, DeploymentMode, Fidelity, ResourceKindKey, Status, AlertAction -AutoSize
}

$failedCount = @($results | Where-Object Status -eq 'FAILED').Count
if ($failedCount -gt 0) {
    throw "$failedCount of $($results.Count) rule(s) failed. Review '$resolvedResultsPath'."
}
