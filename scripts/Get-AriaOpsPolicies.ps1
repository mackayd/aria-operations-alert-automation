#Requires -Version 7.0

<#
.SYNOPSIS
    Lists Aria Operations policy names and GUIDs.

.DESCRIPTION
    Connects to the Aria Operations REST API, retrieves all policy summaries,
    and returns PowerShell objects containing the policy name, GUID, and
    default-policy status. No configuration is changed.

    The server name and credentials are supplied at runtime and are not stored
    in this script. If Credential is omitted, PowerShell prompts for it.

.PARAMETER Server
    Aria Operations FQDN, hostname, or base URL.

.PARAMETER Credential
    Aria Operations credential. If omitted, the script prompts securely.

.PARAMETER OutputPath
    Optional path at which to export the policy list as a CSV file.

.PARAMETER SkipCertificateCheck
    Allows an untrusted HTTPS certificate. Intended for lab environments.

.EXAMPLE
    $credential = Get-Credential
    .\Get-AriaOpsPolicies.ps1 -Server 'ariaops.example.com' `
        -Credential $credential -SkipCertificateCheck

.EXAMPLE
    .\Get-AriaOpsPolicies.ps1 -Server 'ariaops.example.com' `
        -Credential $credential -SkipCertificateCheck `
        -OutputPath '.\ariaops-policies.csv'

.NOTES
    Author: Drew Mackay
    Requires PowerShell 7 or later. This script performs read-only policy
    queries and releases its API token when finished.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$SkipCertificateCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Invoke-AriaOpsRequest {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Get', 'Post')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [AllowNull()]
        [object]$Body
    )

    $parameters = @{
        Method      = $Method
        Uri         = $Uri
        ErrorAction = 'Stop'
    }
    if ($null -ne $Headers) {
        $parameters.Headers = $Headers
    }
    if ($SkipCertificateCheck) {
        $parameters.SkipCertificateCheck = $true
    }
    if ($PSBoundParameters.ContainsKey('Body')) {
        $parameters.Body = $Body | ConvertTo-Json -Depth 5
        $parameters.ContentType = 'application/json'
    }

    try {
        return Invoke-RestMethod @parameters
    }
    catch {
        $safePath = try { ([uri]$Uri).PathAndQuery } catch { $Uri }
        $details = @(
            $_.Exception.Message
            if ($null -ne $_.ErrorDetails) { $_.ErrorDetails.Message }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        throw "Aria Operations API $Method $safePath failed: $($details -join ' | ')"
    }
}

$serverRoot = Get-NormalizedServerUri -Value $Server
$apiBaseUri = "$serverRoot/suite-api/api"
if ($null -eq $Credential) {
    $Credential = Get-Credential -Message "Enter credentials for $serverRoot"
}
if ($null -eq $Credential) {
    throw 'No credential was supplied.'
}

$apiToken = $null
$apiHeaders = $null

try {
    $authenticationBody = @{
        username = $Credential.UserName
        password = $Credential.GetNetworkCredential().Password
    }
    $authentication = Invoke-AriaOpsRequest -Method Post `
        -Uri "$apiBaseUri/auth/token/acquire" -Headers @{ Accept = 'application/json' } `
        -Body $authenticationBody
    $authenticationBody = $null

    $apiToken = [string]$authentication.token
    if ([string]::IsNullOrWhiteSpace($apiToken)) {
        throw 'Aria Operations returned an empty authentication token.'
    }
    $apiHeaders = @{
        Authorization = "vRealizeOpsToken $apiToken"
        Accept        = 'application/json'
    }

    $policyItems = [System.Collections.Generic.List[object]]::new()
    $page = 0
    $pageSize = 1000
    do {
        $response = Invoke-AriaOpsRequest -Method Get `
            -Uri "$apiBaseUri/policies?page=$page&pageSize=$pageSize" -Headers $apiHeaders
        $pageItems = @($response.policySummaries | Where-Object { $null -ne $_ })
        foreach ($policy in $pageItems) {
            $policyItems.Add($policy)
        }

        $pageInfoProperty = $response.PSObject.Properties['pageInfo']
        if ($null -eq $pageInfoProperty -or $null -eq $pageInfoProperty.Value) {
            $totalCount = $policyItems.Count
        }
        else {
            $totalCount = [int]$pageInfoProperty.Value.totalCount
        }
        $page++
    } while ($pageItems.Count -gt 0 -and $policyItems.Count -lt $totalCount)

    $policies = @(
        $policyItems |
            Sort-Object @{ Expression = { [bool]$_.defaultPolicy }; Descending = $true }, name |
            ForEach-Object {
                [pscustomobject]@{
                    Name          = [string]$_.name
                    PolicyGuid    = [string]$_.id
                    DefaultPolicy = [bool]$_.defaultPolicy
                }
            }
    )

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        $parentDirectory = Split-Path -Path $resolvedOutputPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($parentDirectory) -and
            -not (Test-Path -LiteralPath $parentDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $parentDirectory -Force | Out-Null
        }
        $policies | Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Policy list exported to '$resolvedOutputPath'."
    }

    Write-Output $policies
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($apiToken) -and $null -ne $apiHeaders) {
        try {
            Invoke-AriaOpsRequest -Method Post -Uri "$apiBaseUri/auth/token/release" `
                -Headers $apiHeaders | Out-Null
        }
        catch {
            Write-Warning "The API token could not be released: $($_.Exception.Message)"
        }
    }
}
