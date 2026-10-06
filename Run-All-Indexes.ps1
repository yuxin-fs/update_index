$ErrorActionPreference = "Stop"

# ------------------------
# Configuration
# ------------------------
$SearchService = "srch-ai-kmagent-dev-ccanada-001"
$ApiVersion = "2026-05-01-preview"

$IndexerPollSeconds = 30
$IndexerTimeoutMinutes = 180

$PythonExe = (Get-Command python).Source
$UpdateAclScript = "D:\Users\YuXin.Li\Desktop\research\update_index\update_acls_backup.py"
$LogDirectory = Join-Path $PSScriptRoot "logs"
$LogPath = Join-Path $LogDirectory (
    "Run-All-Indexes-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
)

# Replace with your actual prefixes
$Prefixes = @(
    "ai-sharepoint",
    "digitalguestexperience-sharepoint",
    "esg-sharepoint",
    "foodbeverage-sharepoint",
    "guestrecognition-sharepoint",
    "iconnect-sharepoint",
    "leadwithcare-sharepoint",
    "lnd-sharepoint",
    "pt-sharepoint",
    "spa-sharepoint"
)

# ------------------------
# Functions
# ------------------------

function Get-SearchToken {
    $token = az account get-access-token `
        --resource https://search.azure.com `
        --query accessToken `
        -o tsv

    if (-not $token) {
        throw "Failed to acquire Azure AI Search access token."
    }

    return $token
}

function Invoke-IndexerRun {
    param(
        [Parameter(Mandatory)]
        [string]$IndexerName,

        [Parameter(Mandatory)]
        [string]$Token
    )

    $url = "https://$SearchService.search.windows.net/indexers/$IndexerName/run?api-version=$ApiVersion"

    $headers = @{
        Authorization = "Bearer $Token"
        Accept        = "application/json"
    }

    # Record this immediately before starting the indexer.
    $requestedAt = [datetimeoffset]::UtcNow

    Write-Host (
        "Starting indexer: {0} RequestedAt={1:u}" -f
        $IndexerName,
        $requestedAt
    )

    try {
        $null = Invoke-RestMethod `
            -Method Post `
            -Uri $url `
            -Headers $headers `
            -TimeoutSec 60
    }
    catch {
        $statusCode = $null

        if ($_.Exception.Response) {
            try {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            catch {
                $statusCode = $null
            }
        }

        if ($statusCode -eq 401) {
            Write-Warning "Search token expired before indexer start. Refreshing token."

            $Token = Get-SearchToken
            $headers.Authorization = "Bearer $Token"

            $null = Invoke-RestMethod `
                -Method Post `
                -Uri $url `
                -Headers $headers `
                -TimeoutSec 60
        }
        else {
            throw
        }
    }

    return $requestedAt
}

function Wait-Indexer {
    param(
        [Parameter(Mandatory)]
        [string]$IndexerName,

        [Parameter(Mandatory)]
        [string]$Token,

        [Parameter(Mandatory)]
        [datetimeoffset]$RequestedAt
    )

    $url = "https://$SearchService.search.windows.net/indexers/$IndexerName/status?api-version=$ApiVersion"
    $deadline = [datetimeoffset]::UtcNow.AddMinutes($IndexerTimeoutMinutes)

    # Allow for minor clock precision differences between the VDI and Azure.
    $requestedAtTolerance = $RequestedAt.AddSeconds(-5)

    while ([datetimeoffset]::UtcNow -lt $deadline) {
        try {
            $status = Invoke-RestMethod `
                -Method Get `
                -Uri $url `
                -Headers @{
                    Authorization = "Bearer $Token"
                    Accept        = "application/json"
                } `
                -TimeoutSec 60
        }
        catch {
            $statusCode = $null

            if ($_.Exception.Response) {
                try {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                catch {
                    $statusCode = $null
                }
            }

            if ($statusCode -eq 401) {
                Write-Warning "Search token expired. Refreshing token."

                $Token = Get-SearchToken

                $status = Invoke-RestMethod `
                    -Method Get `
                    -Uri $url `
                    -Headers @{
                        Authorization = "Bearer $Token"
                        Accept        = "application/json"
                    } `
                    -TimeoutSec 60
            }
            else {
                throw
            }
        }

        $last = $status.lastResult

        if ($null -eq $last) {
            Write-Host "Indexer=$IndexerName has no lastResult yet. Waiting..."
            Start-Sleep -Seconds $IndexerPollSeconds
            continue
        }

        $startTime = $null
        $endTime = $null

        if ($last.startTime) {
            $startTime = [datetimeoffset]::Parse(
                [string]$last.startTime
            )
        }

        if ($last.endTime) {
            $endTime = [datetimeoffset]::Parse(
                [string]$last.endTime
            )
        }

        $resultStatus = ([string]$last.status).ToLowerInvariant()
        $serviceStatus = ([string]$status.status).ToLowerInvariant()

        $isCurrentRun = (
            $null -ne $startTime -and
            $startTime -ge $requestedAtTolerance
        )

        Write-Host ((
            "Indexer={0} ServiceStatus={1} ResultStatus={2} " +
            "RequestedAt={3:u} Start={4} End={5} " +
            "Processed={6} Failed={7} CurrentRun={8}"
        ) -f
            $IndexerName,
            $serviceStatus,
            $resultStatus,
            $RequestedAt,
            $startTime,
            $endTime,
            $last.itemsProcessed,
            $last.itemsFailed,
            $isCurrentRun
        )

        # Check for completion BEFORE looking at the top-level status.
        if ($isCurrentRun -and $null -ne $endTime) {
            if (
                $resultStatus -eq "success" -or
                $resultStatus -eq "warning"
            ) {
                Write-Host ((
                    "Indexer completed: {0} " +
                    "Result={1} Start={2:u} End={3:u}"
                ) -f
                    $IndexerName,
                    $resultStatus,
                    $startTime,
                    $endTime
                )

                return
            }

            $errorsJson = "[]"

            if ($last.errors) {
                $errorsJson = $last.errors |
                    ConvertTo-Json -Depth 10 -Compress
            }

            throw (
                "Indexer '$IndexerName' completed with status " +
                "'$resultStatus'. Errors: $errorsJson"
            )
        }

        # The requested execution has started but has not completed.
        if ($isCurrentRun -and $null -eq $endTime) {
            Write-Host "Indexer '$IndexerName' is still executing."
            Start-Sleep -Seconds $IndexerPollSeconds
            continue
        }

        # The API is still returning a previous execution.
        Write-Host (
            "Waiting for the newly requested run of '$IndexerName' " +
            "to appear in lastResult."
        )

        Start-Sleep -Seconds $IndexerPollSeconds
    }

    throw (
        "Indexer '$IndexerName' timed out after " +
        "$IndexerTimeoutMinutes minutes."
    )
}

function Invoke-AclRefresh {
    param(
        [string]$IndexName
    )

    Write-Host "Running ACL refresh: $IndexName"

    & $PythonExe `
        $UpdateAclScript `
        --index-name $IndexName `
        --search-service $SearchService

    if ($LASTEXITCODE -ne 0) {
        throw "update_acls failed for $IndexName"
    }
}

# ------------------------
# Main
# ------------------------

if (-not (Test-Path $LogDirectory)) {
    $null = New-Item -ItemType Directory -Path $LogDirectory
}

Start-Transcript -Path $LogPath -Force | Out-Null
Write-Host "Logging to: $LogPath"

$Failures = @()
$exitCode = 0

try {

foreach ($prefix in $Prefixes) {

    try {

        $indexers = @(
            "$prefix-pages-indexer"
        )

        if ($prefix -eq "spa-sharepoint") {
            $indexers += "spa-sharepoint-docs-indexer"
        }

        foreach ($indexer in $indexers) {

            $token = Get-SearchToken

            $requestedAt = Invoke-IndexerRun `
                -IndexerName $indexer `
                -Token $token

            Wait-Indexer `
                -IndexerName $indexer `
                -Token $token `
                -RequestedAt $requestedAt
        }

        $chunksIndex = "$prefix-chunks-index"

        Invoke-AclRefresh `
            -IndexName $chunksIndex

        Write-Host "Completed: $prefix"
    }
    catch {

        Write-Host "Failed: $prefix"
        Write-Host "Exception: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Host $_.ScriptStackTrace
        }

        $Failures += [PSCustomObject]@{
            Prefix = $prefix
            Error  = $_.Exception.Message
        }
    }
}

# ------------------------
# Summary
# ------------------------

if ($Failures.Count -gt 0) {

    Write-Host ""
    Write-Host "FAILED PREFIXES"

    $Failures | Format-Table -AutoSize

    $exitCode = 1
}
else {
    Write-Host ""
    Write-Host "All indexers and ACL refreshes completed successfully."
}
}
finally {
    Stop-Transcript | Out-Null
}

exit $exitCode