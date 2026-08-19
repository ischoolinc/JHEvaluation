param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [string]$BaselineVariant = "baseline",

    [string]$CandidateVariant = "optimized",

    [int]$MinimumSamples = 10,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path)) {
    throw "Metrics file not found: $Path"
}

$rows = @(Import-Csv -LiteralPath $Path | Where-Object { $_.Success -eq "true" })
if ($rows.Count -eq 0) {
    throw "The metrics file contains no successful samples."
}

function Get-PercentileValue {
    param(
        [double[]]$Values,
        [double]$Percentile
    )

    $sorted = @($Values | Sort-Object)
    $index = [Math]::Ceiling($Percentile * $sorted.Count) - 1
    if ($index -lt 0) { $index = 0 }
    return [double]$sorted[$index]
}

$statistics = @(
    $rows |
        Group-Object Variant, DeploymentMode, Stage |
        ForEach-Object {
            $first = $_.Group[0]
            $values = [double[]]@($_.Group | ForEach-Object {
                [double]::Parse($_.ElapsedMilliseconds, [Globalization.CultureInfo]::InvariantCulture)
            })

            [pscustomobject]@{
                Variant = $first.Variant
                DeploymentMode = $first.DeploymentMode
                Stage = $first.Stage
                Samples = $values.Count
                MedianMs = [Math]::Round((Get-PercentileValue -Values $values -Percentile 0.5), 3)
                P95Ms = [Math]::Round((Get-PercentileValue -Values $values -Percentile 0.95), 3)
                HasEnoughSamples = $values.Count -ge $MinimumSamples
            }
        } |
        Sort-Object Variant, DeploymentMode, Stage
)

$comparisons = @()
foreach ($baseline in $statistics | Where-Object { $_.Variant -eq $BaselineVariant }) {
    $candidate = $statistics | Where-Object {
        $_.Variant -eq $CandidateVariant -and
        $_.DeploymentMode -eq $baseline.DeploymentMode -and
        $_.Stage -eq $baseline.Stage
    } | Select-Object -First 1

    if ($null -eq $candidate) { continue }

    $medianChangePercent = if ($baseline.MedianMs -eq 0) { 0 } else {
        (($candidate.MedianMs - $baseline.MedianMs) / $baseline.MedianMs) * 100
    }
    $p95ChangePercent = if ($baseline.P95Ms -eq 0) { 0 } else {
        (($candidate.P95Ms - $baseline.P95Ms) / $baseline.P95Ms) * 100
    }

    $passesThreshold = if ($baseline.Stage -ne "total") {
        $null
    } elseif ($baseline.MedianMs -gt 200) {
        $candidate.MedianMs -le ($baseline.MedianMs * 0.8) -and
        $candidate.P95Ms -le ($baseline.P95Ms * 1.1)
    } else {
        $candidate.MedianMs -le $baseline.MedianMs -and
        $candidate.P95Ms -le $baseline.P95Ms
    }

    $comparisons += [pscustomobject]@{
        DeploymentMode = $baseline.DeploymentMode
        Stage = $baseline.Stage
        BaselineSamples = $baseline.Samples
        CandidateSamples = $candidate.Samples
        BaselineMedianMs = $baseline.MedianMs
        CandidateMedianMs = $candidate.MedianMs
        MedianChangePercent = [Math]::Round($medianChangePercent, 2)
        BaselineP95Ms = $baseline.P95Ms
        CandidateP95Ms = $candidate.P95Ms
        P95ChangePercent = [Math]::Round($p95ChangePercent, 2)
        PassesThreshold = $passesThreshold
    }
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $comparisons | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

[pscustomobject]@{
    Statistics = $statistics
    Comparisons = $comparisons
}
