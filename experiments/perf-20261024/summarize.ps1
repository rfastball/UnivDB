$ErrorActionPreference = 'Stop'
$records = Get-Content (Join-Path $PSScriptRoot 'measurements.jsonl') | ForEach-Object { $_ | ConvertFrom-Json }
$summary = $records | Where-Object { $_.event -eq 'perf_probe' -and $_.data.sample -ge 0 } | Group-Object { $_.data.case } | ForEach-Object {
    $values = @($_.Group.data.ms | Sort-Object)
    [pscustomobject]@{
        case = $_.Name
        samples = $values.Count
        median_ms = $values[[int][math]::Floor($values.Count / 2)]
        min_ms = $values[0]
        max_ms = $values[-1]
    }
}
$summary | ConvertTo-Json | Set-Content (Join-Path $PSScriptRoot 'summary.json')
$summary | Format-Table -AutoSize
$failed = @($records | Where-Object { $_.kind -eq 'A' -and -not $_.ok })
if ($failed.Count) { throw "Measurement assertions failed: $($failed.Count)" }
