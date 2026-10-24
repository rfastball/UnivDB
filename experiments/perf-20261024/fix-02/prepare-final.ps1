$ErrorActionPreference = 'Stop'
# Rebuild the disposable lab using current product source; baseline evidence remains immutable.
& (Join-Path $PSScriptRoot '..\prepare.ps1')
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$benchPath = Join-Path $repo '_dbg\perf-20261024\src\TestBenchDb.bas'
$bench = [IO.File]::ReadAllText($benchPath)
$append = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'probe-final.txt'))
# Paired join control is pinned to the pre-change source, not a rewritten approximation.
$engine = (git -C $repo show '9ec684a:src/modEngine.bas') -join "`n"
$start = $engine.IndexOf('Public Function JoinTable(')
$end = $engine.IndexOf('End Function', $start) + 'End Function'.Length
$originalJoin = $engine.Substring($start, $end-$start).Replace('JoinTable', 'ProbeOriginalJoin')
$start = $engine.IndexOf('Private Function IsHeaderExists(')
$end = $engine.IndexOf('End Function', $start) + 'End Function'.Length
$append = $append + "`n`n" + $originalJoin + "`n`n" + $engine.Substring($start, $end-$start)
[IO.File]::WriteAllText($benchPath, (($bench.TrimEnd()+"`n`n"+$append.TrimEnd()) -replace "`r?`n","`r`n")+"`r`n", [Text.UTF8Encoding]::new($false))
