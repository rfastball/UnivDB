$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$lab = Join-Path $repo '_dbg\perf-20261024'
New-Item -ItemType Directory -Force -Path "$lab\src" | Out-Null
Copy-Item -LiteralPath "$repo\UnivDB.xlsm" -Destination "$lab\UnivDB.xlsm"
Get-ChildItem -LiteralPath "$repo\src" -File | Where-Object { $_.Extension -in '.bas','.cls','.frm','.frx' } | Copy-Item -Destination "$lab\src"
Copy-Item -LiteralPath "$repo\vbaloop.toml" -Destination "$lab\vbaloop.toml"
$utf8 = [Text.UTF8Encoding]::new($false)
$bench = [IO.File]::ReadAllText("$repo\src\TestBenchDb.bas")
$append = [IO.File]::ReadAllText("$PSScriptRoot\probe-appendix.txt")
[IO.File]::WriteAllText("$lab\src\TestBenchDb.bas", (($bench.TrimEnd()+"`n`n"+$append.TrimEnd()) -replace "`r?`n","`r`n")+"`r`n", $utf8)
$row = [IO.File]::ReadAllText("$repo\src\DbRowSet.cls").Replace('DbRowSet','ProbeRowSet')
$row = $row.Replace('Private mMaxRow As Long', "Private mMaxRow As Long`r`nPrivate mCapacity As Long")
$row = $row.Replace('    mMaxRow = 0', "    mMaxRow = 0`r`n    mCapacity = 0")
# The control is used only for AddRow/ToArray, never intersection or union.
$start = $row.IndexOf('Private Sub EnsureCapacity(')
$row = $row.Substring(0,$start) + @'
Private Sub EnsureCapacity(ByVal rowId As Long)
    If rowId <= mMaxRow Then Exit Sub
    mMaxRow = rowId
    If rowId <= mCapacity Then Exit Sub
    If mCapacity = 0 Then mCapacity = 256
    Do While mCapacity < rowId
        mCapacity = mCapacity * 2
    Loop
    ReDim Preserve mFlags(1 To mCapacity)
End Sub
'@
[IO.File]::WriteAllText("$lab\src\ProbeRowSet.cls", ($row.TrimEnd() -replace "`r?`n","`r`n")+"`r`n", $utf8)
Get-FileHash -LiteralPath "$repo\UnivDB.xlsm" | ConvertTo-Json | Set-Content "$lab\original-workbook-hash.json"
Write-Output $lab
