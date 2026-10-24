UnivDB v2.9.0 final performance evidence

Product source base: 9ec684a (first performance fix).
Original measurement baseline: 06f5efe800444aca6c23fd93ab1fdd15ebba78e0.
Final changes: bulk append/drop, candidate-only filter copy, unused join copy removal.

Reproduction (repository root, after vbaloop push):
  powershell -File experiments/perf-20261024/fix-02/prepare-final.ps1
  vbaloop push --config _dbg/perf-20261024/vbaloop.toml --json
  vbaloop run TestBenchDb.ProbeFinal --config _dbg/perf-20261024/vbaloop.toml --timeout 600 --json
  vbaloop run TestBenchDb.ProbeJoinPair --config _dbg/perf-20261024/vbaloop.toml --timeout 600 --json

prepare-final rebuilds the disposable _dbg/perf-20261024 lab. It never overwrites
the source baseline evidence in experiments/perf-20261024. Diagnostic macros and
the pinned old join body are injected only into that lab, never the product workbook.
Source: original probe-appendix.txt plus probe-final.txt. QPC, one warmup, three samples.

measurements.jsonl: latest ProbeFinal run only (92 assertions passed).
before-after.json: baseline median compared with final median for matching cases.
join-paired.jsonl / join-paired-summary.json: alternating old/new warm join in one
process, 20k x 30 left and 5k x 2 right; keys overlap on 2000 rows; 32 assertions passed.
tests-after.json / validation.json: product full suite, compile, lint and timing.

Interpretation:
- Later runs show substantial host variability even in unchanged controls. Historical
  join timings alone do not establish regression or improvement; use paired evidence.
- Append-only timing excludes constructing the source table. with_prepare includes
  source LoadFromArray and ID setters. Both exclude initial target construction.
- The historic 10k repeated Insert took 362887.21ms in one execution, not a 3-run median.
  It is not repeated in the permanent test suite. Builder/Upsert equivalence applies
  only to unique new keys; AppendFrom also preserves duplicate keys.
- Absolute timing is recorded, never used as a fragile pass/fail gate.
- Product test count: 126 -> 129; new cases cost 0.871 seconds in the final run.
  Whole suite: 26.565 seconds, including variable unrelated benchmarks.
- CSR peak memory is unmeasured. No claim of universal speedups or memory reduction.

Full Korean implementation/decision report:
  src/docs/tasks/2026-10-24_performance_fixes.md
