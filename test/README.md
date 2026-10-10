# Tests

Three test layers plus a load test harness. See `AGENTS.md` for the full
guidance and the Excel automation lessons.

## 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

The test project compiles the production sources directly (no add-in build), so
a new production file used by the tests must be added to its `Compile` list.
Coverage includes the JSON conversions, config load/sanitize, the
connection/subscription managers, the publish-dedup LRU cache, `TickGate`, the
update-check helpers, the `...NonVolatile` signature-parity reflection test,
the write-mode dispatch (`SyncWrite`/`AsyncWrites` parsing, per-host
serialization, no host overlap, the synchronous path and the caller
refusal/fallback) and the offline `RedisWriteObservable` tests (single
delivery, error text, OnNext-throw still completed, one-shot subscribe with
duplicate delivery, no-op dispose while queued, synchronous enqueue).

## 2. Smoke tests (requires Redis, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
```

## 3. Excel end-to-end (requires Excel + Redis + built XLLs)

```powershell
dotnet build RedisExcel.sln -c Release
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

The script builds the test workbook (`RedisExcel.Test.xlsx`), asserts UDF/RTD
values and reproduces the Pub/Sub regression scenario (workbook copy/close and
connection blip). It also covers the write modes (`sync`/`fireforget`/
`fireforget-all`, async on and off) and asserts that a `...NonVolatile` write
runs once (a worksheet recalculation must not re-send it). With `-AsyncWrites`
it additionally checks the pending marker (via `WorksheetFunction.IsNA`, so
the check is locale-independent), the write reaching Redis while the cell is
pending, both identical formulas writing, and an argument change dispatching
the new write. Before the sample
can be committed it is sanitized: the local save path and the personal document
metadata are removed. For remote hosts the workbook goes to `%TEMP%` and the
destructive `CLIENT KILL` step is skipped automatically; `-AsyncWrites` runs
also keep the workbook in `%TEMP%`.

Useful parameters: `-RepoRoot <path>` (repository root; defaults to the
script's parent folder, the repo root), `-RedisHost`, `-KeyPrefix`,
`-RealChannel`, `-RealPattern`, `-SkipClientKill`, `-RedisCli`,
`-KeepExcelOpen`. Supplying a custom `-RedisCli` (e.g.
`docker exec my-redis redis-cli`) disables the automatic `CLIENT KILL` step.

> Building the test workbook only uses `localhost:6379` and
> `test.redisexcel.*` keys/channels. Never commit workbooks or configs that
> reference private hosts.

## 4. Load tests (requires Redis)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "127.0.0.1:6379" 10 2 1
```

Modes: `manager` (RedisSubscriptionManager) and `raw` (plain
StackExchange.Redis baseline). Parameters: host, seconds, publisher threads
(0 = listen-only with an external generator such as
`redis-benchmark -t publish`), listeners, pattern and an optional `[channel]`
(7th argument; defaults to a random `load:<guid>` channel). It reports
throughput, allocated bytes per received message and GC counts. Read-only stress runs
against real servers are allowed, but pass the host only as a command-line
argument (never commit it) and keep the runs short.
