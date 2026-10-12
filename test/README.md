# Tests

Five test layers: unit, smoke, Excel end-to-end, load and liveness. See
`AGENTS.md` for the full guidance and the Excel automation lessons.

## 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

The test project compiles the production sources directly (no add-in build), so
every top-level production file except the intentionally excluded `RedisRtd.cs`
must be in its `Compile` list (a parity test enforces it).
Coverage includes the JSON conversions and the matrix size/total-cell budget,
config load/sanitize (including alias trim/case-insensitivity, the undefined-style
reset and the tolerant per-value parsing), the real-time conflation window
(`ConflationMs` sanitize/resolve/due), the connection/subscription managers
(the shutdown fences, the connect-failure memo, eviction protection and the
idle-down drop), the publish-dedup
LRU cache and the pattern-join marker clearing (every marker of the host),
`TickGate`, the update-check helpers, the `...NonVolatile` signature-parity and
delegation tests, the write-mode dispatch (`SyncWrite`/`AsyncWrites` parsing,
per-host serialization, no host overlap, the synchronous path and the caller
refusal/fallback), the offline `RedisWriteObservable` tests (single delivery,
error text, OnNext-throw still completed, one-shot subscribe with duplicate
delivery, no-op dispose while queued, synchronous enqueue, the enqueue-failure
latch and pathological error messages), the `WaitBounded` pipelined-wait bound
(completed/faulted tasks return true, a never-completing task reports the
timeout), and a project-file parity test that keeps the unit project's
`Compile` list complete (`RedisRtd.cs` is the intentional exclusion).

## 2. Smoke tests (requires Redis, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
```

Covers the ref-counted Pub/Sub behavior (two listeners on one channel, the
last-listener unsubscribe, re-subscription, duplicate suppression,
literal/pattern independence, late joiners, double dispose, origin counters,
argument validation) plus the concurrent-dedup check (4 x 50,000 distinct
payloads against a dedicated Redis: a disposable `redis:7-alpine` container
with a random name on a free ephemeral port when a Linux Docker daemon is
available, else the main host). It also drives the real `SyncWrite` modes
against Redis - the CI-safe end-to-end for the write path, no Excel: the exact
cell text per mode (`OK`, `OK (fire and forget)`, `OK (fire and forget: all)`,
`0 readers`), the reply shapes and the `RedisUDFChannelPublish` markers - and a
live `RedisUDFUpdateAvailable` check against the GitHub release API (skipped
when offline).

## 3. Excel end-to-end (requires Excel + Redis + built XLLs)

```powershell
dotnet build RedisExcel.sln -c Release
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

The script builds the test workbook (`RedisExcel.Test.xlsx`), asserts UDF/RTD
values and covers the multi-workbook Pub/Sub scenario (workbook copy/close and
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
`-SyncWrite <mode>` (write mode for the run: `sync`/`fireforget`/
`fireforget-all`, default `fireforget`), `-AsyncWrites`, `-KeepExcelOpen`.
Supplying a custom `-RedisCli` (e.g.
`docker exec my-redis redis-cli`) disables the automatic `CLIENT KILL` step.

Security software (AV/EDR) is the top cause of flaky Excel COM automation, so
the script reports the AV state up front and a `COM retries: total=...` summary
at the end, and it retries transient COM rejections (including method calls
such as `CalculateFull`). `test\Set-Antivirus.ps1 -Action Status|Disable|Enable`
reports the state and can toggle Defender real-time, but a third-party engine
with self-defense cannot be paused by any generic script: pause it from its UI
and use `Status` to record which state a run had.

Measured per mode (same script; the add-in itself is never the source of the
retries). Excel COM automation is **inherently flaky**: the counts below are the
same with the antivirus active and paused, so a retry total is normal, not a red
flag:

| Mode | Result | COM retries |
| --- | --- | --- |
| `fireforget` (default) | ALL PASS | ~2-5 |
| `sync` | ALL PASS | ~2-5 |
| `fireforget-all` | ALL PASS | ~1-5 |
| `-AsyncWrites` | ALL PASS | ~180-190 |

The `-AsyncWrites` mode drives ~180 method calls, so ~1 retry each accumulates.
Its pending-marker check is deterministic: it points the queued write at an
unreachable host, so the cell stays in Excel's pending `#N/A` for the whole
connect timeout and the marker is always observable (the local-Redis pending
window is a diagnostic only, since it closes in microseconds). Every suite
(unit via a test
output line, smoke and load as `Antivirus: ...`, liveness as
`[LIVENESS] antivirus=...`) prints the active AV state, so a result is
attributable to the machine it ran on.

To compare AV on vs off, `test\Run-ExcelE2E-Matrix.ps1 -Tag <name>` runs all
four write modes (`default`/`sync`/`ffall`/`async`) and logs each to
`%TEMP%\e2e-<tag>-<mode>.log`; run it once with the AV active and once paused,
then `test\Compare-ExcelE2E.ps1 -TagA <on> -TagB <off>` prints them side by
side (status, COM-retry total, SKIP count per mode).

> Building the test workbook only uses `localhost:6379` and
> `test.redisexcel.*` keys/channels. Never commit workbooks or configs that
> reference private hosts.

## 4. Load tests (requires Redis)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "127.0.0.1:6379" 10 2 1
```

Modes: `manager` (RedisSubscriptionManager) and `raw` (plain
StackExchange.Redis baseline). Parameters: host, seconds, publisher threads
(0 = listen-only with an external generator; `redis-benchmark -t publish` is
a silent no-op on Redis 7.4 - use
`redis-benchmark -n 20000 -q -P 16 PUBLISH <channel> <payload>` instead),
listeners, pattern and an optional `[channel]`
(7th argument; defaults to a random `load:<guid>` channel). It reports
throughput, allocated bytes per received message and GC counts. Read-only stress runs
against real servers are allowed, but pass the host only as a command-line
argument (never commit it) and keep the runs short.

## 5. Liveness tests (requires Redis; Docker optional)

```powershell
dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False"
```

Drives the real managers under adversarial conditions (ThreadPool starvation,
subscribe/dispose churn, lock contention on every public entry point,
`CLIENT KILL` storm, async queue/observable bursts, dedup re-join) and enforces
the liveness bounds: a stall past **5 s** fails the run, delivery and the queue
must resume within **30 s**, and the async backlog must drain within **20 s**.
CI runs the quick mode after the smoke tests (`--quick --skip-restart`); the
fuller modes below are meant to be run manually.

Container behavior: without `--container` the suite runs against whatever
Redis answers on the host (the `CLIENT KILL` storm runs only on loopback hosts;
the restart sub-phase prints `SKIP`). With `--container <name>` it manages its
own disposable `redis:7-alpine` container - started mapped to the host's port
(any previous instance of that name is removed first), restarted mid-stream in
phase C and removed at the end - which requires a loopback host and a working
Docker CLI; when Docker is unavailable the container phase is skipped with a
printed reason. `--skip-restart` skips only the restart sub-phase; `--quick`
shortens the attack windows for a fast sanity run. `--matrix` runs a scripted
failure matrix (stop/start, restart, pause, kill storms, `CLIENT PAUSE`,
flapping, reload under traffic, writes during the fault) asserting recovery and
resource stability after every fault; `--soak [minutes]` (default 5) keeps mixed
traffic running and injects a random fault every 20-40 s. Only run the layer
against a disposable/local Redis: the kill storm and the restart are destructive.

> Pass the host only as a command-line argument; never commit it.
