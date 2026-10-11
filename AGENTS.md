# AGENTS.md

Guidance for AI coding agents (and humans) working on RedisExcel.

## What this project is

Excel add-in (XLL) written in C# / .NET Framework 4.8 with Excel-DNA:

- **UDFs** — worksheet functions (`RedisUDF*`, `RedisRTD*` status helpers).
- **RTD server** — streaming/polling topics via `=RTD("RedisRtd",, ...)`.
- Backends: StackExchange.Redis, NLog, Newtonsoft.Json.

## Golden rules

1. **English only** — code, comments, identifiers, log messages, docs and commit
   messages are written in English.
2. **Never change the public Excel surface** — function names, argument order,
   `[ExcelArgument]`/`[ExcelFunction]` attributes, RTD commands
   (`GET`, `HGET`, `HGETALL`, `SUB`, `PSUB`) and the ProgId `RedisRtd`.
   Users' spreadsheets depend on every one of them. New additive functions
   (e.g. the v1.3.0 `...NonVolatile` twins) are fine; renames, argument
   changes and removals are not.
3. Keep the `RedisExcel.json` shape backward compatible
   (`RTD` / `UDF` / `Servers`).
4. All Redis access goes through `RedisRuntime`
   (`RedisConnectionManager` / `RedisSubscriptionManager`). Never create a
   `ConnectionMultiplexer` anywhere else.
5. Subscription listeners are registered/removed as `IDisposable` tokens.
   Never call `ISubscriber.Unsubscribe(channel)` without a handler — it would
   tear down every other listener on that channel.
6. Any code running inside a `System.Timers.Timer` callback must catch and log
   exceptions (the timer swallows them silently). The shared `CreateTimer`
   helper in `RedisRtd.cs` already does this.
7. State touched by both the Excel thread and timer threads must be
   thread-safe (`ConcurrentDictionary`, `Interlocked`, or a private lock).
8. Keep `net48` and C# 7.3-compatible syntax.
9. **No private data in tracked files** — no company names, customer names,
   private hostnames or local machine paths. Examples use `localhost:6379` and
   relative paths.

## Layout

| File | Responsibility |
| --- | --- |
| `Config.cs` | `AppConfig`: loads `RedisExcel.json` once, safe defaults, host alias resolution. |
| `RedisConnectionManager.cs` | Single place that creates/caches `ConnectionMultiplexer` per host and pool (`RtdData`/`RtdSub`/`UdfData`). |
| `RedisSubscriptionManager.cs` | Ref-counted Pub/Sub: one StackExchange.Redis handler per `(host, channel, pattern)` broadcasting to N listeners; channels survive reconnects (StackExchange.Redis re-subscribes automatically). |
| `RedisRuntime.cs` | Process-wide singleton wiring connections + subscriptions; shutdown on add-in unload. |
| `RedisRtd.cs` | RTD server lifecycle, topic registries, timers, polling; `RedisRtdStatus` functions. |
| `RedisUDF.cs` | `[ExcelFunction]` implementations; thin wrappers over the managers; 24 `...NonVolatile` write twins. |
| `RedisUdfAsync.cs` | Optional async write dispatch (`AsyncWrites`, default off): per-host FIFO queue + Excel-DNA `ExcelAsyncUtil.Observe` (`RedisWriteObservable`); pure sync passthrough when disabled. |
| `ExcelJson.cs` | `RedisUDFMatrixToJSON` / `RedisUDFJSONToMatrix`. |
| `RedisResultFormatter.cs` | Value formatting sent to Excel (HGETALL as valid JSON). |
| `TickGate.cs` | Non-blocking reentrancy gate for timer callbacks; shared `StripedLocks` helper. |
| `Conflation.cs` | Real-time delivery conflation window (`ConflationMs`): latest value wins per window; sanitize/resolve/due helpers. |
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task; exposed via RedisUDFUpdateAvailable). |

## Runtime model (why it is the way it is)

- Excel may create more than one RTD server per process. Connections and
  subscriptions are **process-wide** (`RedisRuntime`); topics are per RTD
  server instance.
- One StackExchange.Redis handler per channel broadcasts to every listener;
  a channel is unsubscribed only when its last listener leaves.
- Subscribe/unsubscribe network handoff is serialized per registry key
  (striped gates) across channel-state generations: StackExchange.Redis reuses
  its internal per-channel subscription object, so a joiner attaching while
  the previous generation was torn down could skip the wire SUBSCRIBE and
  leave the channel permanently deaf. The disposed-state wait in `Subscribe`
  is bounded (after 500 yields the joiner evicts the stale mapping itself),
  and a failed rollback unsubscribe poisons the state (the next Subscribe
  builds a fresh one) instead of risking a duplicate handler registration.
- Subscription listeners carry an origin tag ("RTD"/"UDF"): the RTD status
  counters (`RedisRTDSubscriptionCount`, `RedisRTDChannelCount`) report RTD
  listeners only, so UDF subscriptions no longer inflate them. After the RTD
  server shuts down the RTD status helpers return `0`/`false`.
- RTD status functions that report the default host, update rates, the
  real-time flag and the message counter reflect the LAST started RTD server;
  the counts (connections/topics/subscriptions/channels) aggregate across
  instances.
- StackExchange.Redis re-subscribes channels automatically after a reconnect;
  no custom resubscribe code is needed (verified live during the v1.1.0 work).
- A `SUB`/`PSUB` RTD topic that fails to subscribe at connect time shows
  `#ERROR` and is retried with a 1s..30s backoff, capped per tick (v1.2.7);
  blank `SUB`/`PSUB` channels are rejected at `ConnectData`; polled commands
  (`GET`/`HGET`/`HGETALL`) retry on every tick.
- RTD push model: the poll timer (`RedisUpdateRateMs`) reads values; the Excel
  timer (`ExcelUpdateRateMs`) flushes dirty values. Real-time updates are
  coalesced per topic by default (`CoalesceRealtimeUpdates`, on): the latest
  value wins per window instead of one Excel update per message.
- Timer callbacks are protected against reentrancy: a tick that fires while the
  previous one is still running is skipped, not queued, so a slow tick never
  overlaps the next one. `ExcelUpdateStyle: Automatic` still drives the
  real-time/timer switch through `MessageCounterThreshold` (messages in the
  last second): above the threshold the Excel tick disables real-time delivery,
  and the 1s tick re-enables it once the rate drops; `<= 0` disables the
  switch. With `CoalesceRealtimeUpdates` on (default) delivery is coalesced per
  window regardless, but with it off the switch decides per-message pushes vs
  dirty-value flushes on the Excel tick.
- `HGETALL` output is valid JSON: `{"field":"value",...}`; a missing hash
  returns `{}` (not a sentinel string).
- `RedisUDFChannelPublishIfChanged` dedup: in `sync` the last payload is
  remembered per host/channel only after it was actually delivered (a publish
  with zero readers drops the marker), and the marker is cleared when a
  listener (re)joins - RTD `SUB`/`PSUB` as well as UDF `ChannelLatest`; a
  pattern subscription clears the matching channels of that host - and by an
  explicit `RedisUDFChannelUnsubscribe` that removes that host/channel's local
  listener. Generic listener leaves (for example an RTD topic disconnecting)
  do not clear the marker. So a (re)joining local subscriber is never starved
  by a publish it did not see. In the fire-and-forget modes the marker is
  recorded without a confirmed delivery (an external subscriber joining later
  can miss an unchanged payload until it changes); local listener (re)joins
  still clear it. The cache is safe under concurrent recalculation and
  LRU-capped by `PublishDedupCacheSize` (default 10000).
- Pub/Sub delivery order is **not** guaranteed: StackExchange.Redis hands every
  channel callback to the thread pool, so two messages in flight together can
  complete out of order (the raw baseline shows inversions in every run under
  load). Nothing is dropped and no thread stalls; the visible effect is that a
  cell (`SUB`/`PSUB`/`ChannelLatest`) may transiently hold the older of two
  overlapping values until the next message, and a burst can leave the dedup
  marker on the older payload. Ordering-sensitive feeds must carry a
  timestamp/sequence field and compare it in the formula; polled reads
  (`GET`/`HGET`/`HGETALL`) are single round trips per tick and are unaffected.
- `ChannelLatest` lifecycle: per-key epochs + per-listener locks. An
  unsubscribe/reset (including add-in reload) bumps the epoch before removing,
  so a subscribe that raced it declines to install (no ghost listener), and a
  callback that passed its closed check cannot re-add a payload after the
  removal; the network Subscribe runs outside the lifecycle lock and the
  install re-checks the epoch under it.
- Reads and status functions are volatile by design: they re-execute on every
  recalculation (F9/edit) so they stay fresh. Every write function also has an
  additive `...NonVolatile` twin (same args/defaults, thin delegation, no
  `IsVolatile`): Excel evaluates it only on entry and when an argument cell
  changes - `F9`/edits do not re-run it, `Ctrl+Alt+F9` (full recalculation)
  does. Reads have no twins (`ChannelLatest` must refresh by itself), and the
  pure JSON conversions (`RedisUDFMatrixToJSON`/`RedisUDFJSONToMatrix`) were
  already non-volatile. Keep a reflection signature-parity test for every pair.
- Write delivery is configurable (`SyncWrite`): `sync` blocks for every reply
  (pre-v1.3.0 behavior); `fireforget` (default) sends result-agnostic writes
  (`Set`, `SetJSON`, `SetKV`/`SetKVPair`, `SetEx`, `Rename`, `HashSet`,
  `HashSetMultiple`, list pushes, channel publishes) with
  `CommandFlags.FireAndForget` and returns `OK (fire and forget)`, while
  reply-dependent writes (`Del`, `Incr`, `IncrBy`, `Expire`, `SetAdd`,
  `SetRemove`, `HashDel`, list pops) stay blocking; `fireforget-all` sends
  every write FireAndForget (reply-dependent writes return
  `OK (fire and forget: all)`). `ChannelUnsubscribe` always removes the local listeners
  deterministically (never fire-and-forget). In fire-and-forget modes an
  error detected after the dispatch (or a delivery failure) is only logged
  (the cell keeps the marker) while validation/config failures before the
  dispatch still surface as `Error:`, and publishes return the marker instead
  of the readers count.
- `AsyncWrites` (default false) dispatches writes through Excel-DNA's
  Observe-based async support (`ExcelAsyncUtil.Observe` with a custom
  `RedisWriteObservable`, RTD-based) so the Excel thread never blocks: the cell
  shows the pending marker (`#N/A`) and then the real value/error (or the
  fire-and-forget marker). The write is enqueued exactly once per registered
  call - Excel-DNA creates the observable at registration and its `Subscribe`
  runs synchronously on the Excel thread during the internal RTD `ConnectData`,
  and the recalculation that delivers the result returns the cached value for
  the same identity (Excel-DNA state lookup) without re-subscribing; a
  duplicate `Subscribe` on the same observable never re-enqueues the write and
  still receives the queued result (one-shot guard). No
  thread-pool thread is held per pending write: the queue continuation delivers
  `OnNext`/`OnCompleted`, so a large same-host burst no longer throttles the
  pool (the old classic `ExcelAsyncUtil.Run` dispatch blocked one pool thread
  per queued item). The async identity is the calling cell + resolved host +
  the UDF's own arguments (the cell reference is structurally equal across the
  completed re-call), so different cells never share one call and an argument
  change dispatches a new write. Repeated evaluations with unchanged arguments
  return the cached value while the internal topic stays connected (a volatile
  write is not re-sent by AsyncWrites); when Excel detaches the topic (e.g. an
  unchanged recalculation) the next evaluation re-registers the call and a
  volatile write is issued again, so the dedup is best-effort, not
  exactly-once across the sheet lifetime; inserting/moving rows or columns
  changes the cell reference and re-issues the write; a call without a
  worksheet caller is refused with an Error cell (the identity would be shared
  or unstable). Same-host writes are serialized by a per-host FIFO queue fed on
  the Excel thread during `Subscribe`, so the order is the formula evaluation
  order again. `AsyncWrites` decides where a write blocks (Excel thread vs
  worker) and `SyncWrite` decides whether the reply is awaited on that thread,
  so `sync` + async yields real replies without blocking Excel.
- `AutoOpen` raises the ThreadPool floor (min 16 worker/IO threads, best
  effort) so a starved pool (other add-ins, hosted CLR, policy) cannot stall
  the async write queue or RTD timer continuations; a failing ComServer
  registration is logged instead of aborting `AutoOpen`.
- Connection counters report live multiplexers per pool only (closed/failed
  entries are not counted; the RTD connection count is `RtdData` + `RtdSub`
  only), and shutdown no longer blocks on closing connections.
- Cached multiplexer eviction (past the 512-entry cap) vetoes hosts with live
  subscription listeners (`RedisConnectionManager.SetEvictionProtection` wired
  to `RedisSubscriptionManager.HasActiveSubscribers`), reclaims entries whose
  Lazy factory never materialized, and invalidates the cached
  `IDatabase`/`ISubscriber` wrappers of the evicted multiplexer so later calls
  rebuild instead of throwing `ObjectDisposedException` forever.
- A connect failure is memoized for 2 s per (host, pool): a burst of `SUB`
  topics or volatile reads against a dead host pays one connect attempt
  instead of N x ConnectTimeout; a successful connect clears the memo, and the
  retry loops/recalculations recover after the window.
- Every pipelined batch wait is bounded (`RedisUDF.WaitBounded` with the
  configured response timeout + 500 ms): an orphaned batch task (multiplexer
  disposed mid-execute) surfaces as a timeout instead of stalling a timer or
  the Excel thread forever. Covers RTD polled reads and the UDF batch
  `...Multiples` functions.
- `CoalesceRealtimeUpdates`/`ConflationMs`: in real-time mode a topic that keeps
  changing is pushed to Excel at most once per window with the latest value
  (latest wins), instead of one push per message. The root `ConflationMs`
  (ms, `0` = off, capped at 3600000) is the explicit window; when it is absent
  the legacy boolean decides - `true` = the Excel tick interval (the pre-existing
  behaviour), `false` = per-message. This is the main lever against screen
  flicker and the recalculation cascade of dependent formulas on fast feeds; a
  window below ~50-100 ms barely helps because the flush rides the Excel tick.
- `RedisRuntime.ResetAfterAddInReload` supports a same-process add-in reload
  without reusing the previous managers.
- Values and identifiers (keys, hash keys, fields, channels, patterns) written
  to Redis always use the invariant culture (decimal point), regardless of the
  Excel locale.
- The host argument may be a full connection string, so it can carry a
  `password=...`. Every log line and every verbose error cell masks credentials
  (`AppConfig.MaskHost`: `password=****` / `pass=****`); host/port/options stay
  readable. This covers the UDF/RTD error text, the `#ERROR` RTD cell and the
  `ConnectData` argument log.
- Excel error cells are rejected in every scalar argument position; date/time cells
  arrive as their Excel serial number and boolean cells serialize as
  `true`/`false`.
- An empty-string cell is a valid Redis name/pattern while a truly blank cell
  is a missing argument.
- 2x2 pair ranges are read as 2 rows x 2 columns (vertical pairs).
- Identical consecutive payloads are skipped before decoding for literal
  subscriptions and GET/HGET polling (`SkipRepeatedMessages`, default on):
  payloads are compared without a decode step, using `RedisValue` equality;
  PSUB patterns are never deduplicated because channels interleave. Unchanged
  HGETALL hashes are compared field-by-field and skipped too. `RedisValue`
  equality normalizes numeric formatting (e.g. `"1.00"` equals `"1"`), so
  formatting-only changes are treated as duplicates.
- The config file is read once per Excel process (restart Excel after editing
  it) and searched in: user profile, Excel folder, `C:\Windows` (first existing
  wins). A malformed first-existing file uses safe defaults instead of falling
  through to a lower-priority file. When no config file is found at all, the
  legacy no-file default for `ExcelUpdateRateMs` applies (1000 ms, vs the
  loaded-file default of 100 ms). An undefined `ExcelUpdateStyle` falls back to
  `Automatic`, and host aliases are matched after trimming and
  case-insensitively (exact match first).

## Build

```powershell
dotnet build RedisExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`:

- `RedisExcel-packed.xll` (32-bit) and `RedisExcel64-packed.xll` (64-bit),
- `RedisExcel.dll`, plus the dependency assemblies are packed inside the XLLs.

CI (`.github/workflows/build.yml`) runs on pull requests and main pushes; on
`v*` tags it passes `/p:InformationalVersion=<tag>`, which is embedded in the
Redis `ClientName` for diagnostics (shown as `dev` for local builds). The quick
liveness step is wrapped in a bounded retry (up to 3 attempts, 5s apart): the
shared runner can stall the whole process for a moment, which trips the
delivery-gap/`udfErrors` bounds on infrastructure noise; the bounds themselves
are unchanged, so a real regression fails every attempt.

## Tests

### 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

Covers: `ExcelJson` conversions and the matrix size/total-cell budget,
`AppConfig` load/sanitize (including alias trim/case-insensitivity, the
undefined-`ExcelUpdateStyle` reset and the tolerant per-value parsing),
`ResolveHostCore` (exact-match priority), the real-time conflation window
(`ConflationMs` sanitize/resolve/due), the
`RedisConnectionManager`/`RedisSubscriptionManager`
behavior (the shutdown fences, the connect-failure memo, eviction protection and
the idle-down drop), the `PublishIfChanged` dedup LRU cache and the
pattern-join marker clearing (every marker of the host), subscription keys,
HGETALL formatting, the `TickGate` reentrancy helper, `UpdateCheckTests`
(`IsNewer`/`NormalizeTag`), `RedisValueLocaleTests` (de-DE culture), the
`...NonVolatile` signature-parity and offline delegation tests (the volatile
write set is pinned to the 24 twins), the write-mode
(`SyncWrite`/`AsyncWrites`) parsing plus async write dispatch (per-host serial
order, no host overlap, the synchronous path, the caller refusal and the
invalid-host fallback), the offline `RedisWriteObservable` tests (single
delivery + completion, error text, an observer whose `OnNext` throws is still
completed, one-shot subscribe with duplicate delivery, no-op dispose while
queued, synchronous enqueue, the enqueue-failure latch and pathological error
messages), the `WaitBounded` pipelined-wait bound (completed/faulted tasks
return true, a never-completing task reports the timeout), and a project-file
parity test that keeps the
unit project's `Compile` list complete (RedisRtd.cs is the intentional
exclusion).

The unit, smoke, load and liveness test projects compile the production
sources directly (linked `Compile` items), so a new production `.cs` needed by
tests must be added to their `Compile` lists; a unit parity test pins the unit
project's list (every top-level source except the intentionally excluded
`RedisRtd.cs`). `dotnet test` no longer builds or packs the
add-in; CI builds it with msbuild.

### 2. Smoke tests (requires a Redis server, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
```

Covers the ref-counted Pub/Sub behavior: two listeners on one channel, one
leaves and the other keeps receiving; channel unsubscribed only after the last
listener leaves; re-subscription works.

### 3. Excel end-to-end

```powershell
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Uses a hidden Excel instance: loads the packed XLL, builds UDF/RTD test sheets,
asserts values against Redis, then reproduces the reported regression
(`v1.1.0`): opens a COPY of the workbook, publishes messages, closes the copy
and verifies the original keeps receiving; finally kills the Pub/Sub
connections server-side and verifies automatic recovery. It also exercises the
write modes (`sync`/`fireforget`/`fireforget-all`, async on and off) and
asserts that a `...NonVolatile` write runs once (a worksheet recalculation must
not re-send it). With `-AsyncWrites` it additionally checks the pending marker
(`WorksheetFunction.IsNA`, locale-independent), that the queued write reaches
Redis while the cell is still pending and the delivery recalculation does not
re-run it, that two identical formulas in different cells both write, and that
an argument change dispatches the new write.

Useful parameters:

- `-RepoRoot <path>` — repository root; defaults to the parent folder of the
  script (the repo root).
- `-RedisHost <host:port>` — default `localhost:6379`.
- `-KeyPrefix <prefix>` — default `test.redisexcel` (keys and channels).
- `-RealChannel <name>` / `-RealPattern <pattern>` — optionally subscribe to
  real read-only channels to check live data.
- `-SkipClientKill` — never run `CLIENT KILL TYPE pubsub`
  (auto-skipped for non-local hosts).
- `-RedisCli <command>` — custom Redis CLI command (e.g.
  `docker exec my-redis redis-cli`); a custom CLI disables the automatic
  `CLIENT KILL` step.
- `-SyncWrite <mode>` — write mode for the run: `sync`, `fireforget`
  (default) or `fireforget-all`; written to the user's `RedisExcel.json` for
  the run and restored afterwards.
- `-AsyncWrites` — run with async write dispatch (same config handling), so
  the pending-marker and single-delivery checks above are exercised.
- `-KeepExcelOpen` — leave Excel open for debugging.

For local hosts the workbook is saved to `test\RedisExcel.Test.xlsx` (committed
sample). For remote hosts it is saved to `%TEMP%` so the host never lands in
the repository. The committed workbook only references `localhost:6379` and
`test.redisexcel.*`.

### 4. Load tests (requires Redis)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "127.0.0.1:6379" 10 2 1
```

`manager` exercises the subscription broadcast path; `raw` is the plain
StackExchange.Redis baseline. Parameters: host, seconds, publishers (0 =
listen-only with an external generator: `redis-benchmark -t publish` is a
silent no-op on Redis 7.4 - use
`redis-benchmark -n 20000 -q -P 16 PUBLISH <channel> <payload>` instead),
listeners, pattern and an optional channel (default random `load:<guid>`;
with `pattern` and listen-only mode the default is `*`). Reports throughput,
allocated bytes per received message and GC counts. Read-only stress runs
against real servers are allowed, but pass the host only as a command-line
argument (never commit it) and keep the runs short.

### 5. Liveness tests (requires Redis; Docker optional)

```powershell
dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False"
```

Drives the real managers under adversarial conditions — ThreadPool starvation,
subscribe/dispose churn, lock contention on every public entry point,
`CLIENT KILL TYPE pubsub` storms, async queue/observable bursts and dedup
re-join — and enforces the liveness bounds with a dedicated process heartbeat
and watchdog thread: a stall past **5 s** fails the run, delivery/queue must
resume within **30 s** and the async backlog must drain within **20 s**. A
`RedisTimeoutException` on a contending synchronous UDF call is transient (the
same class as the worker timeouts) and is tolerated: only non-timeout UDF errors
fail the run, with the transient timeouts bounded so a collapse still fails. The
classification matches the whole StackExchange.Redis timeout family
(`Timeout performing ...`, `Timeout awaiting response ...`,
`The message timed out in the backlog ...`), not just the backlog wording.

With `--container <name>` the suite manages its own disposable
`redis:7-alpine` container (started on the host's port, restarted mid-stream in
phase C, removed at the end; loopback hosts and a working Docker CLI required).
`--skip-restart` skips only the restart sub-phase (the kill storm still runs on
loopback hosts) and `--quick` shortens the attack windows for a fast sanity
run. `--matrix` runs a scripted failure matrix (stop/start, restart, pause,
kill storms, `CLIENT PAUSE`, flapping, reload under traffic, writes during the
fault) asserting recovery and resource stability after every fault;
`--soak [minutes]` (default 5) keeps mixed traffic running and injects a random
fault every 20-40 s. CI runs `--quick --skip-restart` after the smoke tests;
run the fuller modes manually against a disposable/local server.

## Excel automation lessons (hard-won)

These cost real debugging time — read before writing automation.

1. **Antivirus/EDR can intermittently block Excel COM automation.** Symptoms
   are misleading errors (`InvalidCastException: cannot convert Int32 to
   String`, `RPC_E_CALL_REJECTED`, COM calls hanging or aborting). The E2E
   script retries with backoff and has a warm-up phase; if a run fails this
   way, just run it again. The script prints the registered antivirus
   products (Windows Security Center) before creating Excel, so an
   AV-enabled environment is visible immediately. Add an exception for
   `EXCEL.EXE` and the script to
   the security product when possible. Never assume such an error means the
   add-in is broken.
2. **Elevated processes ignore per-user COM classes.** The add-in registers
   `RedisRtd` under `HKCU\Software\Classes` in `AutoOpen`. If Excel (or the
   whole automation session) runs elevated, `=RTD("RedisRtd", ...)` shows
   `#N/D` and `CLSIDFromProgID('RedisRtd')` fails with `0x800401F3`, even
   though the registry keys exist. Workaround for test machines: register the
   ProgID/CLSID under `HKLM\Software\Classes` (admin), pointing at the packed
   XLL. This is an environment workaround, not part of the normal install.
3. **`CLIENT KILL TYPE pubsub` is destructive** — it disconnects every Pub/Sub
   client of that server. Only use it against a disposable/local Redis. The
   E2E script skips it for non-local hosts.
4. **Never kill `EXCEL.EXE` processes blindly** — the user may have workbooks
   open. The E2E script creates its own hidden instance and quits it; check
   for leftover processes before re-running (a blocked instance shows an empty
   window title).
5. **PowerShell + COM quirks seen in the wild:**
   - Mixing value types on the same call site can hit a member-cache bug
     (e.g. writing a string first, then an int, to `Range.Value2`). Cast
     explicitly: `[string]$value` / `[double]$value`.
   - `Application.CalculateBeforeSave` may be rejected with `0x800A03EC`
     depending on the build; avoid or wrap it.
   - Prefer `Range("A1")` addresses over `Cells.Item(r, c)` for clear errors.
   - Scriptblocks invoked with `&` see variables through dynamic scoping; pass
     parameters explicitly when in doubt.
6. **RTD cells display `#N/D` when the RTD server could not be created** —
   check COM registration before suspecting the add-in logic.

## Release

Push an annotated `v*` tag (from v1.2.7 on); the CI workflow builds and
publishes the packed XLLs plus `NLog.config` and `RedisExcel.json` as release
assets. Current release: `v1.4.1`; the next planned version (`v1.5.0`/`v2.0.0`)
is not scheduled yet (see the gitignored `DESIGN-v2.0.md`).
