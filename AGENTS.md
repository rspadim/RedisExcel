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
| `RedisUdfAsync.cs` | Optional async write dispatch (`AsyncWrites`, default off): per-host FIFO queue + Excel-DNA `ExcelAsyncUtil.Run`; pure sync passthrough when disabled. |
| `ExcelJson.cs` | `RedisUDFMatrixToJSON` / `RedisUDFJSONToMatrix`. |
| `RedisResultFormatter.cs` | Value formatting sent to Excel (HGETALL as valid JSON). |
| `TickGate.cs` | Non-blocking reentrancy gate for timer callbacks. |
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task; exposed via RedisUDFUpdateAvailable). |

## Runtime model (why it is the way it is)

- Excel may create more than one RTD server per process. Connections and
  subscriptions are **process-wide** (`RedisRuntime`); topics are per RTD
  server instance.
- One StackExchange.Redis handler per channel broadcasts to every listener;
  a channel is unsubscribed only when its last listener leaves.
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
  overlaps the next one. The `Automatic` threshold machinery is kept for
  status/config compatibility, while delivery is coalesced by default
  (`CoalesceRealtimeUpdates`).
- `HGETALL` output is valid JSON: `{"field":"value",...}`; a missing hash
  returns `{}` (not a sentinel string).
- `RedisUDFChannelPublishIfChanged` dedup: in `sync` the last payload is
  remembered per host/channel only after it was actually delivered, and the
  marker is cleared when any listener joins or leaves - RTD `SUB`/`PSUB` as
  well as UDF `ChannelLatest`; a pattern subscription clears the matching
  channels of that host - so a late subscriber is never starved by a publish
  it did not see. In the fire-and-forget modes the marker is recorded without
  a confirmed delivery (an external subscriber joining later can miss an
  unchanged payload until it changes); local listener joins still clear it.
  The cache is safe under concurrent recalculation and LRU-capped by
  `PublishDedupCacheSize` (default 10000).
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
  `CommandFlags.FireAndForget` and returns `OK FireForget`, while
  reply-dependent writes (`Del`, `Incr`, `IncrBy`, `Expire`, `SetAdd`,
  `SetRemove`, `HashDel`, list pops) stay blocking; `fireforget-all` sends
  every write FireAndForget (reply-dependent writes return
  `OK-FireForgetAll`). `ChannelUnsubscribe` always removes the local listeners
  deterministically (never fire-and-forget). In fire-and-forget modes an
  error detected after the dispatch (or a delivery failure) is only logged
  (the cell keeps the marker) while validation/config failures before the
  dispatch still surface as `Error:`, and publishes return the marker instead
  of the readers count.
- `AsyncWrites` (default false) dispatches writes through Excel-DNA's async
  support (`ExcelAsyncUtil.Run`, RTD-based) so the Excel thread never blocks:
  the cell shows the pending marker (`#N/A`) and then the real value/error (or
  the fire-and-forget marker). The write is enqueued exactly once per
  registered call - the enqueue lives inside Excel-DNA's single-shot delegate,
  and the recalculation that delivers the result returns the cached value for
  the same identity instead of re-running the write. The async identity is the
  calling cell + resolved host + the UDF's own arguments (the cell reference is
  structurally equal across the completed re-call), so different cells never
  share one call and an argument change dispatches a new write. Same-host
  writes are serialized by a per-host FIFO queue; the order is the dispatch
  order (strict formula order is not guaranteed). `AsyncWrites` decides where a
  write blocks (Excel thread vs worker) and `SyncWrite` decides whether the
  reply is awaited on that thread, so `sync` + async yields real replies
  without blocking Excel.
- Connection counters report live multiplexers per pool only (closed/failed
  entries are not counted; the RTD connection count is `RtdData` + `RtdSub`
  only), and shutdown no longer blocks on closing connections.
- `RedisRuntime.ResetAfterAddInReload` supports a same-process add-in reload
  without reusing the previous managers.
- Values and identifiers (keys, hash keys, fields, channels, patterns) written
  to Redis always use the invariant culture (decimal point), regardless of the
  Excel locale.
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
  through to a lower-priority file.

## Build

```powershell
dotnet build RedisExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`:

- `RedisExcel-packed.xll` (32-bit) and `RedisExcel64-packed.xll` (64-bit),
- `RedisExcel.dll`, plus the dependency assemblies are packed inside the XLLs.

CI (`.github/workflows/build.yml`) runs on pull requests and main pushes; on
`v*` tags it passes `/p:InformationalVersion=<tag>`, which is embedded in the
Redis `ClientName` for diagnostics (shown as `dev` for local builds).

## Tests

### 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

Covers: `ExcelJson` conversions, `AppConfig` load/sanitize and
`ResolveHostCore`, the `RedisConnectionManager`/`RedisSubscriptionManager`
behavior, the `PublishIfChanged` dedup LRU cache, subscription keys, HGETALL
formatting, the `TickGate` reentrancy helper, `UpdateCheckTests`
(`IsNewer`/`NormalizeTag`), `RedisValueLocaleTests` (de-DE culture), the
`...NonVolatile` signature-parity reflection test, and the write-mode
(`SyncWrite`/`AsyncWrites`) parsing plus async write dispatch (per-host serial
order and the synchronous path).

The unit, smoke and load test projects compile the production sources directly
(linked `Compile` items), so a new production `.cs` needed by tests must be
added to their `Compile` lists. `dotnet test` no longer builds or packs the
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
not re-send it).

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
listen-only with an external generator like `redis-benchmark -t publish`),
listeners, pattern and an optional channel (default random `load:<guid>`;
with `pattern` and listen-only mode the default is `*`). Reports throughput,
allocated bytes per received message and GC counts. Read-only stress runs
against real servers are allowed, but pass the host only as a command-line
argument (never commit it) and keep the runs short.

## Excel automation lessons (hard-won)

These cost real debugging time — read before writing automation.

1. **Antivirus/EDR can intermittently block Excel COM automation.** Symptoms
   are misleading errors (`InvalidCastException: cannot convert Int32 to
   String`, `RPC_E_CALL_REJECTED`, COM calls hanging or aborting). The E2E
   script retries with backoff and has a warm-up phase; if a run fails this
   way, just run it again. Add an exception for `EXCEL.EXE` and the script to
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
assets. Current release: `v1.3.0`; next planned version: TBD.
