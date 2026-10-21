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
   (e.g. the `...NonVolatile` write twins) are fine; renames, argument
   changes and removals are not.
3. The `RedisExcel.json` shape is **v1.6**: the server is declared once
   (`Default` + `Servers`, each server a string or `{ host, connectTimeout,
   syncTimeout }`; `Timeout` holds the global defaults). RTD owns its rates plus
   `CoalesceRealtimeUpdates`. A wrong VALUE on a known key rejects the whole file
   (`#CONFIG`); an extra/unknown key is a log warning only.
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
| `Config.cs` | `AppConfig`: loads `RedisExcel.json` once (v1.6 shape: `Default`/`Servers`/`Timeout`), safe defaults, server-name resolution. |
| `ConfigSchema.cs` | Embedded JSON Schema for `RedisExcel.json`: validates with NJsonSchema (comments accepted); a wrong value rejects the file (fail-closed), an extra key warns; flags the removed v1 keys. |
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
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task; exposed via RedisUDFUpdateAvailable). `BuildInfo` (in `RedisConnectionManager.cs`) resolves the readable version reported by `RedisUDFVersion`. |
| `installer/` | Inno Setup script (`RedisExcel.iss`) + README: a per-user helper installer (copies the files, keeps the matching XLL, writes a starter config). The wizard ships **10 languages** (English, both Chinese, pt-BR, es, ru, fr, ko, de, ja; the two Chinese `.isl` live in `installer/languages/`, the rest come with Inno). The add-in inside Excel is **English-only** on purpose (localized cell text would make volatile functions locale-dependent). Built by CI; see `installer/README.md`. |
| `installer/banner.ps1` | The console header the helper scripts print (5-row ASCII wordmark + version + project URL). ASCII only, because the console renders in an OEM codepage; degrades to one compact line when the window is narrow or the output is redirected. Dot-sourced defensively. |
| `docs/`, `overrides/`, `mkdocs.yml`, `tools/make-social-cards.py` | The landing page (MkDocs Material, pt-BR/en/es via `mkdocs-static-i18n`) published by `pages.yml`. Each language has its own social card (`docs/assets/social-card*.png`) used as the page hero and as `og:image` (set per page through the `og_image` front matter, read by `overrides/main.html`); `tools/make-social-cards.py` regenerates them from the original by repainting only the subtitle, so a new language needs one line there. |

## Runtime model (why it is the way it is)

- Excel may create more than one RTD server per process. Connections and
  subscriptions are **process-wide** (`RedisRuntime`); topics are per RTD
  server instance.
- One StackExchange.Redis handler per channel broadcasts to every listener;
  a channel is unsubscribed only when its last listener leaves.
- Subscribe/unsubscribe network handoff is serialized per registry key
  (striped gates) across channel-state generations: StackExchange.Redis reuses
  its internal per-channel subscription object, so a joiner attaching while
  the previous generation is being torn down could skip the wire SUBSCRIBE and
  leave the channel permanently deaf. The disposed-state wait in `Subscribe`
  is bounded (after 500 yields the joiner evicts the stale mapping itself),
  and a failed rollback unsubscribe poisons the state (the next Subscribe
  builds a fresh one) instead of risking a duplicate handler registration.
- Subscription listeners carry an origin tag ("RTD"/"UDF"): the RTD status
  counters (`RedisRTDSubscriptionCount`, `RedisRTDChannelCount`) report RTD
  listeners only, so UDF subscriptions do not inflate them. After the RTD
  server shuts down the RTD status helpers return `0`/`false`.
- RTD status functions that report the default host, update rates, the
  real-time flag and the message counter reflect the LAST started RTD server;
  the counts (connections/topics/subscriptions/channels) aggregate across
  instances.
- StackExchange.Redis re-subscribes channels automatically after a reconnect;
  no custom resubscribe code is needed.
- A `SUB`/`PSUB` RTD topic that fails to subscribe at connect time shows
  `#ERROR` and is retried with a 1s..30s backoff, capped per tick;
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
  do not clear the marker. A (re)joining local subscriber is thereby never
  starved by a publish it did not see. The marker matches across host
  spellings of the same endpoint (`localhost` vs `127.0.0.1`, or the same host
  with different options): the join purge normalizes both sides to the parsed
  host:port before comparing. In the fire-and-forget modes the marker is
  recorded without a confirmed delivery (an external subscriber joining later
  can miss an unchanged payload until it changes); local listener (re)joins
  still clear it. The cache is safe under concurrent recalculation and
  LRU-capped by `PublishDedupCacheSize` (default 10000).
- Pub/Sub delivery order is **not** guaranteed: StackExchange.Redis hands every
  channel callback to the thread pool, so two messages in flight together can
  complete out of order (the raw baseline shows inversions in every run under
  load). Nothing is dropped and no thread stalls; the visible effect is that a
  cell (`SUB`/`PSUB`/`ChannelLatest`) may transiently hold the older of two
  overlapping values until the next message. The per-channel dedup marker is
  deliberately kept aligned with the payload delivered LAST (it is re-stamped
  when a later-finishing delivery notices that another callback replaced it),
  so an inversion can never make an identical republish of the current value be
  suppressed and turn the transient inversion into permanent staleness.
  Ordering-sensitive feeds must still carry a timestamp/sequence field and
  compare it in the formula; polled reads (`GET`/`HGET`/`HGETALL`) are single
  round trips per tick and are unaffected.
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
  pure JSON conversions (`RedisUDFMatrixToJSON`/`RedisUDFJSONToMatrix`) are
  non-volatile. Keep a reflection signature-parity test for every pair.
- Write delivery is configurable (`SyncWrite`): `sync` blocks for every reply;
  `fireforget` (default) sends result-agnostic writes
  (`Set`, `SetJSON`, `SetKV`/`SetKVPair`, `SetEx`, `Rename`, `HashSet`,
  `HashSetMultiple`, list pushes, channel publishes) with
  `CommandFlags.FireAndForget` and returns `OK (fire and forget)`, while
  reply-dependent writes (`Del`, `Incr`, `IncrBy`, `Expire`, `SetAdd`,
  `SetRemove`, `HashDel`, list pops) stay blocking; `fireforget-all` sends
  every write FireAndForget (reply-dependent writes return
  `OK (fire and forget: all)`). `ChannelUnsubscribe` always removes the local listeners
  deterministically (never fire-and-forget). In fire-and-forget modes the reply
  is DISCARDED by StackExchange.Redis, so a write the server rejects (WRONGTYPE,
  OOM, quota) is neither awaited nor detected: the cell keeps the marker and
  nothing is logged for it. Only validation/config failures BEFORE the dispatch
  surface as `Error:`, and publishes return the marker instead of the readers
  count. The marker means "sent", not "accepted" - use `SyncWrite: "sync"` when a
  write's success must be observable.
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
  `OnNext`/`OnCompleted`, so a large same-host burst cannot throttle the
  pool (a classic per-item pool-thread dispatch would block one thread per
  queued item). The async identity is the calling cell + resolved host +
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
  only), and shutdown does not block on closing connections.
- Native Redis HA (Sentinel and Cluster) works through StackExchange.Redis:
  a `host` that is a Sentinel connection string (`serviceName=`) or any node of
  a Redis Cluster connects, follows MOVED/ASK, re-subscribes after a failover,
  and runs every key-routed command/read/write and Pub/Sub. Server-wide
  operations span every master: `RedisUDFKeys` (SCAN) and
  `RedisUDFPubSubChannelsInfo` (PUBSUB CHANNELS/NUMSUB, subscriber sums) iterate
  the cluster masters (`GetMasterServers`) instead of the first node; a Redis
  Cluster has only database 0, so a `defaultDatabase != 0` connect is rejected
  with a clear error; a multi-key command that spans shards (MGET/MSET/RENAME)
  fails with the client's "single slot" error plus a hash-tag hint. Standalone
  and Sentinel keep the single-endpoint behavior.
- Cached multiplexer eviction (past the 512-entry cap) vetoes hosts with live
  subscription listeners (`RedisConnectionManager.SetEvictionProtection` wired
  to `RedisSubscriptionManager.HasActiveSubscribers`), reclaims entries whose
  Lazy factory never materialized, and invalidates the cached
  `IDatabase`/`ISubscriber` wrappers of the evicted multiplexer so later calls
  rebuild instead of throwing `ObjectDisposedException` forever.
- A connect-failure memo self-extends only on a real new attempt: the
  fast-fail skip does not refresh the timestamp, so a host polled faster than
  the 2 s window still retries once the window expires. A burst of `SUB`
  topics or volatile reads against a dead host pays one connect attempt instead
  of N x ConnectTimeout; a successful connect clears the memo, and the retry
  loops/recalculations recover after the window. Eviction clears the memo and
  the idle-down clock of the evicted host.
- Every pipelined batch wait is bounded (`RedisUDF.WaitBounded` with the
  configured response timeout + 500 ms): an orphaned batch task (multiplexer
  disposed mid-execute) surfaces as a timeout instead of stalling a timer or
  the Excel thread forever. Covers RTD polled reads and the UDF batch
  `...Multiples` functions.
- `CoalesceRealtimeUpdates`/`ConflationMs`: in real-time mode a topic that keeps
  changing is pushed to Excel at most once per window with the latest value
  (latest wins), instead of one push per message. The root `ConflationMs`
  (ms, `0` = off, capped at 3600000) is the explicit window; when it is absent
  the boolean decides - `true` = the Excel tick interval, `false` = per-message.
  This is the main lever against screen
  flicker and the recalculation cascade of dependent formulas on fast feeds; a
  window below ~50-100 ms barely helps because the flush rides the Excel tick.
- `RedisRuntime.ResetAfterAddInReload` supports a same-process add-in reload
  without reusing the previous managers.
- Values and identifiers (keys, hash keys, fields, channels, patterns) written
  to Redis always use the invariant culture (decimal point), regardless of the
  Excel locale.
- The host argument may be a full connection string, so it can carry
  `password=...` / `user=...`. Every log line and every verbose error cell masks
  credentials (`AppConfig.MaskHost`: `password=****` / `pass=****` /
  `user=****` / `username=****`); host/port/options stay readable. This covers
  the UDF/RTD error text, the `#ERROR` RTD cell, the `RedisRTDDefaultHost` cell
  and the `ConnectData` argument log.
- Excel error cells are rejected in every scalar argument position; date/time cells
  arrive as their Excel serial number and boolean cells serialize as
  `true`/`false`.
- An empty-string cell is a valid Redis name/pattern while a truly blank cell
  is a missing argument.
- 2x2 pair ranges are read as 2 rows x 2 columns (vertical pairs).
- Identical consecutive payloads are skipped before decoding for literal
  subscriptions and GET/HGET polling (`SkipRepeatedMessages`, default on):
  payloads are compared by ORDINAL text, so `"007"` and `"7"` are different
  values (a zero-padded id must not be swallowed). RedisValue's numeric equality
  is NOT used: it would collapse `"7"`/`"007"`/`"+1"` and leave the cell stale,
  and it disagreed with the write side (`PublishIfChanged` already compares
  ordinals); PSUB patterns are never deduplicated because channels interleave.
  Unchanged HGETALL hashes are compared field-by-field (ordinal) and skipped too.
- The config file is read once per Excel process (restart Excel after editing
  it) and searched in: the add-in folder, the user profile, the Excel folder and
  `C:\Windows` (first existing wins). The file may contain `//` and `/* */`
  comments and is validated against the embedded JSON Schema (`ConfigSchema.cs`)
  in a fail-closed way: a wrong **value** on a known key rejects the **whole
  file** (every UDF/RTD cell shows `#CONFIG: ...` until it is fixed), while an
  extra/unknown **key** is only a log warning. The removed v1 keys
  (`RTD.host`/`UDF.host`/`RTD.timeout`/`UDF.timeout`, root
  `CoalesceRealtimeUpdates`) are errors, so a pre-1.6 file fails closed instead of
  silently reverting to localhost; a `Default` that is a bare word which is not a
  defined `Servers` name and not a recognizable connection string is a WARNING
  only (it may be a valid port-less host like `localhost`), used as the host. A
  structural failure (not JSON, or a root that is not a JSON object) also fails
  closed (`#CONFIG`), and the first-existing candidate stops the search. Every
  cell's `#CONFIG` reason names the offending **file path**, so a stale copy in
  one search location is identifiable from the cell. `=RedisCheckConfigFile()`
  reports the problems in a cell (blank =
  the config in use, or pass a path). When no config file is found at all, the
  no-file default for `ExcelUpdateRateMs` applies (1000 ms, vs the loaded-file
  default of 100 ms). An undefined `ExcelUpdateStyle` falls back to `Automatic`,
  and `Servers` names are matched after trimming and case-insensitively (exact
  match first). **Upgrading:** the installer keeps an existing `RedisExcel.json`
  the add-in accepts, and for one it would reject (pre-1.6 shape or a wrong value)
  it saves `RedisExcel.json.pre-1.6.bak` and writes a working file - reading the
  old address first (`read-config-host.ps1`) so the replacement does not silently
  become `localhost`; the wizard's host field is pre-filled with it and a silent
  install reuses it. The installer also **unregisters an older copy** of the
  add-in: any `OPEN`/`OPENn` value in `Excel\Options` that points at a *different*
  RedisExcel XLL is deleted (plus dead "Add-in Manager" labels left by a renamed
  or moved copy, the `redisd`-style line in File > Options > Add-ins), because
  otherwise Excel loads two copies - duplicated functions, one of them reading a
  different `RedisExcel.json`. Other add-ins' entries are never touched, and this
  install's own entry is registered after the cleanup.
- The installer's host page has a **Test connection** button (same authenticated
  probe as the end-of-install check, run without leaving the page or writing the
  config). Its message comes from the helper's STABLE exit codes (0 ok, 2 unknown
  host, 3 timeout, 4 refused/other socket error, 5 AUTH rejected, 6 unexpected
  reply) and never from the helper's console text, which a `cmd.exe` redirect
  re-encodes into the OEM codepage and would mangle. `RedisUDFVersion` is the
  in-Excel counterpart: it reports the build and the XLL path actually loaded, so
  a duplicated install is visible from a cell too.

## Build

```powershell
dotnet build RedisExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`:

- `RedisExcel-packed.xll` (32-bit) and `RedisExcel64-packed.xll` (64-bit),
- `RedisExcel.dll`, plus the dependency assemblies are packed inside the XLLs.

CI (`.github/workflows/build.yml`) runs on pull requests and main pushes; on
`v*` tags it passes `/p:InformationalVersion=<tag>`, which is embedded in the
Redis `ClientName` for diagnostics (shown as `dev` for local builds). On a tag it
also: verifies the CHANGELOG top heading matches the tag (anchored, so a short
tag cannot pass against a longer heading), asserts each of the five release
assets exists (an upload with `if-no-files-found: error` only fails when the whole
set is empty), and uses that CHANGELOG section as the release body. The release
job checks out the repository so it can read `CHANGELOG.md` (the checkout runs
before the artifact download so it cannot wipe the downloaded assets).
The build job also installs Inno Setup and compiles `installer\RedisExcel.iss`
on every run (a broken `.iss` fails the build, not only a tag), uploads the
installer (`RedisExcelSetup-<version>.exe` plus the version-less
`RedisExcelSetup.exe` copy, which is what makes
`releases/latest/download/RedisExcelSetup.exe` a link that never goes stale) as a
standalone artifact, and on a `v*` tag adds both to the release assets.
The smoke step retries once (2 attempts, 3s apart) and the quick liveness step up
to three times (5s apart): a shared runner can stall the whole process or trip a
best-effort race on infrastructure noise, and the bounds themselves are
unchanged, so a real regression fails every attempt. The unit step has no
retry (a unit regression cannot be masked).
The Redis endpoint the suites run against comes from the **portable Windows build**
(a direct download from the redis-windows releases), with the Memurai Chocolatey
package as the fallback: the community feed flakes, and the choco path used to be
primary, so every run paid for its retries before falling back. Both paths are
fail-closed (the step throws when neither provisioned a server).

## Tests

### 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

Covers: `ExcelJson` conversions and the matrix size/total-cell budget,
`AppConfig` load/sanitize (including alias trim/case-insensitivity, the
undefined-`ExcelUpdateStyle` reset, the JSON-Schema per-value validation/drop and
the comment-tolerant parse), `ResolveHostCore` (exact-match priority), the
real-time conflation window
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
return true, a never-completing task reports the timeout), the public-surface
baseline parity (the shipped assembly against the frozen 79-function surface and
the linked sources, plus the twin/signature pins), and the unbounded-blocking
scan of the root sources.

The unit, smoke, load and liveness test projects compile the production
sources directly (linked `Compile` items) from ONE shared list,
`test\TestSupport\ProductionSources.props`; a new production `.cs` is added
there ONCE and every suite picks it up (a per-project copy drifted - the round-K
LoadTests defect). `ProjectFileParityTests` enforces that every top-level
production source is listed there (except the intentionally excluded
`RedisRtd.cs`) and that no project declares production sources itself.
`PublicSurfaceParityTests` freezes the public Excel surface in
`test\RedisExcel.Tests\PublicSurface.baseline.txt` (refresh deliberately with
`REDISEXCEL_REFRESH_PUBLIC_SURFACE=1`), and `NoUnboundedBlockingTests` scans the
root sources for an unbounded wait (`.Result` / `.Wait(` /
`GetAwaiter().GetResult()`) outside its allow-list. `dotnet test` does not build
or pack the add-in; CI builds it with msbuild.
`test\TestSupport\Run-Tests.ps1` builds once and runs the layers locally.

### 2. Smoke tests (requires a Redis server, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
```

Covers the ref-counted Pub/Sub behavior (two listeners on one channel, the
last-listener unsubscribe, re-subscription, duplicate suppression, origin
counters, argument validation) plus the concurrent-dedup check, the real
`SyncWrite` modes and publish markers against Redis, a live
`RedisUDFUpdateAvailable` check and the AV banner. See `test/README.md` for the
full list.

### 3. Excel end-to-end

```powershell
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Uses a hidden Excel instance: loads the packed XLL, builds UDF/RTD test sheets,
asserts values against Redis, then covers the multi-workbook case: opens a COPY
of the workbook, publishes messages, closes the copy
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

Environment control: `test\Set-Antivirus.ps1 -Action Status|Disable|Enable` reports
the AV state (Windows Security Center + AV-looking services) and can turn
Defender real-time off/on (`-WhatIfOnly` dry-runs it). A third-party engine with
self-defense/tamper protection cannot be paused by any generic script (the
service refuses to stop and the process cannot be killed — verified: Kaspersky
home refuses `Stop-Service` and even `taskkill /F`, and its `STOP Protection`
CLI hangs waiting for a password/GUI prompt), so for the AV-off comparison the
operator pauses it from the product UI/tray and the harness records the state;
note that pausing/exiting one AV usually lets Defender re-engage, so "off" often
means Defender-on, which the status output shows. To compare the two states,
`test\Run-ExcelE2E-Matrix.ps1 -Tag <name>` runs the four write modes and
`test\Compare-ExcelE2E.ps1 -TagA <on> -TagB <off>` prints them side by side
(result, COM-retry total, skip count per mode).

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

1. **Excel COM automation is inherently flaky; antivirus/EDR only adds to it.**
   Measured on this repo: even with **no** antivirus, every `Range.Value2`
   write is rejected once (`RPC_E_CALL_REJECTED`) and the `-AsyncWrites` mode
   needs one retry ~180 times (Excel is busy computing), so a retry count is
   NORMAL, not a red flag. AV on vs off showed the same counts in this
   environment; security software widens the window for a *flake* (e.g. a
   sub-second timing check), it is not the source of the retries. Symptoms of
   COM interference are misleading errors (`InvalidCastException: cannot
   convert Int32 to String`, `RPC_E_CALL_REJECTED`, COM calls hanging or
   aborting); the E2E retries with backoff and has a warm-up phase, so **if a
   run fails this way, run it again**. The script prints the AV products
   (Windows Security Center — real-time state is the productState **bit 12**:
   `0x41000` = on, `0x42000` = paused) before creating Excel and a
   `COM retries: total=...` line at the end. When the interference exceeds the
   retry budget, the failure names the active product. Never assume such an
   error means the add-in is broken. `test\Set-Antivirus.ps1` reports the AV
   state and can turn Defender real-time off/on, but a **third-party product
   with self-defense/tamper-protection cannot be disabled by any generic
   script** — the service refuses to stop and the process cannot be killed;
   pause it from its own UI/tray (the helper says so) and use its `Status`
   output to record which state a run had. Every suite prints the same AV
   banner (`AntivirusInfo.Describe()`: unit via a test output line, smoke/load
   as `Antivirus: ...`, liveness as `[LIVENESS] antivirus=...`), so a run's
   result is always attributable to its environment.
2. **Elevated processes ignore per-user COM classes.** The add-in registers
   `RedisRtd` under `HKCU\Software\Classes` in `AutoOpen`. If Excel (or the
   whole automation session) runs elevated, `=RTD("RedisRtd", ...)` shows
   `#N/D` and `CLSIDFromProgID('RedisRtd')` fails with `0x800401F3`, even
   though the registry keys exist. Workaround for test machines: register the
   ProgID/CLSID under `HKLM\Software\Classes` (admin), pointing at the packed
   XLL. This is an environment workaround, not part of the normal install.
   A **stale** `HKLM` registration is itself a trap: it makes Excel load a
   *second* copy of the XLL from a different folder, whose RTD server reads
   that other folder's `RedisExcel.json` and never sees the config the run
   under test wrote. In that state `=RTD("RedisRtd",...)` can return a value
   while the UDFs in the tested XLL return `#CONFIG` - the fail-closed guard is
   **per XLL instance**, not per process, so a machine-wide ProgID can hide it.
   Confirm only ONE XLL module is loaded (`Get-Process EXCEL | Get Modules`)
   before trusting an RTD result.
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
6. **RTD cells display `#N/D` when the RTD server could not be created** -
   check COM registration before suspecting the add-in logic.
7. **A sibling Excel-automation suite may be under test at the same time**
   (`../natsexcel` here, and its E2E also kills its own hidden Excel). Its
   instances - hidden `/automation -Embedding` and the Office `/restore` relaunch
   its crash test triggers - are indistinguishable from ours by command line, so
   the E2E identifies ITS victim by the **XLL module loaded**
   (`RedisExcel-packed.xll` / `RedisExcel64-packed.xll`), never by "an EXCEL.EXE
   outside the pre-run baseline", and auto-closes a `/restore` only when it
   appeared within 20s of its own kill. Everything else is reported, never closed:
   killing a concurrent run's instance would break that run (and closing the wrong
   recovery pane can hide a real crashed-document list). When Excel processes keep
   appearing every ~15-30s with no XLL loaded, the other repo is being tested -
   not a leak of ours.
8. **A `redis-cli.exe` must be run from the folder that carries its DLLs.** The
   Windows msys2 build needs `msys-*.dll` beside the executable: a copy taken out
   of that folder exits `0xC0000135` (STATUS_DLL_NOT_FOUND) and prints nothing, so
   the E2E reported "Redis is not responding" against a server that was answering
   PONG. The message now names the CLI and its exit code
   (`CLI 'redis-cli ' exit=-1073741515`), and `-RedisCli "docker exec <container>
   redis-cli"` sidesteps the local CLI entirely.
9. **The installer's interactive flow can be verified without stealing focus.**
   `SetForegroundWindow` is refused to a background process (and
   `AttachThreadInput` fails when the current foreground window belongs to a
   higher-integrity process), so keystrokes would land in whatever the user has
   focused. Capture each page with `PrintWindow` (it renders an unfocused window)
   and advance by posting `BM_CLICK` to the wizard's own button controls. Inno's
   option pages are ONE owner-drawn `TNewCheckListBox` - the items have no child
   windows and it is not a Win32 LISTBOX - so select an option by posting
   `VK_DOWN` to that list. Two traps: the installer and the uninstaller
   re-launch themselves from `{tmp}` (`RedisExcelSetup-*.tmp`, `_unins.tmp`), so
   track the window rather than the launched process, and the uninstaller does
   nothing until its confirmation is answered Yes.

## Release

Push an annotated `v*` tag; the CI workflow builds and
publishes the packed XLLs, `RedisExcel.dll`, `NLog.config`, `RedisExcel.json` and
the installer (`RedisExcelSetup-<version>.exe` plus the version-less
`RedisExcelSetup.exe`, the stable direct-link name) as release assets. Current
release: `v1.7.2` (the installer's end-of-setup console holds for 3s so its banner and
result can be read - on top of the `v1.7.1` probe that reads instead of PINGing, so a
server that refuses writes is no longer reported as dead - on top of the `v1.7.0`
fix for a Pub/Sub cell that could stay on an older value after a
transient out-of-order delivery, and the installer's connection test now drives
StackExchange.Redis itself instead of a raw `PING`; on top of the v1.6.4
`RedisUDFVersion`/stale-install cleanup, the v1.6.3 config migration and the
v1.6.0 named-server configuration - see the root `CHANGELOG.md`).

On a version bump: (1) the CHANGELOG top heading must be `## v<tag>` (the CI
guard fails the release otherwise and names this line); work in progress keeps an
`(unreleased)` suffix on that heading, which the tag step removes. (2) update the
`Current release` value above, and (3) if the Excel surface changed, refresh the
frozen baseline - `PublicSurfaceParityTests.BaselineHeader`/the generated
`test\RedisExcel.Tests\PublicSurface.baseline.txt` - with
`REDISEXCEL_REFRESH_PUBLIC_SURFACE=1` and review the diff. The landing pages
(`docs\`) carry no version, so they only change when their text changes.

## Shell for agents

This is a **Windows** project (.NET Framework 4.8, Excel-DNA, Inno Setup), so the
tools that matter are Windows ones and they are all reachable from PowerShell:

- **PowerShell is available and preferred over bash.** Build with `dotnet build
  RedisExcel.sln -c Release`, test with `dotnet test test\RedisExcel.Tests\...`,
  compile the installer with `& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
  "/DAppVersion=<ver>" "installer\RedisExcel.iss"`, and drive/inspect the
  registry with `Get-ItemProperty` / `Set-ItemProperty` (Excel's auto-load list
  lives in `HKCU:\Software\Microsoft\Office\<ver>\Excel\Options`). None of this
  needs a POSIX shell.
- **When a bash-style command is unavoidable, prefer the PowerShell equivalent.**
  `grep` -> `Select-String`, `sed`/`awk` -> `-replace`/`-split`, `head`/`tail` ->
  `Select-Object -First/-Last`, `which` -> `Get-Command`, `ls -R` ->
  `Get-ChildItem -Recurse`. Piping between them works the same way.
- **Quoting differs between the two shells** and silently corrupts output: in
  bash a heredoc (`<<'EOF'`) and `$(...)` work; in PowerShell they do not
  (`<<` is reserved and `$(...)` is subexpression syntax). Write the body to a
  file and pass `-File`, or use a `.ps1` script, instead of inlining a long
  multi-line command.
- **Never rewrite a text file with `Set-Content`/`Out-File` just to change a few
  bytes.** Both re-encode the file, which corrupted this repo's `.iss` (mixed
  ASCII/CJK/accented text) during development. Use the editor's targeted
  replace, or read/write with an explicit encoding:
  `[IO.File]::WriteAllText($p, $text, (New-Object Text.UTF8Encoding($false)))`.
  Check with a strict decode (`New-Object Text.UTF8Encoding($false, $true)`)
  when in doubt: the console showing `?` does NOT mean the file is broken, but
  a failed strict decode does.
- **Console output is not a reliable view of the text.** The Windows console
  renders in an OEM codepage (437/850/866/932), so accented and CJK characters
  appear as `?` or mojibake even when the file is perfect. Verify non-ASCII text
  by code point instead of trusting what a `Write-Host`/redirect shows.
- **Locating windows and automating dialogs** works through `user32.dll`
  (`EnumWindows`, `PrintWindow`, `PostMessage`): `PrintWindow` renders an
  unfocused window, so a wizard/screenshot check never steals focus. These
  P/Invoke types do not survive between separate `powershell.exe` invocations -
  run the whole capture in ONE script.
- **Long/hanging commands**: an Excel COM call or a dialog can block forever.
  Bound it (`Start-Job` + `Wait-Job -Timeout`, or a script with an explicit
  timeout) instead of leaving the shell waiting. Never leave an automation Excel
  (`EXCEL.EXE /automation -Embedding`) behind - but never kill a user's Excel
  either: identify your own by the XLL module loaded / the `/automation` flag.
