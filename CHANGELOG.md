# Changelog

## v1.2.7 (unreleased)

### Fixed

- `RedisUDFChannelPublishIfChanged` only remembers a payload that was actually
  delivered and clears the marker when any listener (RTD `SUB`/`PSUB` or UDF
  `ChannelLatest`) joins or rejoins, with glob matching for pattern
  subscriptions, so a late subscriber is no longer starved by a publish it
  never saw; the cache is safe under concurrent recalculation and LRU-capped
  (`PublishDedupCacheSize`).
- Empty/whitespace keys are valid Redis names in the batch writers (`SetKV`,
  `SetKVPair`, `HashSetMultiple`) instead of being silently dropped; a missing
  required key/field returns a friendly `Error:` cell.
- RTD: an invalid host is surfaced in the cell instead of being hidden; a
  failed `SUB`/`PSUB` subscribe is retried with a 1s..30s backoff (capped per
  tick) and the first failure logs once with the exception; blank `SUB`/`PSUB`
  channels are rejected with a clear `#ERROR`; a missing hash returns `{}` from
  `HGETALL` (was `(no value)`); `RedisRTDMessagesCounter` reports the last full
  second; extra topic arguments are rejected; polling runs per host in
  parallel; `HGETALL` comparison is O(n) and compares field names byte-exact;
  topic logs truncate payloads to 64 chars (full payloads only at Trace level).
- `RedisConnectionManager` shutdown no longer blocks on closing connections and
  the connection counters report live multiplexers per pool (the RTD
  connection count is `RtdData` + `RtdSub` only); the RTD status helpers
  return `0`/`false` after shutdown.
- The update check keeps the last known tag when a GitHub response lacks
  `tag_name`, and retries sooner after a failed attempt instead of waiting the
  full refresh window.
- `ExcelJson`: a UTF-8 BOM is tolerated even with whitespace before it, JSON
  nesting depth is capped at 256, and oversized matrices are rejected. Date-like
  JSON strings stay text (as since v1.2.6); only programmatic `DateTime` values
  now serialize as ISO-8601 text.

### Changed

- New config key `PublishDedupCacheSize` (default 10000): LRU cap for the
  `PublishIfChanged` publish-dedup cache.
- `RedisRuntime.ResetAfterAddInReload` resets the process-wide managers for a
  same-process add-in reload.
- Release tags are annotated from v1.2.7 on (the `v*` tag trigger is unchanged).

### Tests

- Unit coverage for the connection/subscription managers, config loading and
  the publish-dedup LRU cache, plus new E2E regression rows.

## v1.2.6

### Fixed

- A listener that attaches to an already-active Pub/Sub channel now receives
  the next payload even if it is identical to the last one (the per-channel
  dedup could starve late joiners).
- `RedisUDFChannelPublishIfChanged` actually suppresses unchanged payloads
  (returns `No change` instead of publishing again).
- `RedisUDFChannelUnsubscribe` only removes the given host's listeners (new
  optional host argument; omitted = default host) and rejects blank channels;
  publish/latest reject empty channels with a clear message.
- `RedisUDFSetEx` validates the TTL like `RedisUDFExpire` (0/negative no longer
  reaches the server).
- `TTLMultiples` reports fractional seconds like `TTL` (was rounded).
- Batch writers with no valid entries return an `Error:` cell instead of a
  silent `OK`; null range arguments return `Error: a range is required`.
- `RedisConnectionManager` wrappers respect the shutdown fence; a malformed
  high-priority config file no longer falls through to a lower-priority file;
  malformed hosts produce a stable English message.
- RTD: an existing empty value no longer displays as `(no value)`; the
  `(ConnectData)` placeholder stays until the first real value.

### Changed

- `RedisUDFChannelUnsubscribe` now scopes removal to a host: sheets that
  subscribed with a non-default host must pass the host to unsubscribe
  (previously it removed every host's listeners for that channel).

### Tests

- Smoke late-joiner dedup case; unit + E2E coverage for the new validations
  (rows 44-45).

## v1.2.5

### Fixed

- Batch writers (`SetKV`, `SetKVPair`, `HashSetMultiple`) validate their range
  shapes instead of silently writing only part of the data; `SetKVPair`/
  `HashSetMultiple` also accept horizontal 2-row ranges.
- `RedisUDFChannelUnsubscribe` rejects a blank channel instead of reporting
  success for nothing.
- `RedisUDFGetMultiple`'s `multipleColumns` accepts TRUE/FALSE, 0/1 and
  TRUE/FALSE text, and rejects anything else instead of silently meaning FALSE.
- `ToInt64Invariant` rejects booleans and fractional numbers (TTL/increment
  arguments must be whole numbers).

### Tests

- Unit + E2E coverage for the new validations (rows 42-43).

## v1.2.4

### Fixed

- Scalar UDF arguments reject multi-cell ranges and array constants (they
  previously became the literal text `System.Object[,]`, silently creating keys,
  values or channels).
- Excel error cells are rejected in every scalar UDF argument position: they
  previously became `null`/`""` for values and the raw error code (`#N/A` = 42,
  `#DIV/0!` = 7, ...) for TTL, increment and index arguments. Documented
  exception: matrix cells passed to `RedisUDFSetJSON`/JSON publish still map
  error cells to JSON `null` (unchanged, pre-existing behavior).
- Optional host arguments accept only text: blank/missing keeps the default
  host, any other type surfaces an error instead of silently using the default.
- Multi-key UDFs (`GetMultiple`, `ExistsMultiples`, `TTLMultiples`,
  `HashGetFieldMultipleKeys`) flatten any range shape row-major; a range with
  more than one row and column used to silently query only its first row.
- `RedisUDFKeys` rejects a missing or empty-string pattern instead of scanning
  every key; pass `"*"` to match all keys.
- `RedisUDFJSONToMatrix` rejects non-finite JSON numbers (`NaN`, `Infinity`,
  `-Infinity`) with an `Error:` cell (Excel rendered them as `#NUM!`).
- Numbers written to Redis round-trip exactly: values whose compact form would
  parse back to a different double (e.g. `double.MaxValue`) now use full
  precision.
- Boolean cells serialize as `true`/`false`, matching the JSON path (was
  `True`/`False`); this applies to identifiers too - a boolean cell used as a
  key name changes from `True` to `true`.
- `RedisUDFGetMultiple` returns a single `Error:` cell when no valid key
  remains; a rejected RTD topic no longer logs a fake ERROR at disconnect.

### Changed

- Documented behaviors: date/time cells are stored as their Excel serial number
  (use `TEXT()` for a date string); missing-value sentinels differ per surface
  (`Get` returns an empty string, `GetMultiple` `(null)`, RTD `(no value)`);
  `GetMultiple` skips blank keys while the other multi-key functions keep one
  row per input.

### Tests

- Unit tests pin the new conversion contracts (multi-cell range and error-cell
  rejection, boolean casing, double round-trip, non-finite JSON).
- Unit tests also cover the JSON argument/fill behavior (matrix error cells map
  to `null`; empty or ragged JSON uses the fill value), the `GetMultiple`
  validation messages and Excel missing/empty sentinel handling.
- E2E regression checks for the multi-cell-range rejection, the no-valid-key
  error, 2x2 row-major flattening and the blank keys pattern.

## v1.2.3

### Fixed

- `RedisUDFHashGetFieldMultipleKeys` with an empty input returns a blank cell
  instead of a zero-width matrix (completes the v1.2.2 zero-row fix).
- RTD `ConnectData`/`PollHost` reject only null keys/fields again - empty and
  whitespace key/field names are valid Redis names (v1.2.2 had over-rejected
  them).
- RTD topic value updates and dirty flushes are serialized per topic, so a tick
  flush can no longer overwrite a newer value with an older one; a failed push
  stays dirty for the next tick.

### Changed

- The E2E script sanitizes the saved sample workbook (removes the local
  `absPath` and personal document metadata) before it can be committed.
- E2E gained regression checks for the RTD argument validation (a topic with a
  missing key returns `#ERROR` without disturbing the host's other topics) and
  for a whitespace-only key.
- Removed the dead `InternalsVisibleTo`; docs note the linked-sources test
  model.

## v1.2.2

### Fixed

- `RedisUDFSetRemove` removes the empty member for null/empty member cells (a
  null member made the client throw `A null value is not valid in this
  context`); this matches `RedisUDFSetAdd`.
- `RedisUDFMatrixToJSON`: cells of exactly 2^63 no longer overflow to a
  negative long (the boundary check is strict now).
- `RedisUDFJSONToMatrix`: JSON integers beyond Int64 are returned as invariant
  text instead of `#VALUE!`; ISO date strings are kept as text
  (DateParseHandling.None) instead of becoming Excel serial numbers; empty
  inner arrays (`[[]]`) return the fill value instead of a zero-width matrix.
- RTD `ConnectData` validates the required arguments per command (`GET`/
  `HGETALL` need the key, `HGET` needs key and field, `SUB`/`PSUB` need the
  channel) and returns `#ERROR: ConnectData: ...` instead of registering a
  topic that would later abort the host's polling; `PollHost` also skips
  malformed topics defensively.
- Multi-key read functions (`ExistsMultiples`, `TTLMultiples`,
  `HashGetFieldMultipleKeys`) treat blank cells as empty keys and keep row
  positions instead of failing the whole call.
- Zero-row results (`Keys` with no matches, `HashGetAll` on a missing hash,
  empty multi-key inputs) return a blank cell instead of a zero-width matrix
  that Excel renders as `#VALUE!`.
- The JSON publish wrappers no longer publish an `Error: ...` string when the
  matrix-to-JSON conversion fails, and `RedisUDFMatrixToJSON` uses the
  exception message only.
- Dirty topics are flushed on every Excel tick in every mode, so a queued value
  can no longer be stranded when the `Automatic` style re-enables real-time
  updates.

### Changed

- `HashEquals` documents its unique-field-name precondition; `TickGate.Exit`
  doc clarifies it is only safe after a successful `TryEnter`.
- Test loop is much faster: the unit test project compiles the production
  sources directly instead of building the add-in (no ExcelDna packing per
  run), and the E2E polls at 120 ms with shorter fixed sleeps (suite ~28 s,
  down from ~2 min).

## v1.2.1

### Fixed

- Null/empty values are now stored as `""` on every write path (key-value and
  matrix setters, list pushes and channel publishes included); previously some
  of them reached Redis as a null `RedisValue` (deleting the key or throwing).
- `RedisUDFJSONToMatrix` returns the fill value for empty/missing input cells
  instead of an error cell.
- A failed Redis connect is now replaced with a fresh entry atomically
  (compare-and-swap), so the next call retries and a concurrently created
  entry is never removed.

### Changed

- Test hardening: invariant-culture tests run under a comma-decimal culture
  (de-DE) and require the invariant parsing behavior; `TickGate` gained a
  free-gate contention test; the smoke suite adapts when
  `SkipRepeatedMessages` is disabled by a local config; the E2E script got
  safer cleanup/save paths (crash-safe sample replace) and better failure
  diagnostics.
- CI release step declares `permissions: contents: write` explicitly.

## v1.2.0

### Changed

- Multi-key worksheet functions (`RedisUDFExistsMultiples`,
  `RedisUDFTTLMultiples`, `RedisUDFHashGetFieldMultipleKeys`) now pipeline all
  keys in a single round trip instead of one command per key.
- `IDatabase` and `ISubscriber` wrappers are cached per host/pool, removing an
  allocation from every UDF call.
- Unchanged `HGETALL` hashes are compared field-by-field and no longer
  re-formatted or pushed to Excel (extends `SkipRepeatedMessages` to hash
  polling).
- Real-time updates are coalesced per topic by default
  (`"CoalesceRealtimeUpdates": true`): in real-time mode each topic sends at
  most one value per `ExcelUpdateRateMs` window (the latest value wins)
  instead of one Excel update per incoming message. Set it to false to
  restore per-message delivery.

### Added

- New worksheet functions:
  - `RedisUDFDel(key, optionalHost)` - deletes a key and returns the deleted
    count.
  - `RedisUDFSetEx(key, value, ttlSeconds, optionalHost)` - set with TTL.
  - `RedisUDFExpire(key, ttlSeconds, optionalHost)` - sets the TTL of a key.
  - `RedisUDFIncr(key, optionalHost)` / `RedisUDFIncrBy(key, increment,
    optionalHost)` - atomic counters.
  - `RedisUDFListPushRight(key, value, optionalHost)` /
    `RedisUDFListPushLeft(key, value, optionalHost)` - list push (returns the
    list length).
  - `RedisUDFListRange(key, start, stop, optionalHost)` - list range.
  - `RedisUDFSetAdd(key, value, optionalHost)` - set add (returns the number
    of members added).
  - `RedisUDFSetMembers(key, optionalHost)` - set members.
  - `RedisUDFHashDel(hashKey, field, optionalHost)` - deletes a hash field.
  - `RedisUDFSetRemove(key, value, optionalHost)` - removes a set member.
  - `RedisUDFListPopRight(key, optionalHost)` /
    `RedisUDFListPopLeft(key, optionalHost)` - pops a list element.
  - `RedisUDFType(key, optionalHost)` - returns the key type.
  - `RedisUDFRename(key, newKey, optionalHost)` - renames a key.
  - `RedisUDFKeys(pattern, optionalHost, pageSize)` gained an optional
    `pageSize` argument (SCAN page size).

### Robustness

- Timer reentrancy gates: a slow poll/update tick no longer overlaps the next
  one; `TickGate`, a tested helper, implements the guard.
- RTD polling isolates GET-multi failures per host, so `HGET`/`HGETALL` topics
  still run in that tick.
- Failed Redis connects are no longer cached: the failed entry is dropped and
  retried on the next call; a concurrently created entry is restored.
- RTD polling commits the dedup state only after the value is accepted for
  delivery (queued or pushed without error); a failed Excel push restores the
  dirty flag so the value is retried on the next tick. `ServerTerminate` marks
  topics disconnected before stopping the timers.
- Subscription manager hardened: network I/O moved outside the per-channel
  lock, a manager dispose flag, safe channel removal (a stale removal restores
  the currently installed entry instead of evicting it) and
  host-length-prefixed subscription keys (no collision when hosts/channels
  contain control characters); a late unsubscribe can no longer tear down a
  freshly installed handler (the network gate is held across Unsubscribe).
- Subscription manager shutdown race closed: a channel state created
  concurrently with `Dispose()` can no longer subscribe after shutdown.
- Joiner retry: every `Subscribe` re-ensures the single StackExchange.Redis
  handler (idempotent fast path), so a failed subscribe followed by another
  listener no longer leaves a channel silently unsubscribed.
- Subscription listeners carry an origin tag ("RTD"/"UDF"); the RTD status
  counters (`RedisRTDSubscriptionCount`, `RedisRTDChannelCount`) now report RTD
  listeners only.
- `RedisRuntime` publication order fixed (connections/subscriptions no longer
  observable half-initialized).
- `HGETALL` comparison is now order-insensitive: a rehash that reorders fields
  no longer triggers a pointless Excel update.

### Tests / CI

- CI now installs Memurai (Redis for Windows) on the runner and executes the
  smoke suite before the tagged build.

### Fixed

- Numeric cell values are written to Redis with the invariant culture
  (`67000.5`, not `67000,5` on comma-decimal locales), covering single sets,
  matrix/key-value setters, hash fields, channel publishes and list pushes.
  Identifiers (keys, hash keys, fields, channels, patterns) and values are
  converted with the invariant culture. Null/`ExcelMissing`/`ExcelEmpty`/
  `ExcelError` cell values become null and are never sent as a Redis key;
  empty/null values are stored as `""` for key and hash setters (they no longer
  reach Redis as a null `RedisValue` that would delete the key).
- `RedisUDFExpire` rejects non-positive TTLs; `RedisUDFType` reports `unknown`
  for unrecognized key types; `RedisUDFJSONToMatrix` accepts numeric cells
  invariantly.

## v1.1.3

### Changed

- Subscription broadcast hot path optimized: listeners are kept in a
  copy-on-write snapshot, so each incoming message no longer allocates a list
  copy from the `ConcurrentDictionary`. Measured against a live server
  (30s, `PSUBSCRIBE *`): broadcast overhead per message dropped from ~212 B to
  ~68 B (total allocation per message ~30% lower, less GC pressure) with no
  delivery regression.
- Duplicate suppression for literal subscriptions and GET/HGET polling:
  identical consecutive payloads are compared as raw bytes and skipped before
  the string decode and fan-out (`"SkipRepeatedMessages": true`, default on;
  PSUB patterns are never deduplicated because channels interleave). Feeds
  that republish identical payloads as a liveness signal should set it false.
- Added `test/LoadTests`: reusable load test harness for the subscription path
  (throughput, bytes/message, GC counts; internal or external load).

## v1.1.2

### Changed

- Update check simplified: `UpdateCheck` is a plain boolean in
  `RedisExcel.json` (`"UpdateCheck": true`, default on) and a single worksheet
  function, `RedisUDFUpdateAvailable()`, returns TRUE/FALSE. The intermediate
  `{ "enabled": ... }` shape and the `RedisUDFUpdateInfo` matrix from v1.1.1
  are gone. The check runs in the background at add-in load and refreshes at
  most every 6 hours when the function recalculates; it never blocks Excel.

## v1.1.1

### Added

- Non-blocking update check against GitHub releases, enabled by default and
  controlled by the `UpdateCheck.enabled` setting in `RedisExcel.json`.
  Exposed through the `RedisUDFUpdateAvailable` and `RedisUDFUpdateInfo`
  worksheet functions; it runs in a background task and never blocks Excel.

## v1.1.0

### Fixed

- Pub/Sub subscriptions could be silently lost after copying/reopening a
  workbook or when several workbooks used the same channel. Disconnecting one
  topic removed every other topic's handler on that channel, and state shared
  between multiple RTD server instances could unsubscribe each other.
- The UDF channel listener froze permanently after a connection blip
  (`RedisUDFChannelLatest`).
- `System.Timers.Timer` callback exceptions were swallowed silently; timers
  now log errors and are properly disposed.
- `HGETALL` now returns valid JSON: `{"field":"value",...}`.
- `ExcelUpdateStyle` `Timer`/`Realtime` are honored (previously both behaved as
  realtime).
- RTD connect/sync timeouts now respect `RedisExcel.json` (previously
  hardcoded).
- Config files missing the `RTD`/`UDF` sections no longer throw
  `NullReferenceException`; defaults are applied.
- `RedisUDFJSONToMatrix` object row-fill loop fixed; `[null]` elements are
  handled.

### Changed

- Connection/subscription handling refactored into `RedisConnectionManager`,
  `RedisSubscriptionManager` and `RedisRuntime` (single connection point,
  ref-counted channels; subscriptions survive reconnects).
- Polling uses `MGET` for `GET` topics and a pipeline for `HGET`/`HGETALL`.
- Added unit tests (`test/RedisExcel.Tests`) and an Excel end-to-end test
  (`test/Run-ExcelE2E.ps1`); smoke tests added under `test/SmokeTests`.
- All Excel function names, arguments and descriptions are unchanged.
