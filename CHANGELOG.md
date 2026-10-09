# Changelog

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
- Failed Redis connects are no longer cached: the failed attempt is replaced
  with a fresh entry atomically (compare-and-swap) so the next call retries
  instead of the host staying poisoned until Excel restarts; a concurrently
  created entry is never removed.
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
  empty/null values are stored as `""` (they no longer reach Redis as a null
  `RedisValue` that would delete the key), now also covering key-value/matrix
  setters, list pushes and channel publishes.
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
