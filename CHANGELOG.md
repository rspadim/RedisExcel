# Changelog

## v1.1.3 (unreleased)

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
