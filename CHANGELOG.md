# Changelog

## v1.1.0 (unreleased)

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
  ref-counted channels, automatic re-subscribe on reconnects).
- Polling uses `MGET` for `GET` topics and a pipeline for `HGET`/`HGETALL`.
- Added unit tests (`test/RedisExcel.Tests`) and an Excel end-to-end test
  (`test/Run-ExcelE2E.ps1`); smoke tests moved to `test/SmokeTests`.
- All Excel function names, arguments and descriptions are unchanged.
