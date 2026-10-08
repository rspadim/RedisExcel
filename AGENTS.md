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
   Users' spreadsheets depend on every one of them.
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
| `RedisSubscriptionManager.cs` | Ref-counted Pub/Sub: one StackExchange.Redis handler per `(host, channel, pattern)` broadcasting to N listeners; auto-resubscribe on reconnect. |
| `RedisRuntime.cs` | Process-wide singleton wiring connections + subscriptions; shutdown on add-in unload. |
| `RedisRtd.cs` | RTD server lifecycle, topic registries, timers, polling; `RedisRtdStatus` functions. |
| `RedisUDF.cs` | `[ExcelFunction]` implementations; thin wrappers over the managers. |
| `ExcelJson.cs` | `RedisUDFMatrixToJSON` / `RedisUDFJSONToMatrix`. |
| `RedisResultFormatter.cs` | Value formatting sent to Excel (HGETALL as valid JSON). |

## Runtime model (why it is the way it is)

- Excel may create more than one RTD server per process. Connections and
  subscriptions are **process-wide** (`RedisRuntime`); topics are per RTD
  server instance.
- One StackExchange.Redis handler per channel broadcasts to every listener;
  a channel is unsubscribed only when its last listener leaves.
- StackExchange.Redis does **not** re-subscribe after reconnects;
  `ConnectionRestored` → `RedisSubscriptionManager.ResubscribeHost` does it.
- RTD push model: the poll timer (`RedisUpdateRateMs`) reads values; when
  real-time is off (`Automatic` over threshold, or `Timer` style), the Excel
  timer (`ExcelUpdateRateMs`) flushes dirty values.
- `HGETALL` output is valid JSON: `{"field":"value",...}`.
- Config file is searched in: user profile, Excel folder, `C:\Windows`
  (first found wins).

## Build

```powershell
dotnet build RedisExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`:

- `RedisExcel-packed.xll` (32-bit) and `RedisExcel64-packed.xll` (64-bit),
- `RedisExcel.dll`, plus the dependency assemblies are packed inside the XLLs.

CI (`.github/workflows/build.yml`) builds on `v*` tags and defines `GIT_TAG`,
which is embedded in the Redis `ClientName` for diagnostics.

## Tests

### 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

Covers: `ExcelJson` conversions, `AppConfig.Sanitize`/`ResolveHostCore`,
subscription keys, HGETALL formatting.

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
connections server-side and verifies automatic recovery.

Useful parameters:

- `-RedisHost <host:port>` — default `localhost:6379`.
- `-KeyPrefix <prefix>` — default `test.redisexcel` (keys and channels).
- `-RealChannel <name>` / `-RealPattern <pattern>` — optionally subscribe to
  real read-only channels to check live data.
- `-SkipClientKill` — never run `CLIENT KILL TYPE pubsub`
  (auto-skipped for non-local hosts).
- `-KeepExcelOpen` — leave Excel open for debugging.

For local hosts the workbook is saved to `test\RedisExcel.Test.xlsx` (committed
sample). For remote hosts it is saved to `%TEMP%` so the host never lands in
the repository. The committed workbook only references `localhost:6379` and
`test.redisexcel.*`.

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

Push a `v*` tag; the CI workflow builds and publishes the packed XLLs plus
`NLog.config` and `RedisExcel.json` as release assets. Next planned version:
`v1.1.0`.
