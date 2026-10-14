---
description: Excel add-in (XLL) that connects Excel to Redis through RTD and worksheet functions
---

# RedisExcel

![RedisExcel — Excel and Redis, via UDF, RTD and Pub/Sub](../assets/redisexcel.jpg){: .rdx-hero }

**Excel add-in (XLL) that connects Excel to [Redis](https://redis.io/)** through
RTD (Real-Time Data) and UDFs (worksheet functions): type a formula, the Redis
value lands in the cell — and in the streaming functions it stays live. The
other direction works too: you can publish to Redis straight from a cell.

Supports polling (`GET`, `HGET`, `HGETALL`) and Pub/Sub (`SUB`, `PSUB`), keys,
hashes, lists, sets and counters, JSON ↔ matrix conversion and end-to-end
Sentinel/Cluster. Built with Excel-DNA on .NET Framework 4.8, using
StackExchange.Redis, NLog and Newtonsoft.Json.

Windows binaries are **digitally signed**; code signing is provided free of
charge to open-source projects by the
[SignPath Foundation](https://signpath.org/).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/RedisExcel/blob/main/LICENSE)
[![Latest release](https://img.shields.io/github/v/release/rspadim/RedisExcel?sort=semver)](https://github.com/rspadim/RedisExcel/releases/latest)


Requirements
------------

- **Windows** with Microsoft Office (Excel), 64-bit or 32-bit.
- **.NET Framework 4.8**, included in Windows 10 and Windows 11.
- A **Redis-compatible server** — Redis, Valkey or Memurai. The installer can
  set one up for you (Docker/WSL); Docker is optional and never required.
- No account, sign-up or cloud service: the add-in talks straight to the server
  you configure.


What it does
------------

| Area | Coverage |
| --- | --- |
| Live cells (RTD) | `GET`, `HGET`, `HGETALL` polling and real-time `SUB`/`PSUB`, with automatic retry and coalescing to keep the sheet smooth on fast feeds |
| Worksheet functions (UDF) | read and write keys, hashes, lists, sets, counters and TTL, `SCAN` patterns and JSON ↔ matrix conversion |
| Controlled writes | 24 `...NonVolatile` functions: they write once (when entered or when an argument changes) instead of re-running on every recalculation |
| Pub/Sub | `ChannelPublish` (text and JSON), `ChannelLatest`, `ChannelUnsubscribe` and `PubSubChannelsInfo` |
| High availability | Native Sentinel and Cluster via StackExchange.Redis: follows MOVED/ASK and re-subscribes after a failover |
| Diagnostics | RTD counters, `RedisUDFConnectionCount`, `RedisUDFUpdateAvailable` and `RedisCheckConfigFile` (validates the configuration inside the sheet) |


Download
--------

Download the latest version from the
**[releases page](https://github.com/rspadim/RedisExcel/releases/latest)**:

| File | What it is for |
| --- | --- |
| `RedisExcelSetup-<version>.exe` | One-click installer (per user; wizard in 10 languages) |
| `RedisExcel64-packed.xll` | Add-in for 64-bit Office |
| `RedisExcel-packed.xll` | Add-in for 32-bit Office |
| `RedisExcel.dll` | Managed assembly of the add-in (ships next to the XLL) |
| `NLog.config` | Log configuration |
| `RedisExcel.json` | Starter configuration (example) |

Releases are digitally signed; the signing service is provided by the
[SignPath Foundation](https://signpath.org/), which issues certificates for
open-source projects.

No Redis server yet? The installer can set one up for you (Redis or Valkey via
Docker/WSL, or Memurai) and tests the connection at the end.


Install
-------

1. Run `RedisExcelSetup-<version>.exe`. The wizard copies the add-in to
   `C:\RedisExcel`, keeps only the XLL that matches your Office bitness, writes
   a starter `RedisExcel.json` and can test the connection and open a sample
   workbook.
2. Open Excel. If the add-in does not load by itself, enable it under
   **File → Options → Add-ins → Manage: Excel Add-ins → Browse…** and pick
   `RedisExcel64-packed.xll` (`RedisExcel-packed.xll` for 32-bit Office).
3. Point `RedisExcel.json` at your server — the default is `localhost:6379`;
   you can use named aliases and full connection strings, including a password
   (which never appears in errors or in the log).

!!! tip "No Redis server yet?"
    The **“No Redis server yet? Run one on Windows”** guide in
    [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md)
    walks through it (Docker, Memurai or WSL2). They all listen on
    `localhost:6379` by default, which is what the starter configuration
    expects.

!!! warning "SmartScreen notice"
    Windows may warn about the XLL when it first loads: click
    **More info → Run anyway**. If a downloaded file stays blocked, right-click
    it → **Properties** → **Unblock** before loading the add-in.

The installer wizard is available in 10 languages; inside Excel the add-in is
always in English, so formula results do not change with the Windows language.


Examples
--------

```excel
=RedisUDFSet("btc_price", A1)                        ' writes a key with the value of cell A1
=RedisUDFGet("btc_price")                            ' reads the key (no character limit)
=RedisUDFHashGetAll("portfolio")                     ' the whole hash as a two-column matrix
=RedisUDFChannelPublish("alerts", "HIGH")            ' publishes to a channel
=RedisUDFChannelLatest("alerts")                     ' latest message on the channel, live
=RTD("RedisRtd", , "SUB", "alerts")                  ' live cell via RTD
=RedisUDFJSONToMatrix(RedisUDFChannelLatest("feed")) ' JSON payload → sheet matrix
```

And the other direction, publishing from outside with Python:

```python
import redis

r = redis.StrictRedis(host="localhost", port=6379, db=0)
r.set("btc_price", "67000.50")   # Excel reads it with  =RedisUDFGet("btc_price")
r.publish("alerts", "HIGH")      # shows up in  =RedisUDFChannelLatest("alerts")
```


Documentation
-------------

- **Full guide** (install, RTD, UDF, configuration, tests):
  [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md)
- **Changelog:**
  [CHANGELOG.md](https://github.com/rspadim/RedisExcel/blob/main/CHANGELOG.md)
- **Configuration:** example
  [`RedisExcel.json`](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.json),
  [schema](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.schema.json)
  and [NLog.config](https://github.com/rspadim/RedisExcel/blob/main/NLog.config)
- **Installer:**
  [installer/README.md](https://github.com/rspadim/RedisExcel/blob/main/installer/README.md)
- **Code signing:**
  [SIGNING.md](https://github.com/rspadim/RedisExcel/blob/main/SIGNING.md)


Privacy
-------

RedisExcel collects no user data and sends nothing to us. The only request it
makes on its own is a background check against the GitHub releases API to see
whether a newer version exists (turn it off with `"UpdateCheck": false` in
`RedisExcel.json`); failures are ignored. It writes a local log
(`RedisExcel.log`) next to the add-in files. Redis traffic goes only to the
servers you configure, and credentials in connection strings are masked in
errors and in the log.


Support and contributing
------------------------

Questions, bug reports and pull requests are welcome at
**[github.com/rspadim/RedisExcel](https://github.com/rspadim/RedisExcel/issues)**.
When opening a bug, include the installed version, your Office architecture and
the error message (from the cell and/or `RedisExcel.log`).

**Translations are welcome.** Want to keep this page in your language? Open a
pull request touching `docs/` (the site content) — we have Portuguese, English
and Spanish today.


License
-------

RedisExcel is distributed under the **MIT License** — see
[LICENSE](https://github.com/rspadim/RedisExcel/blob/main/LICENSE).
Bundled third-party components: Excel-DNA (Zlib), StackExchange.Redis (MIT),
Newtonsoft.Json (MIT), NLog (BSD-3-Clause) and NJsonSchema/Namotion.Reflection
(MIT).

> RedisExcel is an independent project; it is not affiliated with or endorsed by
> Redis Ltd.
