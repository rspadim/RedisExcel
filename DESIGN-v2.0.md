# RedisExcel 2.0.0 — Server Clusters, Read/Write Policies, Client-Side Failover

Status: design draft for discussion. Local file, not committed. Target release: v2.0.0.

## 1. Goals

- Named clusters usable anywhere a host is accepted (`Default.host`, the
  per-section `RTD`/`UDF` overrides, and the optional host argument of every
  UDF and RTD topic).
- Centralized server choice: one router decides the endpoint(s) for every
  Redis call — UDF commands, RTD polling, RTD subscriptions and UDF
  long-lived subscriptions.
- Cost-aware selection: prefer the cheapest healthy member. While parked on an
  expensive fallback, keep evaluating the cheap member and move back
  automatically when it recovers (failback), including Pub/Sub.
- Explicit read policies: `single` (one member: cheapest / weighted /
  round-robin / latency) and `multi` (several members at the same time).
- Explicit write policies: `single` (the member replicates the write; master
  or Sentinel) and `multi` (members are independent and do not replicate).
- Redis Sentinel supported transparently as a member `target`.
- No secrets in logs: connection strings are always masked.
- Strict v2 configuration: invalid files fail closed with a visible `#CONFIG`
  error (no silent fallback), fixable at runtime with `RedisUDFConfigReload`.

## 2. Non-goals (2.0.0)

- Conflict resolution / reconciliation across independent members.
- Sharded Pub/Sub (`SPUBLISH`) semantics.
- A cross-process command channel; convergence uses the optional file watcher
  (section 4.8.1).
- A separate watchdog service; everything runs inside the add-in process.
- Redis Cluster client-side routing (StackExchange.Redis already handles it
  when a `target` points at a cluster).

## 3. Concepts

| Term | Meaning |
| --- | --- |
| Server | Physical target plus metadata: `cost`, `writeAccepted`, `weight`. |
| Cluster | Logical name: ordered members plus read/write policies. |
| Member | A server referenced by a cluster. |
| Router | Process-wide component resolving names to endpoints; owns health and migrations. |
| Logical subscription | `(cluster, channel/pattern)` as seen by the user. |
| Physical subscription | `(endpoint, channel/pattern)` as tracked by `RedisSubscriptionManager`. |

Resolution order for any host argument: `Servers` name, then `Clusters` name,
then `UnknownHosts` name, then the raw connection string. The validator warns
when a name exists in more than one section.

## 4. Configuration

Layout principle: connection defaults live in one place; `RTD` and `UDF` keep
only their own settings and optional overrides. `Servers` holds physical
definitions, `Clusters` holds logical groups and policies. The client decides
the topology through the read/write modes; the add-in does not infer it.

### 4.0 Simplified layout (v2, strict)

v1 duplicated `host`/`timeout` in `RTD` and `UDF`. v2 introduces a single
`Default` block; the sections keep optional overrides:

```json
{
  "ConfigVersion": 2,

  "Default": { "host": "prod", "timeout": 1000 },

  "RTD": {
    "RedisUpdateRateMs": 1000,
    "ExcelUpdateRateMs": 100,
    "MessageCounterThreshold": 1000,
    "ExcelUpdateStyle": "Automatic",
    "UseGetMultiple": true,
    "CoalesceRealtimeUpdates": true,
    "host": "feed",                 // optional override
    "timeout": 1000                 // optional override
  },

  "UDF": {
    "host": "internal",             // optional override; can be omitted
    "timeout": 1000                 // optional override; can be omitted
  },

  "Servers":    { "...": "..." },
  "Clusters":   { "...": "..." },
  "UnknownHosts": { "...": "..." },

  "UpdateCheck": true,
  "SkipRepeatedMessages": true,
  "LearnUnknownHosts": true,
  "WatchConfig": false
}
```

- Minimal single-server file: `{ "Default": { "host": "localhost:6379" } }`;
  every other block defaults.
- `RTD` and `UDF` may be omitted entirely. `CoalesceRealtimeUpdates` lives
  inside `RTD` (it is RTD-only).
- A cluster with only `members` uses the defaults: `read` =
  single/cheapest with standard failback, `write` = single/cheapest-accepted.
- `members` accepts a name from `Servers` or an inline server object:
  ```json
  "members": [
    { "target": "cheap.example:6379", "cost": 0, "writeAccepted": false },
    "aws-sentinel"
  ]
  ```
  Inline members get an internal stable name (`<cluster>#<index>`) used in
  logs, health keys and `write.pinned`. Named entries are recommended when a
  member is shared by clusters or referenced directly in formulas.
- v2 only: **no v1 compatibility** in 2.0 (major version). The v1 keys
  `RTD.host`, `UDF.host`, `RTD.timeout`, `UDF.timeout`, the root-level
  `CoalesceRealtimeUpdates` and the `ExcelUpdateRateMS` spelling are not
  accepted; a config containing them (or otherwise not matching v2) is
  rejected as a whole.
- Rejection is fail-closed and visible: no Redis connection is opened and every
  UDF/RTD function returns a `#CONFIG` error identifying the file and the
  first offending item (for example
  `#CONFIG: RedisExcel.config.v2.json:12 unknown property 'RTD.host' (use Default.host);
  run RedisUDFConfigValidate()`). Diagnostic functions (`RedisUDFConfigValidate`,
  `RedisUDFConfigSchema`, `RedisUDFConfigPath`, `RedisRTDStatus*`) remain
  available. Warnings never block; only errors do. This replaces the old silent
  fallback to defaults and the arbitrary "run against localhost" state. After
  fixing the file, `RedisUDFConfigReload()` applies it without restarting
  Excel.
- Single standard location: `C:\RedisExcel\` holds `RedisExcel.config.v2.json`
  (configuration) and `RedisExcel.pointer.json` (section 4.8). No
  per-user/Excel-folder/`C:\Windows` lookup. The directory is created by the
  installer/admin (README documents the ACL recommendation) and is
  machine-wide: a permissive ACL lets any user change the config and read
  stored credentials; a read-only folder disables learning and
  `RedisUDFConfigSet` while everything else keeps working.
- New file name: `RedisExcel.config.v2.json`; 2.0 never reads a v1-named file
  (`RedisExcel.json`). No collision, and v1 and v2 add-ins can coexist with
  their own configs; migration is a documented copy to the standard
  directory, not a runtime compatibility path.
- Migration is a one-time file edit, documented in the release notes/README:
  copy the v1 file to `RedisExcel.config.v2.json` and apply
  `RTD.host`/`UDF.host` -> `Default.host` (or per-section overrides), root
  `CoalesceRealtimeUpdates` -> `RTD.CoalesceRealtimeUpdates`,
  `ExcelUpdateRateMS` -> `ExcelUpdateRateMs`. The committed
  `RedisExcel.config.v2.json` sample and the README configuration section are
  rewritten to the v2 shape as part of P1.
- `ConfigVersion` is optional and only guards the future: a version greater
  than the supported one is rejected with a clear error (same fail-closed
  behavior above). It is not used to detect v1 — that detection is by the
  legacy keys above.
- Unknown-host learning always writes the v2 shape.

**Default host resolution.** A call without a host argument resolves in this
order: the section override (`RTD.host` for RTD topics, `UDF.host` for UDF
calls) -> `Default.host` -> the built-in constant
(`localhost:6379,password=,defaultDatabase=0,ssl=False,abortConnect=False`),
which also applies when no configuration file exists. Each of these values is
a normal host reference (a `Servers` name, a `Clusters` name, an
`UnknownHosts` name or a raw connection string, resolved as in section 3).
When it names a cluster, the cluster policies decide the physical member per
call — `read`/`write` policies, cheapest + failback — so the default is a
logical name, not a fixed endpoint. Configuration values are never treated as
discoveries (section 4.7). A value without connection-string syntax that does
not match a known name is a dangling-reference validation error, not a silent
fallback; there is no implicit "first cluster" default.

### 4.1 Servers

A value may be a plain connection string shorthand or an object:

```json
"Servers": {
  "cheap-replica": { "target": "redis-cheap.example:6379", "cost": 0, "writeAccepted": false },
  "aws-sentinel":  { "target": "sentinel-a.example:26379,sentinel-b.example:26379,serviceName=mymaster", "cost": 9, "writeAccepted": true },
  "mirror-a":      { "target": "redis-a.example:6379", "cost": 1, "writeAccepted": true },
  "mirror-b":      { "target": "redis-b.example:6379", "cost": 1, "writeAccepted": true }
}
```

- `target`: connection string; may be a direct endpoint, a replica, or a
  Sentinel connection string (`serviceName=`).
- `cost`: relative value, lower wins. `0` = cheapest. No currency semantics.
  Accepts an integer (sets ingress = egress) or an object
  `{ "ingress": n, "egress": n }`. Read/sub selection uses `egress` (server ->
  client traffic dominates); write selection uses `ingress`.
- `writeAccepted`: `false` for replicas/read-only nodes. Writes never go there.
- `weight`: used by weighted selection among equal-cost members.

### 4.2 Clusters

```json
"Clusters": {
  "cheap-first": {
    "members": ["cheap-replica", "aws-sentinel"],
    "read": {
      "mode": "single",
      "selection": "cheapest",
      "failback": { "probeMs": 5000, "healthyAfter": 3, "cooldownMs": 30000, "jitterPercent": 20 }
    },
    "write": { "mode": "single", "selection": "cheapest-accepted" }
  },
  "independent-mirrors": {
    "members": ["mirror-a", "mirror-b"],
    "read": { "mode": "multi", "merge": "race" },
    "write": { "mode": "multi", "onPartialFailure": "log", "allowUnsafeMultiWrites": false }
  }
}
```

### 4.3 Read policies

`read.mode = "single"` — one member at a time:

- `selection: "cheapest"` (default): order by `cost`, tie-break by `weight`
  then member order. The active member stays until a strictly cheaper member
  is healthy and confirmed (`failback` hysteresis).
- `selection: "weighted"`: weighted distribution among the cheapest healthy
  cost tier; the next tier is used only when the whole tier is unhealthy.
- `selection: "roundrobin"`: rotate all healthy members. Uses expensive
  members too; intended for capacity, not cost.
- `selection: "latency"`: lowest measured PING RTT (phase 3).
- `failback`: `probeMs` probe interval; `probeTimeoutMs` PING timeout;
  `healthyAfter` consecutive successes required to promote; `cooldownMs`
  minimum wait after a demotion before promoting again; `jitterPercent`
  per-process random spread. Defaults in section 6.4.

`read.mode = "multi"` — several members at the same time:

- Applies to every read command (all UDF reads and RTD
  `GET`/`HGET`/`HGETALL`). In `multi`, `SUB`/`PSUB` use fan-in (below); in
  `single`, subscriptions follow the normal single-member selection.
- Polling commands: every healthy member is queried; the `merge` option
  decides the answer:
  - `race` (default): first successful response wins — mirrored/redundant
    data, lowest latency, tolerates one member being slow.
  - `union` (option): first non-null response in member order, per key — data
    that is split across independent members.
- Subscriptions (`SUB`, `PSUB`): fan-in — subscribe on every healthy member;
  messages from every source are delivered (see dedupe).
- Note: `multi` multiplies traffic. It is for redundancy or partitioned data,
  not for cost. The cost scenario is `single` + `cheapest`.

### 4.4 Write policies

`write.mode = "single"` — one member, replication assumed:

- Candidates: members with `writeAccepted: true` that are healthy.
- `selection: "cheapest-accepted"` (default) or `pinned: "<member>"`.
- Replicated topology: the only `writeAccepted` member is the master — direct
  or the Sentinel member whose master is discovered automatically.
- If no candidate is available the call fails with an explicit error instead
  of a raw READONLY from a replica.

`write.mode = "multi"` — N members, no replication assumed:

- Every healthy `writeAccepted` member, executed in parallel, best-effort.
- `onPartialFailure: "log" | "error"` (default `log`): with `error` any member
  failure fails the call; with `log` the call succeeds when at least one
  member accepted the write and the failures are logged.
- Return value: aggregated across members by the function's natural
  reduction; diverging values are logged at Warn:
  - status (`SET`, `SETJSON`, `SETKV`, `SETEX`, `EXPIRE`, `RENAME`,
    `ChannelPublish`): success only if every required member succeeded;
  - operation counts (`DEL`, `HDEL`, `SADD`, `SREM`): sum of the values
    returned by the members that succeeded;
  - state quantities (`INCR`, `INCRBY`, list push length, list pop value):
    value of the first successful member (guarded by the non-idempotent rule;
    divergence means the members were not true mirrors).
- Non-idempotent commands are rejected unless `allowUnsafeMultiWrites: true`,
  because a partial failure cannot be repaired by re-sending:
  - State-changing on re-run: `INCR`, `INCRBY`, `APPEND`, `LPUSH`, `RPUSH`,
    `LPOP`, `RPOP`.
  - Safe to re-send: `SET`, `SETEX`, `EXPIRE`, `DEL`, `HSET`, `HDEL`, `RENAME`,
    `SADD`/`SREM` (final state is the same).
- `ChannelPublish` is classified as a write and must reach the master in
  replicated topologies (publishing to a replica does not reach subscribers on
  the master).

### 4.5 Full example

```json
{
  "$schema": "https://raw.githubusercontent.com/rspadim/RedisExcel/main/RedisExcel.schema.v2.json",
  "ConfigVersion": 2,

  "Default": { "host": "cheap-first", "timeout": 1000 },

  "RTD": {
    "RedisUpdateRateMs": 1000,
    "ExcelUpdateRateMs": 100,
    "ExcelUpdateStyle": "Automatic",
    "UseGetMultiple": true,
    "CoalesceRealtimeUpdates": true
  },

  "Servers": {
    "cheap-replica": { "target": "redis-cheap.example:6379", "cost": 0, "writeAccepted": false },
    "aws-sentinel":  { "target": "sentinel-a.example:26379,sentinel-b.example:26379,serviceName=mymaster", "cost": 9, "writeAccepted": true }
  },

  "Clusters": {
    "cheap-first": {
      "members": ["cheap-replica", "aws-sentinel"],
      "read": {
        "mode": "single",
        "selection": "cheapest",
        "failback": { "probeMs": 5000, "healthyAfter": 3, "cooldownMs": 30000, "jitterPercent": 20 }
      },
      "write": { "mode": "single", "selection": "cheapest-accepted" }
    }
  },

  "UpdateCheck": true,
  "SkipRepeatedMessages": true,
  "LearnUnknownHosts": true
}
```

`Servers` values accept both forms (string shorthand or object) through a
Newtonsoft converter. Property names are case-sensitive and unknown properties
are validation errors (the fail-closed behavior of section 4.0); `$schema` is
explicitly whitelisted for editors. Comments (`//`, `/* */`) are accepted by
Newtonsoft.

### 4.6 Tooling: embedded schema + static validator

- `RedisExcel.schema.v2.json` is maintained in the repository (for review,
  diffs and editor IntelliSense) and **embedded into `RedisExcel.dll` at
  build time** (`<EmbeddedResource>`). The packed XLL carries the DLL
  byte-for-byte, so its manifest resources are preserved: no extra file is
  installed next to the add-in.
- `ConfigSchema.Text` is the static accessor: the resource is read once into a
  `static readonly string`, cached for the process. No disk reads, no runtime
  schema generation, no dynamic dependencies.
- Validation is static C# (`ConfigValidation`), whose single implementation is
  shared by:
  - the worksheet function `RedisUDFConfigValidate()` -> returns a matrix of
    `(severity, path, line, message)` or `"OK"`; pure, never writes files;
  - an optional console wrapper (`tools/ConfigValidator`) for CI, referencing
    the same code (no duplicated rules).
- Rules: dangling member/host references (`members`, `Default.host`,
  `RTD.host`, `UDF.host`), empty clusters, duplicate targets, name
  collision across `Servers`/`Clusters`/`UnknownHosts` (warning), negative
  `cost`, invalid enums, `write.multi` combined with non-idempotent usage
  (warning), `UnknownHosts` entries duplicating a `Servers` target, unknown
  properties (typo detection matching `additionalProperties: false`; the
  `$schema` key itself is allowed). Property matching is case-sensitive, so
  the canonical spellings (`host`, `timeout`, `ExcelUpdateRateMs`) are the
  only ones accepted. Messages never print unmasked secrets.
- v1 keys are errors, not suggestions: `RTD.host`/`UDF.host`/`RTD.timeout`/
  `UDF.timeout`, root `CoalesceRealtimeUpdates` and the `ExcelUpdateRateMS`
  spelling are reported with the v2 replacement, so the one-time migration can
  be done with the file open in the editor.
- Error severity blocks, warning severity does not: in the rejected state no
  Redis access happens and UDF/RTD return `#CONFIG`; warnings (for example
  non-idempotent multi-write) only log.
- `RedisUDFConfigValidate()` validates the same target a no-argument
  `RedisUDFConfigReload()` would use (the active source, else the standard
  path), even while the active configuration is rejected, so a fix can be
  confirmed and applied without restarting Excel.
- Full JSON Schema validation (strict keyword checking) runs only in the test
  project (NJsonSchema as a test-only dependency), against sample configs and
  against a schema generated from the config model, so the shipped binary
  stays dependency-free. The static checker covers the same ground for users.
- Editor use: `$schema` can point to the raw GitHub URL, or the user can save
  the output of `RedisUDFConfigSchema()` (same embedded text); comments
  (`//`, `/* */`) are parsed natively by Newtonsoft.

### 4.7 Unknown host learning (default on, opt-out)

Root option `"LearnUnknownHosts": true` (set `false` to opt out). A literal
connection string passed in a formula that is not a known name and is not
already mapped is recorded under `UnknownHosts`:

```json
"UnknownHosts": {
  "unknown_localhost_6379": "localhost:6379,defaultDatabase=1"
}
```

- Trigger: only formula arguments that parse as a connection string
  (`ConfigurationOptions.Parse` succeeds, at least one endpoint). Bare names
  are never learned; `Default.host` and section overrides are configuration,
  not discoveries.
- Naming: `unknown_<host>_<port>` (sanitized; default port 6379 when absent).
  On collision with a different target a numeric suffix is added (`_2`). The
  generated name is stable for the same literal.
- Dedupe: the literal (trimmed) is compared against every `Servers` target,
  every `UnknownHosts` value and the in-memory learned map; an identical
  target is never written twice. Two literals that only differ in option order
  or whitespace are different keys (exact-match limitation, documented).
- The literal -> generated name mapping is registered in memory at call time,
  so the current session resolves it immediately; the file is not touched per
  call.
- Persistence: one background writer, batched (timer/task), never on the Excel
  thread.
  - Writes into the active file (standard location, session-adopted by
    `RedisUDFConfigReload`, or forced by a pointer); if no file is active,
    creates `C:\RedisExcel\RedisExcel.config.v2.json` (creating the directory
    when the permissions allow).
  - If the loaded file is not writable (for example `C:\Windows`), learning is
    disabled for the session after a single warning.
  - Round trip via `JObject` with `CommentHandling.Load`: existing comments,
    unknown properties and key order are preserved; formatting may be
    normalized to indented JSON.
  - Atomic replace (temp file in the same directory + `File.Replace`), UTF-8
    without BOM.
  - Immediately before writing, the file is re-read and only the new entries
    are merged, so two Excel processes do not lose each other's discoveries
    (last writer wins per entry).
- Security: on by default; writing the literal through to the config file is
  accepted by design. A literal containing `password=`/`user=` is stored
  verbatim (same trust level as the rest of the config file); every report and
  log masks it. Each newly learned entry is logged once at Info.
- Other processes and future sessions pick entries up on their next start;
  within the current process, `RedisUDFConfigReload()` (section 4.8) re-reads
  the file, and with `WatchConfig` enabled every watching process converges
  automatically.
- The learned section is machine-managed. To promote an entry: copy it into
  `Servers` with a real name, add `cost`/`writeAccepted`, and delete it from
  `UnknownHosts`; the validator warns while both exist pointing at the same
  target. Cluster `members` reference `Servers` names only.

### 4.8 Configuration source and reload

Three functions manage the active configuration at runtime.

`RedisUDFConfigPath()` (no arguments) reports the file in use and how it was
chosen: `C:\RedisExcel\RedisExcel.config.v2.json (standard, loaded 2026-10-09
14:22, hash 1f3a..)`, `D:\feed.json (explicit, loaded ...)`,
`D:\feed.json (forced, loaded ...)`, `(defaults; no config file found)` or
`(invalid: C:\RedisExcel\RedisExcel.config.v2.json; 2 error(s); run
RedisUDFConfigValidate())`. It is diagnostic and always works, even in the
rejected state.

`RedisUDFConfigReload(optionalPath)` re-reads and swaps the active
configuration without restarting Excel:

- no argument: re-reads the active source; if none is active (defaults or
  rejection), uses the standard `C:\RedisExcel\RedisExcel.config.v2.json`;
- non-empty path: validates and adopts that file as the active source for the
  process lifetime (it does not need to be in the standard directory);
- empty string `""`: drops an explicit source and returns to the standard
  path.

`RedisUDFConfigSet(optionalPath)` is the persistent variant:

- non-empty path: validates the file, adopts it now and writes the pointer
  `C:\RedisExcel\RedisExcel.pointer.json`
  (`{ "ConfigPath": "D:\\feed.json" }`, absolute path, atomic write). A write
  without permission (or a missing unwritable directory) returns an error and
  persists nothing.
- empty string `""`: clears the pointer and re-resolves to the standard
  `C:\RedisExcel\RedisExcel.config.v2.json`, now and at startup.
- A broken file is never persisted: validation errors are returned and neither
  the pointer nor the active snapshot changes.
- Startup resolution order: the pointer, then the standard
  `C:\RedisExcel\RedisExcel.config.v2.json`. A pointered file that is missing or
  invalid is a fail-closed `#CONFIG` state naming the pointer and the target —
  never a silent fallback to another file.
- Relative paths are rejected.

`RedisUDFConfigReload(path)` adoption is session-only; `RedisUDFConfigSet` is
what survives restarts.

Reload semantics:

- Transactional: the new file is parsed and validated first. On any error the
  current snapshot stays active (a working session is never killed by a typo)
  and the function returns the `(severity, path, line, message)` matrix, same
  shape as `RedisUDFConfigValidate()`. On success it returns `"OK"` with the
  loaded path and clears a previous `#CONFIG` state, so the startup-error flow
  is: fix the file, reload, keep working. Warnings are returned but never
  prevent the swap.
- If no valid configuration is active (startup rejection), an invalid reload
  keeps the rejected state and returns the errors.
- No-op guard: when the file content hash equals the active snapshot, the swap
  and the reconciliation pass are skipped (accidental double calls are
  harmless).
- What applies when:
  - `Servers`/`Clusters`/`UnknownHosts`/`Default`: immediately for new
    resolutions; a reconciliation pass re-evaluates existing placements and
    the failback machinery migrates RTD and UDF subscriptions accordingly.
  - RTD `RedisUpdateRateMs`/`ExcelUpdateRateMs`: timer intervals are updated;
    `ExcelUpdateStyle`, `MessageCounterThreshold`, `UseGetMultiple`,
    `CoalesceRealtimeUpdates` and `SkipRepeatedMessages`: volatile snapshot
    fields applied on the next tick/message.
  - `timeout`: applies to connections created after the reload; existing
    multiplexers keep the values they were created with (documented).
  - `UpdateCheck` and `LearnUnknownHosts`: next use.
  - `WatchConfig`/`WatchConfigMs`: the watcher is (re)configured immediately
    (enabled starts it, disabled stops it).
- Removed names: an RTD topic whose logical name no longer resolves publishes
  `#CONFIG` and releases its subscription; UDF long-lived subscriptions are
  released and logged. If the name still resolves but the member set changed,
  normal selection policy applies.
- Pending learned hosts are flushed before loading, so a reload never loses a
  discovery; after adopting an explicit path, the learner writes to that file.
- Concurrency: reload and the unknown-host writer are serialized by one lock;
  the configuration snapshot is swapped by reference (immutable) and readers
  never block. Concurrent reload calls run serialized, each doing a fresh
  load.
- The console wrapper can call the same code path.

### 4.8.1 Reload scope (current session vs all Excel instances)

- Each Excel process has its own in-memory configuration snapshot, connections
  and subscriptions. `RedisUDFConfigReload` and `RedisUDFConfigSet` take effect
  immediately in the process that evaluates the formula (the current session).
- Other open Excel processes converge by restarting (the pointer is read at
  startup), by calling reload themselves, or through the optional watcher.
- Optional watcher: root option `"WatchConfig": false` (default off), with an
  optional `WatchConfigMs` (default 5000). When on, a slow periodic check
  re-reads the pointer and hashes the active file inside the health worker;
  any change is applied through the same transactional reload path. Invalid
  intermediate states (a file being saved) are ignored and the previous
  snapshot stays. No IPC channel is used, so it works across processes and
  users without extra permissions.
- This is the supported way to get "all Excel instances" behavior; a direct
  cross-process command is out of scope for 2.0.

## 5. Centralized routing

One component, working name `ClusterRouter`, a single instance in
`RedisRuntime`:

```csharp
enum RedisUse { ReadCommand, WriteCommand, Subscription }

sealed class Route
{
    IReadOnlyList<string> Endpoints;   // 1 for single, N for multi
    bool IsMulti;
}

// UDF commands: resolve per call.
Route Resolve(string hostOrAlias, RedisUse use);

// Long-lived subscriptions: placement plus automatic migration.
IDisposable Subscribe(string hostOrAlias, string channel, bool pattern,
                      Action<string> onMessage, string origin);
```

- `AppConfig.ResolveRtdHost` / `ResolveUdfHost` stay as thin facades during the
  transition; every call site moves to the router.
- `RedisSubscriptionManager` remains physical (per endpoint/channel) and keeps
  its current contract. A handover is composed by the router: two physical
  registrations during the overlap, the old one disposed at the end.
- RTD topics store the logical name (cluster or server); the physical endpoint
  is resolved per polling tick and per migration decision.

### 5.1 Access classification (to verify against `RedisUDF.cs`)

- Read: `Get`, `GetMultiple`, `Type`, `Exists*`, `TTL*`, `HashGetField*`,
  `ListRange`, `SetMembers`, `ServerTime`, `Keys`, `ChannelLatest`.
- Subscription: `ChannelSubscribe`, `ChannelPatternSubscribe`; RTD `SUB`,
  `PSUB`.
- Write: `Set`, `SetJSON`, `SetKV`, `Rename`, `SetEx`, `Expire`, `Del`,
  `Incr`, `IncrBy`, `HashDel`, `ListPush*`, `ListPop*`, `SetAdd`,
  `SetRemove`, `ChannelPublish`; RTD GET/HGET/HGETALL are reads.

### 5.2 Execution wrapper

- Command call sites go through the router instead of resolving an endpoint
  and calling it directly:
  `T Router.Execute<T>(string hostOrAlias, RedisUse use, Func<IDatabase, T> command)`.
- Single mode: run on the preferred member; on a connectivity failure report it
  to the health monitor and, for reads, retry the next healthy member
  immediately (no waiting for the probe cycle).
- Retry rules:
  - reads: retry another member on connectivity errors and timeouts;
  - writes: retry only when the failure proves the command was not sent
    (connection down before send). A timeout is ambiguous — the write may have
    applied — so it is reported as an error instead of being silently
    duplicated on another member;
  - `BacklogPolicy` tuning (fail-fast variants) is considered so a down member
    fails fast instead of queueing until the timeout.
- Multi mode: run on every healthy member and aggregate per section 4.4.
- The wrapper is also the source of command-level health signals (per member
  and pool), complementing the multiplexer events.

## 6. Health monitoring

### 6.1 Worker

- One process-wide background worker, not one per member: the guarded timer
  pattern used elsewhere (`System.Timers.Timer` + `TickGate` + try/catch).
  A scheduler tick (default 500 ms) scans which probes are due, so a single
  worker serves every member. It runs on thread-pool threads, never on the
  Excel thread — a dedicated `Thread` is unnecessary and would fight the
  add-in unload lifecycle.
- Created lazily by `RedisRuntime` when the first cluster is resolved — or
  immediately when `WatchConfig` is enabled, so a rejected/default state can
  still recover without a route — and disposed on `Shutdown`.
- Startup delay: the first scan of each process is delayed by a random
  0..probeMs, avoiding a fleet-wide probe burst when Excel files open.
- Each member+pool has its own `nextProbeAt`, spread by `jitterPercent`, so
  probes from different members and from different Excel processes do not
  align.

### 6.2 Scope and probed pools

- Only members of clusters actually referenced are enrolled (lazy enrollment on
  first resolution).
- One state per `(member, pool)`: `RtdData`, `RtdSub`, `UdfData`. A use checks
  the pool it needs:
  - RTD GET/HGET/HGETALL -> `RtdData`;
  - RTD SUB/PSUB -> `RtdSub`;
  - UDF commands and subscriptions -> `UdfData`.
- Probe = `PING` with `probeTimeoutMs` (default 1000 ms) on the pool
  connection obtained from `RedisConnectionManager`. Connections are created
  on demand and cached; `AbortOnConnectFail=false` keeps the multiplexer alive
  and retrying in the background, so a recovered server is picked up without
  recreating anything.

### 6.3 State machine (per member+pool)

- States: `Up` / `Down`. A pool starts `Up` when its first real use connected
  successfully (so first use is never delayed); `healthyAfter` only gates
  re-promotion.
- Demotion:
  - passive `ConnectionFailed` from the multiplexer: immediate `Down`;
  - default two consecutive probe failures: `Down`;
  - command connectivity failures reported by the execution wrapper feed the
    same counter; a success resets it.
  - On demotion: record `demotedAt`, log at Info (masked target + reason), and
    signal the router so placements on this member can move.
- Promotion:
  - every probe success increments `consecutiveSuccesses`;
  - `Down` + `consecutiveSuccesses >= healthyAfter` +
    `UtcNow - demotedAt >= cooldownMs` -> `Up`;
  - on promotion: log at Info and raise `MemberPromoted`; the router schedules
    failback migrations (section 7.4). Promotion is compare-and-swap protected,
    so a member is promoted once and migrations are not duplicated.
- Flapping: consecutive successes plus cooldown damp it; a demotion within
  `cooldownMs` of a promotion increments a flap counter and logs at Warn.

### 6.4 Parameters

Per-cluster `failback` block; defaults:

| Parameter | Default | Meaning |
| --- | --- | --- |
| `probeMs` | 5000 | interval between probes per member+pool |
| `probeTimeoutMs` | 1000 | PING timeout |
| `healthyAfter` | 3 | consecutive successes required to promote |
| `cooldownMs` | 30000 | minimum time Down before promotion |
| `jitterPercent` | 20 | random spread per member/process |
| probe demotion threshold | 2 | consecutive probe failures to demote |

### 6.5 Thread safety and lifecycle

- State in concurrent dictionaries; counters and timestamps with `Interlocked`;
  promotion with compare-and-swap.
- No lock is held during socket I/O or logging.
- The tick callback catches and logs everything; a tick that finds the
  previous one still running is skipped, not queued.
- Health transitions never dispose connections; reconnect handling stays with
  StackExchange.Redis. The existing "failed initial connect is not cached"
  behavior is kept.
- No cross-process coordination: each Excel process decides for itself; probe
  jitter is the de-synchronization mechanism.

## 7. Failback and migration

### 7.1 Polling topics (GET / HGET / HGETALL)

- No handover: the next tick resolves the current preferred endpoint.
  Demotion is immediate; promotion applies at the next tick.
- `GETMULTI` grouping is by resolved endpoint, per tick.

### 7.2 Subscriptions (SUB / PSUB) — armed handover

1. A cheaper member becomes healthy (promotion) while the current source is
   expensive.
2. Subscribe on the cheap member; keep delivering from the current source
   ("armed").
3. On the first message from the new source, switch the active source and
   dispose the old registration.
4. If the old source becomes unhealthy while armed, switch immediately
   (failover without waiting for a message).

- Quiet channels: no message, no switch — and no traffic cost either. The
  handover stays armed at the cost of one idle subscription.
- Replica lag caveat: the first message from a lagging replica may repeat an
  already-delivered value. Logical dedupe drops exact repeats; coalescing
  masks most of the rest. Documented risk.
- Failure during handover: keep the old source (still subscribed) and retry
  later.

### 7.3 UDF long-lived subscriptions

- Same mechanism through `ClusterRouter.Subscribe`; this centralizes
  server choice for everything (per the agreed scope).
- `RedisUDFChannelUnsubscribe` disposes the router registration; the public
  Excel behavior is unchanged.

### 7.4 Scheduling

- One migration queue per process, serialized, one handover at a time, with
  backoff on failure.
- Triggers: probe promotion, demotion (move away), startup placement, and
  config reload reconciliation (section 4.8).

## 8. Connection lifecycle and cache key

- Cache key stays the physical endpoint + pool; `RedisConnectionManager` does
  not change.
- Both members remain connected after a hop (no eager disposal), so hopping
  back is cheap. Cost: a few idle sockets per cluster member.
- Status functions keep reporting physical connections; cluster-level
  diagnostics are added later.

## 9. Dedup and fan-in

- Current dedup is per physical channel (`ChannelState._lastMessage`). In
  fan-in mode the same value arrives once per source.
- Cluster `dedupe` option: `per-source` (current behavior, default) |
  `logical` (drop a payload identical to the last delivered on the logical
  channel regardless of source) | `none`.
- Patterns keep the current rule (never deduplicated).
- The armed-handover gate prevents old-source delivery after the switch.

## 10. Sentinel

- A member `target` may be a Sentinel connection string (`serviceName=...`).
  StackExchange.Redis discovers the master and follows failovers; the router
  needs nothing extra.
- Writes reach the discovered master; the member is `writeAccepted: true` when
  it is the write target.
- Direct replica members (`writeAccepted: false`) are the cost lever for
  read/sub traffic — subscribers on a replica receive messages published on
  the master (classic replication propagates `PUBLISH` while the link is up).
- Future (phase 3): discover replicas via `SENTINEL REPLICAS` and use them
  without hardcoding endpoints.

## 11. Security and logging

- `AppConfig.MaskConnectionString` redacts `password=`, `user=`,
  `sentinelPassword=`.
- Every target logged is the masked form plus logical identity (`cluster`,
  `member`).
- Health transitions and migrations log at Info: from/to member, reason
  (probe success, failure, handover), result.

## 12. Observability (optional, phase 3)

- New UDFs (`RedisUDFClusterStatus`) and RTD status helpers: current member per
  cluster, degraded flag, last transition, migration counters. Existing
  functions untouched.

## 13. Phases

- **P1** — config model (Servers/Clusters + converter), router + execution
  wrapper, single-read selection, health monitor (worker, passive + probe,
  hysteresis), polling failover/failback, masking, embedded schema
  (`ConfigSchema`) + static validator (`ConfigValidation`) +
  `RedisUDFConfigPath` / `RedisUDFConfigValidate` / `RedisUDFConfigSchema` /
  `RedisUDFConfigReload` / `RedisUDFConfigSet` (+ pointer file, standard
  directory, optional watcher),
  config error propagation to UDF/RTD (`#CONFIG`), unknown-host learning.
- **P2** — armed handover for RTD and UDF subscriptions, fan-in basics,
  logical dedupe.
- **P3** — multi-read polling merge (`race`/`union`), multi-write, latency
  selection, Sentinel replica discovery, status UDFs.

## 14. Tests

- Unit: converter (string/object forms), resolution order, selection state
  machine (cheapest, tie-break, demotion, promotion, cooldown), masking.
- Unit: `ConfigValidation` rules; the embedded schema resource is present,
  parseable and covers every property of the config model (generated with
  NJsonSchema in the test project only).
- Unit: unknown-host learning — name generation/sanitization, exact-target
  dedupe, comment-preserving round trip, atomic write, read-only path
  behavior.
- Unit: config error propagation — an invalid file blocks UDF/RTD with
  `#CONFIG`, warnings do not block, and the validator still works while the
  active configuration is rejected.
- Unit: config reload — transactional swap (invalid reload keeps the previous
  snapshot), no-op on identical hash, `#CONFIG` cleared by a valid reload,
  removed names enter `#CONFIG`, reconciliation/migration triggered when
  member sets change, learner queue flushed first, explicit source adoption
  and `""` reset, `RedisUDFConfigPath()` reporting for
  standard/explicit/forced/defaults/rejected.
- Unit: config source persistence — pointer write/read in the standard
  directory, `RedisUDFConfigSet` adoption and `""` clearing, invalid target
  never persisted, unwritable-directory failure, startup fail-closed on a
  broken pointer, watcher-triggered transactional reload.
- Unit: health monitor — probe scheduling/jitter, demotion thresholds,
  promotion hysteresis/cooldown, compare-and-swap (no duplicate promotions),
  flap counting.
- Smoke: two local Redis instances on different ports —
  failover, failback, fan-in, multi-write, single-write.
- E2E: extend `test/Run-ExcelE2E.ps1` with a failback scenario (stop the cheap
  instance, verify fallback; start it, verify return).
- Load: hops do not leak connections; timers stay bounded.

## 15. Decisions (closed)

1. `read.multi` is available for every read command; `merge` is an option with
   `race` as the default (simplest, safe) and `union` as the alternative.
2. `write.multi` returns an aggregated value, reduced per function category
   (section 4.4).
3. Fan-in `dedupe` is an option; default `per-source` (current behavior,
   safest — never drops a payload a single source really published).
4. `cost` accepts an integer (sets ingress = egress) or an object
   `{ "ingress": n, "egress": n }`.
5. v2 layout: a single `Default` block for `host`/`timeout`, sections keep
   optional overrides, `CoalesceRealtimeUpdates` canonical inside `RTD`, and
   `Clusters.members` accepts names or inline server objects. No v1
   compatibility: 2.0 is a major version; migration is a one-time file edit.
6. Config source persistence: a single standard directory `C:\RedisExcel\`
   holds `RedisExcel.config.v2.json` and the `RedisExcel.pointer.json` written
   by `RedisUDFConfigSet`; startup order is pointer > standard config.

Remaining implementation details only: optional quorum for multi-write
(default: all required) and the exact generated-name pattern for learned
hosts.

## 16. Planned code layout (v2.0)

The project is flat today and stays flat: new concerns get their own files,
not folders. Names below are provisional; identifiers in English.

### Main project (repository root)

| File | Status | Responsibility |
| --- | --- | --- |
| `ConfigModel.cs` | new | POCOs (`DefaultConfig`, `ServerDef`, `ClusterDef`, `ReadPolicy`, `WritePolicy`, `CostValue`) and the Newtonsoft converters (server string/object, cost int/object) |
| `Config.cs` | rewritten | `AppConfig`: loader, immutable snapshot, standard-directory scan, pointer read/write, reload/`ConfigSet` backing, `MaskConnectionString`, host-reference resolution facade |
| `ConfigValidation.cs` | new | static semantic validator returning `(severity, path, line, message)`; single implementation shared with the console tool |
| `ConfigSchema.cs` | new | `ConfigSchema.Text`, read once from the embedded `RedisExcel.schema.v2.json` |
| `ConfigLearner.cs` | new | UnknownHosts learning: naming, dedupe, batched atomic writer, flush-on-reload |
| `ClusterRouter.cs` | new | `Resolve`/`Execute`/`Subscribe`; per-cluster active member; migration scheduling; reconciliation after reload |
| `HealthMonitor.cs` | new | worker timer, probes, Up/Down state machine, `MemberPromoted` events, config watcher hook |
| `RedisUDFConfig.cs` | new | worksheet functions `RedisUDFConfigPath` / `Validate` / `Schema` / `Reload` / `Set` (same pattern as `ExcelJson.cs`, which already hosts functions outside `RedisUDF.cs`) |
| `RedisRuntime.cs` | updated | wires connection manager + subscriptions + router + health; snapshot swap; shutdown |
| `RedisConnectionManager.cs` | updated | raise connection events for health (today they only log); caching unchanged |
| `RedisSubscriptionManager.cs` | mostly unchanged | physical `(endpoint, channel)` ref-count stays; the router composes the handover above it |
| `RedisRtd.cs` | updated | topics keep the logical name; ticks resolve via router; `#CONFIG` propagation; timer intervals updated on reload |
| `RedisUDF.cs` | updated | call sites via router; read/write classification; `#CONFIG` guard |
| `RedisExcel.schema.v2.json` | new | schema, embedded as a manifest resource |
| `RedisExcel.config.v2.json` | new | v2 sample (the v1-named file is never read) |
| `DESIGN-v2.0.md` | local | this document (untracked) |

`RedisExcel.csproj` changes:

- `<EmbeddedResource Include="RedisExcel.schema.v2.json" />`;
- `<Compile Remove="tools\**" />` — SDK-style globs would otherwise compile
  the console tool into the add-in (`test\**` is already excluded).

### Tooling (optional)

- `tools/ConfigValidator/` console project referencing the add-in project's
  `ConfigValidation` (no duplicated rules); used in CI and by hand.
- The tool project is added to `RedisExcel.sln` for convenience.

### Tests

- `test/RedisExcel.Tests/`: new files `ConfigModelTests.cs`,
  `ConfigValidationTests.cs`, `ConfigReloadTests.cs`, `ConfigLearningTests.cs`,
  `RouterSelectionTests.cs`, `HealthMonitorTests.cs`, alongside the existing
  ones.
- `test/SmokeTests`: two local Redis instances (different ports) for
  failover/failback/fan-in/multi-write.
- `test/Run-ExcelE2E.ps1`: extend with a failback scenario.
- `AGENTS.md`: update the layout table, golden rule 3 ("config v2 only, no v1
  compatibility") and the config search-path line (v1: user profile / Excel
  folder / `C:\Windows`; v2: `C:\RedisExcel\`) when implementation starts.
