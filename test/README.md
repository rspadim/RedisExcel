# Tests

Three test layers plus a load test harness. See `AGENTS.md` for the full
guidance and the Excel automation lessons.

## 1. Unit tests (no Redis, no Excel)

```powershell
dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release
```

## 2. Smoke tests (requires Redis, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
```

## 3. Excel end-to-end (requires Excel + Redis + built XLLs)

```powershell
dotnet build RedisExcel.sln -c Release
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

The script builds the test workbook (`RedisExcel.Test.xlsx`), asserts UDF/RTD
values and reproduces the Pub/Sub regression scenario (workbook copy/close and
connection blip). For remote hosts the workbook goes to `%TEMP%` and the
destructive `CLIENT KILL` step is skipped automatically.

Useful parameters: `-RedisHost`, `-KeyPrefix`, `-RealChannel`, `-RealPattern`,
`-SkipClientKill`, `-RedisCli`, `-KeepExcelOpen`. Supplying a custom
`-RedisCli` (e.g. `docker exec my-redis redis-cli`) disables the automatic
`CLIENT KILL` step.

> Building the test workbook only uses `localhost:6379` and
> `test.redisexcel.*` keys/channels. Never commit workbooks or configs that
> reference private hosts.

## 4. Load tests (requires Redis)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "127.0.0.1:6379" 10 2 1
```

Modes: `manager` (RedisSubscriptionManager) and `raw` (plain
StackExchange.Redis baseline). Parameters: host, seconds, publisher threads
(0 = listen-only with an external generator such as
`redis-benchmark -t publish`), listeners, pattern and an optional `[channel]`
(7th argument; defaults to a random `load:<guid>` channel). It reports
throughput, allocated bytes per received message and GC counts. Read-only stress runs
against real servers are allowed, but pass the host only as a command-line
argument (never commit it) and keep the runs short.
