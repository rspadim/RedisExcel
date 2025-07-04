# RedisExcel

![RedisExcel Overview](redisexcel.png)

TL-DR: Read Installing in Excel and [Download Releases](https://github.com/rspadim/RedisExcel/releases)

Integration between Redis and Excel using RTD (Real-Time Data) and UDF (User Defined Functions), developed in C# (.NET Framework 4.8) with [ExcelDna](https://github.com/Excel-DNA/ExcelDna).
Supports Pub/Sub and polling with `GET`, `HGET`, `HGETALL`, `SUB`, `PSUB` commands and various UDF operations to interact with Redis.

---

## 🧹 Installing in Excel

### Quick Installation

1. Download the latest release from GitHub (`.xll`, `.dll`) - [Download Releases](https://github.com/rspadim/RedisExcel/releases)
2. Download [NLog config example files](https://github.com/rspadim/RedisExcel/blob/main/NLog.config)
3. Download [RedisExcel.json file](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.json)
4. Place (`.xll`, `.dll`, `NLog.config`) files in the **same local folder**
5. In Excel: `File > Options > Add-ins`

   * Click **Go...**, then **Browse...**, and select the `.xll` file

6. The `RedisExcel.json` file can be placed in - [Example](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.json)
>
> * The user folder (`%USERPROFILE%`) - Example: c:\Users\rspadim\RedisExcel.json
> * The Excel folder
> * Or `C:\Windows\RedisExcel.json`

---

## 🚀 Example: Publishing from Python

```python
import redis

r = redis.StrictRedis(host='localhost', port=6379, db=0)

# Set a key
r.set("preco_btc", "67000.50")

# Publish to a channel
r.publish("canal_alerta", "ALTA")
```

---

## 🚀 Example: Publishing from Excel (send data)

```Excel
=RedisUDFSet("preco_btc", "67000.50")                          // send SET message (key-value database)
=RedisUDFChannelPublish("canal_alerta", "ALTA")                // send PUB/SUB message
=RedisUDFChannelPublishJSON("range_of_data", A1:C20)           // send PUB/SUB Matrix, encoded as JSON
```

---

## 🚀 Example: Subscribe and UDF from Excel (fetch data)

```Excel
=RedisUDFGet("preco_btc", "67000.50")                          // receive GET (key-value database) with UDF (no character limit)
=RTD("RedisRtd", , "GET", "canal_alerta")                      // receive GET (key-value database - automatic pooling) with RTD (max of 255 characters)
=RedisUDFChannelLatest("canal_alerta")                         // receive PUB/SUB with UDF (no character limit)
=RTD("RedisRtd", , "SUB", "canal_alerta")                      // receive PUB/SUB with RTD (max of 255 characters)
=RedisUDFJSONToMatrix(RedisUDFChannelLatest("range_of_data"))  // receive PUB/SUB Matrix using UDF (no RTD)
```

---

## 🛠️ Build Instructions

### Requirements

* Visual Studio 2022+
* .NET Framework 4.8
* NuGet packages:

  * ExcelDna.AddIn
  * ExcelDna.Integration
  * ExcelDna.IntelliSense
  * StackExchange.Redis
  * NLog
  * Newtonsoft.Json

### Steps

```bash
git clone https://github.com/rspadim/RedisExcel.git
```

1. Open the project in Visual Studio
2. Restore NuGet packages
3. Build in `Release` mode — `.xll` will be generated in `bin\Release`

---

## 🔎 Excel Examples

### Pub/Sub

```excel
=RTD("RedisRtd", , "SUB", "canal_alerta", "localhost:6379")
=RTD("RedisRtd", , "SUB", "canal_alerta", "dev") // can find an alias using `RedisExcel.json` file
=RTD("RedisRtd", , "SUB", "canal_alerta")  // uses default host if omitted
```

You can find other connection string formats in the [StackExchange.Redis configuration manual](https://stackexchange.github.io/StackExchange.Redis/Configuration).

---

## 📊 Available RTD Functions

| Function                    | Description                            | 
| --------------------------- | -------------------------------------- | 
| RedisRTDConnectionCount     | Number of active Redis connections     | 
| RedisRTDSubscriptionCount   | Number of active Redis subscriptions   | 
| RedisRTDTopicCount          | Total number of RTD topics registered  | 
| RedisRTDChannelCount        | Number of distinct subscribed channels | 
| RedisRTDDefaultHost         | Current default Redis host             | 
| RedisRTDExcelUpdateInterval | Excel update interval (ms)             | 
| RedisRTDRedisUpdateInterval | Redis polling interval (ms)            | 
| RedisRTDRealTimeUpdates     | Is real-time update enabled? (bool)    | 

---

## RTD Parameters Reference

When calling the Excel RTD function like this:

```excel
=RTD("RedisRtd", , param1, param2, param3, param4)
```

Each parameter has a specific meaning depending on the RTD command.

### RTD Parameters Breakdown

| Param  | Description                                               |
| ------ | --------------------------------------------------------- |
| param1 | Command: GET, HGET, HGETALL, SUB, PSUB                    |
| param2 | Key, hash, or channel name                                |
| param3 | Host (optional for GET/HGETALL/SUB/PSUB), or field (HGET) |
| param4 | Host (only used in HGET if param3 is the field name)      |

> If `param3` or `param4` are omitted, the default host will be used. You can use aliases from your `RedisExcel.json`.

### Supported RTD Commands

| Command | Description                      | Arguments            |
| ------- | -------------------------------- | -------------------- |
| GET     | Polls a Redis key                | key, \[host]         |
| HGET    | Polls a field in a Redis hash    | hash, field, \[host] |
| HGETALL | Polls all fields in a Redis hash | hash, \[host]        |
| SUB     | Subscribes to a Redis channel    | channel, \[host]     |
| PSUB    | Subscribes to a Redis pattern    | pattern, \[host]     |

> All commands support specifying either a full connection string or a named host defined in the `RedisExcel.json` file.

## 💡 Available UDF Functions

Functions to use directly in Excel cells:

| Function                         | Description                      | Parameters                                |
| -------------------------------- | -------------------------------- | ----------------------------------------- |
| RedisUDFGet                      | Get the value of a key           | key, optionalHost                         |
| RedisUDFSet                      | Set the value of a key           | key, value, optionalHost                  |
| RedisUDFSetJSON                  | Set JSON-encoded data            | key, matrix, optionalHost                 |
| RedisUDFSetKV / SetKVPair        | Set key-value pairs              | keys, values / pairs, optionalHost        |
| RedisUDFGetMultiple              | Get multiple keys                | keys\[], multipleColumnsOpt, optionalHost |
| RedisUDFExists / ExistsMultiples | Check key existence              | key / keys\[], optionalHost               |
| RedisUDFTTL / TTLMultiples       | Time-to-live (TTL) for keys      | key / keys\[], optionalHost               |
| RedisUDFHashSet/Get/...          | Redis Hash operations            | see combinations                          |
| RedisUDFChannelPublish/...       | Pub/Sub operations               | channel, message, optionalHost            |
| RedisUDFJSONToMatrix             | Convert JSON → Excel matrix      | json, nullValue                           |
| RedisUDFMatrixToJSON             | Convert Excel matrix → JSON      | matrix                                    |
| RedisUDFServerTime               | Redis server current time        | optionalHost                              |
| RedisUDFKeys                     | List keys by pattern             | pattern, optionalHost                     |
| RedisUDFConnectionCount          | Number of active UDF connections | None                                      |

---

## 📝 Configuration Files

### Example: NLog.config

```xml
<?xml version="1.0" encoding="utf-8" ?>
<nlog xmlns="http://www.nlog-project.org/schemas/NLog.xsd"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

  <targets>
    <target name="file" xsi:type="File"
            fileName="RedisExcel.log"
            layout="${longdate}|${level:uppercase=true}|${logger}|${message} ${exception:format=toString}"
            archiveFileName="RedisExcel.${environment-user}.{#}.log"
            archiveAboveSize="104857600"
            archiveNumbering="Rolling"
            maxArchiveFiles="5"
            concurrentWrites="true"
            keepFileOpen="true"
            encoding="utf-8" />
  </targets>

  <rules>
    <logger name="*" minlevel="Debug" writeTo="file" />
  </rules>
</nlog>
```

### JSON Configuration (RedisExcel.json)

```json
{
  "RTD": {
    "host": "localhost:6379",
    "timeout": 1000,
    "RedisUpdateRateMs": 1000,
    "ExcelUpdateRateMS": 100,
    "MessageCounterThreshold": 1000,
    "ExcelUpdateStyle": "Automatic",
    "UseGetMultiple": true
  },
  "UDF": {
    "host": "localhost:6379",
    "timeout": 1000
  },
  "Servers": {
    "prod": "localhost:6379,defaultDatabase=0",
    "dev": "localhost:6379,defaultDatabase=1"
  }
}
```

---

## ♻️ Force Update in Excel

* Press `F9`
* Or use VBA:

```vba
Dim nextUpdate As Date

Sub RefreshRTD()
    Sheet1.Calculate
    nextUpdate = Now + TimeValue("00:00:01")
    Application.OnTime nextUpdate, "RefreshRTD"
End Sub

Sub StopUpdate()
    On Error Resume Next
    Application.OnTime nextUpdate, "RefreshRTD", , False
End Sub
```

---

## 📬 Contact

Open an *issue* on GitHub or email as instructed in the repository.
