# RedisExcel installer

A small [Inno Setup 6](https://jrsoftware.org/isinfo.php) installer that copies
the add-in into one folder, writes a starter `RedisExcel.json`, and keeps both
XLLs (removing the one that does not match your Office bitness). It runs
**per-user** (no admin), because the add-in registers its own COM ProgId under
`HKCU` when Excel loads it.

## What it does

1. Copies `RedisExcel.dll` + the packed XLLs + `NLog.config` into a folder you
   pick (default `C:\RedisExcel`).
2. Detects the Office bitness and keeps only the matching XLL
   (`RedisExcel64-packed.xll` for 64-bit, `RedisExcel-packed.xll` for 32-bit).
3. The wizard first asks for a **mode**:
   - **Simple** (default): only the Redis **host / port / password**;
   - **Advanced**: a single editable page with the whole `RedisExcel.json`
     pre-filled and **commented** (every key with its default and a one-line
     explanation - timeouts, rates, `ExcelUpdateStyle`, `MessageCounterThreshold`,
     `Servers` aliases, `UpdateCheck`, `SkipRepeatedMessages`,
     `CoalesceRealtimeUpdates`, `ConflationMs`, `PublishDedupCacheSize`,
     `SyncWrite`, `AsyncWrites`).

   It writes `RedisExcel.json` next to the XLL (the add-in searches its own
   folder first) and a final read-only page summarizes what will be written.
   **Upgrading:** when a `RedisExcel.json` already exists there, the installer
   leaves it alone *if the add-in accepts it*. A file the add-in rejects (the
   pre-1.6 shape, or any wrong value - every cell would show `#CONFIG`) is backed
   up to `RedisExcel.json.pre-1.6.bak` and replaced by a working file, and the
   old server address is carried over: the host field is pre-filled with it (and
   a `/VERYSILENT` install reuses it automatically), so the migration never
   silently moves the user to localhost.
4. **Optional: installs a server.** Choose *No*, *Docker: Redis (latest)*,
   *Docker: Valkey (latest)*, *WSL2: Redis*, *WSL2: Valkey* or *Memurai*. It runs
   `scripts\install-server.ps1`, which uses Docker if available and otherwise
   falls back to WSL or points you at Memurai - so you do not need a server
   beforehand.
5. **End-of-install check.** It authenticates when you gave a password (a server
   with `requirepass` answers `-NOAUTH` to a bare PING, so an unauthenticated
   probe used to report a healthy install as unreachable), then PINGs the server
   and, if you ticked the box, opens a visible Excel workbook (XLL loaded) titled
   *"RedisExcel installed successfully!"* with the server time and a live
   `Set`/`Get` round trip. The messages and the sample sheet are **English**: the
   add-in is English-only inside Excel on purpose, and the sheet is the first
   thing a new user sees. With a password the sheet uses the no-argument form of
   the UDFs, so the secret never lands in a cell formula. The password reaches the
   check through a private temp file, not the argument list (a password with a
   quote or a space cannot be passed as one argument safely, and a command line is
   readable by other local users).
   The connection page also has a **Test connection** button that runs the same
   probe on what is in the fields *right now*, without leaving the page and
   without writing `RedisExcel.json`. It says up front that it waits up to 5
   seconds, and on failure names the reason (unknown host, timeout, refused,
   rejected password) from the helper's stable exit code - never from its console
   text, which a redirect re-encodes.
6. On the last page it explains the one manual step: add the `.xll` in Excel via
   **File > Options > Add-ins > Go... > Browse...** (Excel exposes no safe
   automation for that).
7. The uninstaller asks whether to keep your `RedisExcel.json`.

## Silent / "files only" mode

No wizard, no Excel prompt - just copy the files and use the given (or default)
host:

```powershell
RedisExcelSetup-<version>.exe /VERYSILENT /SUPPRESSMSGBOXES /DIR="C:\RedisExcel" /HOST=localhost /PORT=6379 /PASSWORD=secret /SERVER=none /OPENEXCEL=0
```

`/HOST`, `/PORT`, `/PASSWORD`, `/SERVER` (`none`|`docker-redis`|`docker-valkey`|
`wsl-redis`|`wsl-valkey`|`memurai`) and `/OPENEXCEL` (`0`|`1`) are optional
(defaults: `localhost`, `6379`, empty, `none`, `0`). An unknown `/SERVER` value
aborts the whole install with a non-zero exit code (rather than silently
installing no server), so a typo cannot pass unnoticed in automation. The installer is a normal
`.exe`, so you can also install by hand: copy `RedisExcel-packed.xll` (or
`RedisExcel64-packed.xll`), `RedisExcel.dll`, `NLog.config` and `RedisExcel.json`
into one folder and register the XLL from Excel.

> **Passwords**: the value is embedded in the Redis connection string, which
> StackExchange.Redis splits on commas, so a password containing a **comma**
> cannot be represented (the client silently uses only the part before it, and
> there is no quoting that works). A password with `;` is fine. `Validate-Config.cmd`
> warns when a connection string looks like a truncated password. For such a
> password, set it in `RedisExcel.json` by hand with a form your client accepts.

### Languages

The wizard is localized in 10 languages (English - the default, 简体中文,
繁體中文, Português do Brasil, Español, Русский, Français, 한국어, Deutsch,
日本語); it picks the language of your Windows or you can change it on the first
screen. This is **install-time only**: the add-in always uses English inside
Excel, so volatile functions return the same strings on every machine.

> **Vendored translations:** eight of the ten language files ship with Inno Setup;
> the two Chinese ones are vendored under `installer/languages/` because the
> official Inno package has no Chinese. `ChineseSimplified.isl` matches Inno
> 6.5+; `ChineseTraditional.isl` is a community translation based on Inno 6.4,
> so `iscc` prints ~23 warnings for it (four unknown message names plus ~19
> new 6.5 messages that fall back to English). Those messages belong to the
> download/extraction/verification features this installer does not use, so the
> effect is cosmetic. **When a 6.5+ Traditional Chinese file is available, drop it
> in and re-run `iscc` - the warnings should go to zero.**

The silent uninstaller can also remove a server it can still see. It only
removes what you name:

```powershell
unins000.exe /VERYSILENT /REMOVESERVERS=docker-redis,docker-valkey
```

(Interactive uninstalls show a page with the detected servers instead; by default
nothing is removed.)

## Helper scripts

The installer ships these under `C:\RedisExcel\` (all usable on their own):

- `scripts\install-server.ps1 -Kind docker-valkey -Port 6379` - brings up a
  server (kinds: `docker-redis`, `docker-valkey`, `wsl-redis`, `wsl-valkey`,
  `memurai`; `none` does nothing). Uses the `:latest` Docker image by default;
  pass `-Image redis:7-alpine` to pin a tag. Exit `3` means the method is
  unavailable here.
- `scripts\test-connection.ps1 -ServerHost localhost:6379 [-User u] [-Password p]` -
  PINGs the server (honours AUTH), exit `0`/`1`.
- `scripts\verify-connection.ps1 -AddInDir "C:\RedisExcel" -ServerHost localhost:6379 -OpenExcel` -
  PINGs and opens the sample sheet.
- `scripts\validate-config.ps1` (or double-click `Validate-Config.cmd`) -
  checks `RedisExcel.json` against the shipped `RedisExcel.schema.json` (the same
  schema the add-in embeds) and prints an OK/FAIL report.
- `scripts\detect-servers.ps1` / `scripts\uninstall-server.ps1` - used by the
  uninstaller.

## Uninstall

Run **Start Menu > Uninstall RedisExcel** (or Apps & features). It removes the
files, the Excel auto-load entry and the `RedisRtd` COM ProgId. It then:

- offers to remove any Redis-compatible **servers** it finds (Docker container
  `redisexcel`, Redis/Valkey in WSL2, Memurai) - **nothing is removed unless you
  tick it** (default: keep everything);
- asks whether to keep your `RedisExcel.json`.

## Build it yourself

```powershell
# One-time: install Inno Setup 6 (winget or choco)
choco install innosetup -y

# Build the add-in first (the installer packages bin\Release\net48\publish)
dotnet build RedisExcel.sln -c Release

# Compile the installer (version is injected; use the tag without the 'v')
iscc "/DAppVersion=<version>" "installer\RedisExcel.iss"
# -> installer\output\RedisExcelSetup-<version>.exe
# CI also copies it to installer\output\RedisExcelSetup.exe (version-less), so
# releases/latest/download/RedisExcelSetup.exe is a stable direct link: the
# versioned file name under /releases/latest/ would 404 as soon as the next tag
# is published.
```

CI builds the installer on every run (so a broken `.iss` fails fast) and attaches
both `RedisExcelSetup-<version>.exe` and the version-less `RedisExcelSetup.exe` to
the GitHub release on a `v*` tag.

## Notes / limits

- The XLL is **not code-signed**, so Windows SmartScreen and some antivirus
  products may warn; the installer cannot fix that without a code-signing
  certificate.
- The end-of-install check probes the **wizard's** host/port/password. On the
  advanced page you can edit the whole `RedisExcel.json` after those fields are
  filled, and an edit made *only* there (say a different host) is not what the
  check probes - validate such a file with `scripts\Validate-Config.cmd`.
- Excel must be **closed** while installing/updating (the `.XLL` file is locked
  while loaded).
- The add-in folder is only the *first* place searched for `RedisExcel.json`;
  it still falls back to `%USERPROFILE%`, the Excel folder and `C:\Windows`.
