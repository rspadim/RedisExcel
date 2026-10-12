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
3. A short wizard asks for the Redis **host / port / password** and writes
   `RedisExcel.json` next to the XLL (the add-in searches its own folder first).
4. On the last page, explains the one manual step: open Excel and add the `.xll`
   via **File > Options > Add-ins > Go... > Browse...** (Excel exposes no safe
   automation for that).
5. The uninstaller asks whether to keep your `RedisExcel.json`.

## Silent / "files only" mode

No wizard, no Excel prompt - just copy the files and use the given (or default)
host:

```powershell
RedisExcelSetup-1.5.1.exe /VERYSILENT /SUPPRESSMSGBOXES /DIR="C:\RedisExcel" /HOST=localhost /PORT=6379 /PASSWORD=secret
```

`/HOST`, `/PORT`, `/PASSWORD` are optional (defaults: `localhost`, `6379`,
empty). The installer is a normal `.exe`, so you can also install by hand: copy
`RedisExcel-packed.xll` (or `RedisExcel64-packed.xll`), `RedisExcel.dll`,
`NLog.config` and `RedisExcel.json` into one folder and register the XLL from
Excel.

## Build it yourself

```powershell
# One-time: install Inno Setup 6 (winget or choco)
choco install innosetup -y

# Build the add-in first (the installer packages bin\Release\net48\publish)
dotnet build RedisExcel.sln -c Release

# Compile the installer (version is injected; use the tag without the 'v')
iscc "/DAppVersion=1.5.1" "installer\RedisExcel.iss"
# -> installer\output\RedisExcelSetup-1.5.1.exe
```

CI builds the installer on every run (so a broken `.iss` fails fast) and attaches
`RedisExcelSetup-<version>.exe` to the GitHub release on a `v*` tag.

## Notes / limits

- The XLL is **not code-signed**, so Windows SmartScreen and some antivirus
  products may warn; the installer cannot fix that without a code-signing
  certificate.
- Excel must be **closed** while installing/updating (the `.XLL` file is locked
  while loaded).
- The add-in folder is only the *first* place searched for `RedisExcel.json`;
  it still falls back to `%USERPROFILE%`, the Excel folder and `C:\Windows`.
