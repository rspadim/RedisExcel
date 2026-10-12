; RedisExcel installer + uninstaller (Inno Setup 6).
;
; Build:  ISCC /DAppVersion=1.5.1 .\installer\RedisExcel.iss
; CI builds it on every run and attaches RedisExcelSetup-<version>.exe to a
; v* tag release (see .github/workflows/build.yml).
;
; Flow (in this order, as the wizard runs):
;   1. asks whether to install a server (Docker/WSL/Memurai) and the connection;
;   2. INSTALLS THE SERVER (if chosen);
;   3. INSTALLS RedisExcel (copies the files, keeps the matching XLL, writes
;      RedisExcel.json) and registers the XLL so Excel loads it on start;
;   4. TESTS the configuration (visible OK/fail) and, if asked, opens a sample
;      workbook.
;
; The uninstaller (Start Menu "Uninstall RedisExcel" or Apps & features) removes
; the files, the Excel auto-load entry and the COM ProgId; it asks whether to
; keep your RedisExcel.json.
;
; Silent mode: RedisExcelSetup.exe /VERYSILENT /DIR=... /HOST=... /PORT=...
;   /PASSWORD=... /SERVER=<none|docker-redis|docker-valkey|wsl-redis|wsl-valkey|memurai>
;   /OPENEXCEL=<0|1>

#define AppName "RedisExcel"
#define AppPublisher "rspadim"
#define AppUrl "https://github.com/rspadim/RedisExcel"
#ifndef AppVersion
  #define AppVersion "dev"
#endif
#ifndef BuildDir
  #define BuildDir "..\bin\Release\net48\publish"
#endif
#ifndef RepoRoot
  #define RepoRoot ".."
#endif

[Setup]
AppId={{7F3B9C2A-4E1D-4C7A-9B2E-6A1F0D5C8E33}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
DefaultDirName=C:\RedisExcel
DisableDirPage=no
; Per-user install (no admin): the add-in registers its own COM ProgId under HKCU.
PrivilegesRequired=lowest
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir={#RepoRoot}\installer\output
OutputBaseFilename=RedisExcelSetup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName} {#AppVersion}
UninstallDisplayIcon={app}\RedisExcel64-packed.xll

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "addinconfig"; Description: "Create RedisExcel.json in the install folder (host settings)"; GroupDescription: "Configuration:"
Name: "nlogconfig";  Description: "Create NLog.config (write a log file next to the add-in)"; GroupDescription: "Configuration:"
Name: "excelautoload"; Description: "Register the add-in so Excel loads it automatically (recommended)"; GroupDescription: "Excel:"

[Icons]
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

[Files]
Source: "{#BuildDir}\RedisExcel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\NLog.config"; DestDir: "{app}"; Flags: onlyifdoesntexist; Tasks: nlogconfig
; Helper scripts: server setup, connection test, end-of-install verification,
; server detection/removal for the uninstaller, and the config validator.
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\test-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\verify-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\detect-servers.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\uninstall-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\validate-config.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
; The full JSON Schema (the same one the add-in embeds) and a double-click tool.
Source: "{#RepoRoot}\RedisExcel.schema.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\Validate-Config.cmd"; DestDir: "{app}"; Flags: ignoreversion
; A temporary copy of the server helpers, so the server can be installed
; BEFORE the add-in files are extracted (see PrepareToInstall).
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "{#RepoRoot}\installer\test-connection.ps1"; DestDir: "{tmp}"; Flags: dontcopy

[UninstallDelete]
; The log file our NLog.config writes next to the add-in.
Type: files; Name: "{app}\RedisExcel.log"
Type: filesandordirs; Name: "{app}\scripts"

[Code]
var
  Is64BitOffice: Boolean;
  ConnPage: TInputQueryWizardPage;
  ServerPage: TInputOptionWizardPage;
  OptionsPage: TInputOptionWizardPage;
  ServerInstallCode: Integer;
  TestOk: Boolean;
  InstallCompleted: Boolean;
  UninstallServersPage: TInputOptionWizardPage;
  ServerKinds: TArrayOfString;
  ServerLabels: TArrayOfString;
  ServerPresent: TArrayOfBoolean;

// ---- Office bitness / versions --------------------------------------------
function Detect64BitOffice(): Boolean;
var
  Names: TArrayOfString;
  I: Integer;
  V, P: String;
begin
  Result := True;
  P := 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration';
  if RegQueryStringValue(HKLM, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else if RegQueryStringValue(HKCU, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else if RegGetSubkeyNames(HKCU, 'SOFTWARE\Microsoft\Office', Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
      if RegQueryStringValue(HKCU,
           'SOFTWARE\Microsoft\Office\' + Names[I] + '\Common\General', 'Bitness', V) then
      begin
        Result := (CompareText(V, 'x64') = 0);
        Exit;
      end;
  end;
end;

// Office version subkeys are like "16.0"; use every one that has Excel.
function OfficeExcelKeys(): TArrayOfString;
var
  Names: TArrayOfString;
  I: Integer;
  Acc: TArrayOfString;
begin
  SetArrayLength(Acc, 0);
  if RegGetSubkeyNames(HKCU, 'SOFTWARE\Microsoft\Office', Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if (Length(Names[I]) > 0) and (Names[I][1] >= '0') and (Names[I][1] <= '9') then
      begin
        SetArrayLength(Acc, GetArrayLength(Acc) + 1);
        Acc[GetArrayLength(Acc) - 1] := Names[I];
      end;
  Result := Acc;
end;

// ---- Register the XLL as an Excel auto-open add-in -------------------------
function RegisterExcelAutoLoad(const XllPath: String): Boolean;
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
  Found: Boolean;
begin
  Result := False;
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    // Already registered?
    Found := False;
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and (CompareText(Cur, XllPath) = 0) then
      Found := True;
    N := 1;
    while (not Found) and (N < 100) do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if not RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
      begin
        if RegWriteStringValue(HKCU, BaseKey, ValName, XllPath) then Result := True;
        Found := True; // stop scanning this key (wrote it or failed)
      end
      else if CompareText(Cur, XllPath) = 0 then
        Found := True;
      N := N + 1;
    end;
    if Found and (not Result) then
      Result := True; // already present in this Office version
  end;
end;

procedure UnregisterExcelAutoLoad();
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
begin
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and
       (Pos('RedisExcel', Cur) > 0) then
      RegDeleteValue(HKCU, BaseKey, 'OPEN');
    N := 1;
    while N < 100 do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
      begin
        if Pos('RedisExcel', Cur) > 0 then RegDeleteValue(HKCU, BaseKey, ValName);
      end;
      N := N + 1;
    end;
  end;
end;

// ---- Wizard pages ---------------------------------------------------------
function ParamOr(Param, Default: String): String;
begin
  Result := ExpandConstant('{param:' + Param + '|' + Default + '}');
end;

procedure InitializeWizard();
begin
  Is64BitOffice := Detect64BitOffice();

  ServerPage := CreateInputOptionPage(wpSelectTasks,
    'Redis server', 'Do you want to install a Redis-compatible server now?',
    'RedisExcel needs a Redis (or Valkey) server. Choose an option if you do not ' +
    'already have one - the server is installed BEFORE RedisExcel, and the ' +
    'connection is tested at the end. Docker is the simplest.',
    True, False);
  ServerPage.Add('No - I already have a server (or will set it up later)');
  ServerPage.Add('Yes - Docker: Redis (redis:7-alpine)');
  ServerPage.Add('Yes - Docker: Valkey (valkey/valkey:8-alpine)');
  ServerPage.Add('Yes - WSL2: Redis (apt install redis-server)');
  ServerPage.Add('Yes - WSL2: Valkey (apt install valkey-server)');
  ServerPage.Add('Yes - Memurai (native Windows service, needs admin)');
  ServerPage.SelectedValueIndex := 0;

  ConnPage := CreateInputQueryPage(ServerPage.ID,
    'Redis connection', 'Where should RedisExcel connect?',
    'Press Next for the default (localhost:6379). You can change this later by ' +
    'editing RedisExcel.json.');
  ConnPage.Add('Host (or host:port):', False);
  ConnPage.Add('Port:', False);
  ConnPage.Add('User (optional, for Redis 6 ACL):', False);
  ConnPage.Add('Password (optional):', True);
  ConnPage.Values[0] := ParamOr('HOST', 'localhost');
  ConnPage.Values[1] := ParamOr('PORT', '6379');
  ConnPage.Values[2] := ParamOr('USER', '');
  ConnPage.Values[3] := ParamOr('PASSWORD', '');

  OptionsPage := CreateInputOptionPage(ConnPage.ID,
    'Finish', 'Extra checks',
    'The installer tests the connection when it finishes. Optionally open a ' +
    'sample Excel workbook showing it working.',
    False, False);
  OptionsPage.Add('Open a sample RedisExcel workbook at the end');
  OptionsPage.Values[0] := (ParamOr('OPENEXCEL', '0') = '1');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := WizardSilent() and
            ((PageID = ServerPage.ID) or (PageID = ConnPage.ID) or (PageID = OptionsPage.ID));
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = ConnPage.ID) and (Trim(ConnPage.Values[0]) = '') then
  begin
    MsgBox('Please enter a Redis host (for example localhost).', mbError, MB_OK);
    Result := False;
  end;
end;

// ---- Values ---------------------------------------------------------------
function FieldValue(Index: Integer; const ParamName, Default: String): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[Index])
  else
    Result := ParamOr(ParamName, Default);
end;

function HostString(): String;
var
  H, Port, User, Pw: String;
begin
  H := FieldValue(0, 'HOST', 'localhost');
  Port := FieldValue(1, 'PORT', '6379');
  User := FieldValue(2, 'USER', '');
  Pw := FieldValue(3, 'PASSWORD', '');
  if (Pos(':', H) = 0) and (Port <> '') then Result := H + ':' + Port else Result := H;
  if User <> '' then Result := Result + ',user=' + User;
  if Pw <> '' then Result := Result + ',password=' + Pw;
end;

function ServerKind(): String;
var idx: Integer;
begin
  if WizardSilent() then begin Result := ParamOr('SERVER', 'none'); Exit; end;
  if ServerPage = nil then begin Result := 'none'; Exit; end;
  idx := ServerPage.SelectedValueIndex;
  case idx of
    1: Result := 'docker-redis';
    2: Result := 'docker-valkey';
    3: Result := 'wsl-redis';
    4: Result := 'wsl-valkey';
    5: Result := 'memurai';
  else
    Result := 'none';
  end;
end;

function OpenExcelChecked(): Boolean;
begin
  if WizardSilent() then
    Result := (ParamOr('OPENEXCEL', '0') = '1')
  else
    Result := (OptionsPage <> nil) and OptionsPage.Values[0];
end;

function ConfigJson(): String;
var HostLine: String;
begin
  HostLine := HostString();
  Result :=
    '{' + #13#10 +
    '  "RTD": {' + #13#10 +
    '    "host": "' + HostLine + '",' + #13#10 +
    '    "timeout": 1000,' + #13#10 +
    '    "RedisUpdateRateMs": 1000,' + #13#10 +
    '    "ExcelUpdateRateMs": 100,' + #13#10 +
    '    "MessageCounterThreshold": 10000,' + #13#10 +
    '    "ExcelUpdateStyle": "Automatic",' + #13#10 +
    '    "UseGetMultiple": true' + #13#10 +
    '  },' + #13#10 +
    '  "UDF": {' + #13#10 +
    '    "host": "' + HostLine + '",' + #13#10 +
    '    "timeout": 1000' + #13#10 +
    '  },' + #13#10 +
    '  "UpdateCheck": true,' + #13#10 +
    '  "SkipRepeatedMessages": true,' + #13#10 +
    '  "CoalesceRealtimeUpdates": true,' + #13#10 +
    '  "PublishDedupCacheSize": 10000,' + #13#10 +
    '  "SyncWrite": "fireforget",' + #13#10 +
    '  "AsyncWrites": false' + #13#10 +
    '}' + #13#10;
end;

// ---- Run PowerShell helpers ----------------------------------------------
function RunScript(ScriptName, Args: String): Integer;
var
  Script, Params: String;
  Code: Integer;
begin
  Script := ExpandConstant('{app}\scripts\') + ScriptName;
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" ' + Args;
  if Exec('powershell.exe', Params, ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    Result := Code
  else
    Result := -1;
end;

function ServerName(Kind: String): String;
begin
  if Kind = 'docker-redis' then Result := 'Redis (Docker)'
  else if Kind = 'docker-valkey' then Result := 'Valkey (Docker)'
  else if Kind = 'wsl-redis' then Result := 'Redis (WSL)'
  else if Kind = 'wsl-valkey' then Result := 'Valkey (WSL)'
  else if Kind = 'memurai' then Result := 'Memurai'
  else Result := 'server';
end;

// ---- Steps: server (before files) -> files -> config + test ---------------
// PrepareToInstall runs BEFORE the files are extracted, so the server is
// installed first (the helpers are extracted to {tmp} on demand).
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Kind, Script, Params: String;
  Code: Integer;
begin
  Result := '';
  Kind := ServerKind();
  if Kind = 'none' then Exit;

  ExtractTemporaryFile('install-server.ps1');
  Script := ExpandConstant('{tmp}\install-server.ps1');
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" -Kind ' + Kind +
            ' -Port ' + FieldValue(1, 'PORT', '6379');
  if Exec('powershell.exe', Params, ExpandConstant('{tmp}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    ServerInstallCode := Code
  else
    ServerInstallCode := -1;

  if (ServerInstallCode <> 0) and (not WizardSilent()) then
    MsgBox('Could not set up ' + ServerName(Kind) + ' automatically (exit ' +
           IntToStr(ServerInstallCode) + ').' + #13#10 +
           'You can install a server later and point RedisExcel.json at it.',
           mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath, Lines: String;
begin
  if CurStep = ssPostInstall then
  begin
    // 2) RedisExcel files are copied: keep the matching XLL and write the config.
    if Is64BitOffice then
      DeleteFile(ExpandConstant('{app}\RedisExcel-packed.xll'))
    else
      DeleteFile(ExpandConstant('{app}\RedisExcel64-packed.xll'));

    if WizardIsTaskSelected('addinconfig') then
    begin
      ConfigPath := ExpandConstant('{app}\RedisExcel.json');
      if not FileExists(ConfigPath) then
      begin
        Lines := ConfigJson();
        if not SaveStringToFile(ConfigPath, Lines, False) then
          MsgBox('Could not write ' + ConfigPath + '. Create it manually from the sample.',
                 mbInformation, MB_OK);
      end;
    end;

    if WizardIsTaskSelected('excelautoload') then
    begin
      if Is64BitOffice then
        RegisterExcelAutoLoad(ExpandConstant('{app}\RedisExcel64-packed.xll'))
      else
        RegisterExcelAutoLoad(ExpandConstant('{app}\RedisExcel-packed.xll'));
    end;

    InstallCompleted := True;
  end;
end;

// ---- Finish: TEST the configuration, then optionally open a sample sheet ---
procedure DeinitializeSetup();
var
  HostPort, Args, Msg: String;
begin
  // Only on a successful install (never during uninstall/cancel).
  if not InstallCompleted then Exit;

  HostPort := HostString();
  if Pos(',', HostPort) > 0 then HostPort := Copy(HostPort, 1, Pos(',', HostPort) - 1);
  Args := '-ServerHost "' + HostPort + '"';
  TestOk := (RunScript('test-connection.ps1', Args) = 0);

  if not TestOk then
    Msg := 'Instalado, mas o Redis em ' + HostPort + ' nao respondeu ao PING.' + #13#10 +
           'Verifique se o servidor esta rodando (docker ps) e ajuste o ' +
           'RedisExcel.json se ele estiver em outro host/porta.'
  else
    Msg := 'Tudo certo! O Redis respondeu em ' + HostPort + '.' + #13#10 + #13#10 +
           'Para usar: abra o Excel e, se o add-in nao carregar sozinho, va em ' +
           'File > Options > Add-ins > Go... > Browse... e escolha o .xll em ' +
           ExpandConstant('{app}') + '.';

  if not WizardSilent() then
    MsgBox(Msg, mbInformation, MB_OK);

  if OpenExcelChecked() then
    RunScript('verify-connection.ps1',
      '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '" -OpenExcel -Quiet');
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpFinished) and (not WizardSilent()) then
    WizardForm.FinishedLabel.Caption :=
      'RedisExcel is installed in:' + #13#10 + ExpandConstant('{app}') + #13#10 + #13#10 +
      'If the add-in does not load automatically, enable it in Excel:' + #13#10 +
      'File > Options > Add-ins > Go... > Browse... and pick' + #13#10 +
      'RedisExcel64-packed.xll (or RedisExcel-packed.xll for 32-bit Office).' + #13#10 + #13#10 +
      'Then try =RedisUDFGet("preco_btc") in a cell.';
end;

// ---- Uninstall: detect servers, let the user pick what to remove ----------
// Runs detect-servers.ps1 and fills the arrays; returns how many are present.
function DetectServers(): Integer;
var
  Script, Params, TmpFile, Line: String;
  Code, I, Count, Sep: Integer;
  Lines: TArrayOfString;
  KindS, LblS, Present: String;
begin
  SetArrayLength(ServerKinds, 0);
  SetArrayLength(ServerLabels, 0);
  SetArrayLength(ServerPresent, 0);

  TmpFile := ExpandConstant('{tmp}\redisexcel-servers.txt');
  Script := ExpandConstant('{app}\scripts\detect-servers.ps1');
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" > "' + TmpFile + '" 2>nul';
  if not Exec('powershell.exe', Params, ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code) then
  begin
    Result := 0;
    Exit;
  end;

  if not LoadStringsFromFile(TmpFile, Lines) then
  begin
    Result := 0;
    Exit;
  end;

  Count := 0;
  for I := 0 to GetArrayLength(Lines) - 1 do
  begin
    Line := Trim(Lines[I]);
    if Line <> '' then
    begin
      // kind|label|present
      Sep := Pos('|', Line);
      if Sep > 0 then
      begin
        KindS := Copy(Line, 1, Sep - 1);
        Line := Copy(Line, Sep + 1, Length(Line));
        Sep := Pos('|', Line);
        if Sep > 0 then
        begin
          LblS := Copy(Line, 1, Sep - 1);
          Present := Trim(Copy(Line, Sep + 1, Length(Line)));

          SetArrayLength(ServerKinds, Count + 1);
          SetArrayLength(ServerLabels, Count + 1);
          SetArrayLength(ServerPresent, Count + 1);
          ServerKinds[Count] := KindS;
          ServerLabels[Count] := LblS;
          ServerPresent[Count] := (Present = '1');
          Count := Count + 1;
        end;
      end;
    end;
  end;
  Result := Count;
end;

function InitializeUninstall(): Boolean;
var
  I, Present: Integer;
begin
  Result := True;
  UninstallServersPage := nil;
  try
    DetectServers();
    Present := 0;
    for I := 0 to GetArrayLength(ServerPresent) - 1 do
      if ServerPresent[I] then Present := Present + 1;
    // Only build the page when at least one server was detected.
    if Present > 0 then
    begin
      UninstallServersPage := CreateInputOptionPage(1,
        'Servers', 'Remove Redis-compatible servers too?',
        'These servers were found on this machine. Tick the ones you want to ' +
        'remove with RedisExcel.' + #13#10 +
        'By default NOTHING is removed.',
        False, True);
      for I := 0 to GetArrayLength(ServerKinds) - 1 do
        UninstallServersPage.Add(ServerLabels[I] + '  [found]');
    end;
  except
    // Never block an uninstall on the optional server page.
    UninstallServersPage := nil;
  end;
end;

procedure RunUninstallServers();
var
  I: Integer;
begin
  if UninstallServersPage = nil then Exit;
  for I := 0 to GetArrayLength(ServerKinds) - 1 do
    if UninstallServersPage.Values[I] then
      RunScript('uninstall-server.ps1', '-Kind ' + ServerKinds[I]);
end;

// ---- Uninstall: clean Excel + ProgId, ask about the config ----------------
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterExcelAutoLoad();
    // The add-in registers this ProgId under HKCU on load; remove it.
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\RedisRtd');
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    // Remove any servers the user ticked (the scripts still exist at this step
    // because the files are deleted after usPostUninstall's default handling).
    RunUninstallServers();

    if FileExists(ExpandConstant('{app}\RedisExcel.json')) then
      if MsgBox('Keep your RedisExcel.json settings?', mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
  end;
end;

