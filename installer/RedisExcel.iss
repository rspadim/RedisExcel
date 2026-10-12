; RedisExcel installer (Inno Setup 6).
;
; Build:  ISCC /DAppVersion=1.5.1 .\installer\RedisExcel.iss
; CI builds it on every run and attaches RedisExcelSetup-<version>.exe to a
; v* tag release (see .github/workflows/build.yml).
;
; What it does:
;   * copies RedisExcel.dll + the runtime XLL (+ NLog.config) into a folder the
;     user picks (default C:\RedisExcel);
;   * detects the Office bitness and keeps only the matching packed XLL;
;   * a short wizard asks for host / port / optional password and writes
;     RedisExcel.json next to the XLL (the add-in folder is searched first);
;   * OPTIONALLY brings up a Redis-compatible server (Docker or WSL, or points
;     at Memurai) via installer\install-server.ps1;
;   * at the end, checks the connection and (optionally) opens a sample Excel
;     workbook showing the status via installer\verify-connection.ps1;
;   * silent mode:  RedisExcelSetup.exe /VERYSILENT /DIR=... /HOST=... /PORT=...
;     /PASSWORD=... /SERVER=docker-valkey /OPENEXCEL=1

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

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "addinconfig"; Description: "Create RedisExcel.json in the install folder (host settings)"; GroupDescription: "Configuration:"
Name: "nlogconfig";  Description: "Create NLog.config (write a log file next to the add-in)"; GroupDescription: "Configuration:"

[Files]
Source: "{#BuildDir}\RedisExcel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\NLog.config"; DestDir: "{app}"; Flags: onlyifdoesntexist; Tasks: nlogconfig
; Helper scripts (server setup + end-of-install verification).
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\verify-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion

[Code]
var
  Is64BitOffice: Boolean;
  ConnPage: TInputQueryWizardPage;
  ServerPage: TInputOptionWizardPage;
  OptionsPage: TInputOptionWizardPage;
  InstallDone: Boolean;

// ---- Office bitness -------------------------------------------------------
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

function ParamOr(Param, Default: String): String;
begin
  Result := ExpandConstant('{param:' + Param + '|' + Default + '}');
end;

// ---- Wizard pages ---------------------------------------------------------
procedure InitializeWizard();
begin
  Is64BitOffice := Detect64BitOffice();

  ServerPage := CreateInputOptionPage(wpSelectTasks,
    'Redis server', 'Do you want to install a Redis-compatible server now?',
    'RedisExcel needs a Redis (or Valkey) server on localhost:6379. Choose an ' +
    'option if you do not already have one; Docker is the simplest.',
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
  ConnPage.Add('Password (optional):', True);
  ConnPage.Values[0] := ParamOr('HOST', 'localhost');
  ConnPage.Values[1] := ParamOr('PORT', '6379');
  ConnPage.Values[2] := ParamOr('PASSWORD', '');

  OptionsPage := CreateInputOptionPage(ConnPage.ID,
    'Finish', 'Extra checks',
    'Optional: open an Excel workbook at the end showing the connection status.',
    False, False);
  OptionsPage.Add('Open a sample RedisExcel workbook when finished');
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
function HostString(): String;
var
  H, Port, Pw: String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
  begin
    H := Trim(ConnPage.Values[0]);
    Port := Trim(ConnPage.Values[1]);
    Pw := Trim(ConnPage.Values[2]);
  end
  else
  begin
    H := ParamOr('HOST', 'localhost');
    Port := ParamOr('PORT', '6379');
    Pw := ParamOr('PASSWORD', '');
  end;
  if (Pos(':', H) = 0) and (Port <> '') then Result := H + ':' + Port else Result := H;
  if Pw <> '' then Result := Result + ',password=' + Pw;
end;

function ServerKind(): String;
var idx: Integer;
begin
  if WizardSilent() then
  begin
    Result := ParamOr('SERVER', 'none');
    Exit;
  end;
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

// ---- Run PowerShell helper scripts ---------------------------------------
function RunScript(ScriptName, Args: String): Integer;
var
  Script, Params: String;
  Code: Integer;
begin
  Script := ExpandConstant('{app}\scripts\') + ScriptName;
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" ' + Args;
  // Visible window so WSL/Memurai prompts and any MessageBox are reachable.
  if Exec('powershell.exe', Params, ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    Result := Code
  else
    Result := -1;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath, Lines, Kind: String;
begin
  if CurStep = ssPostInstall then
  begin
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

    // Optional: bring up a server before verifying.
    Kind := ServerKind();
    if Kind <> 'none' then
      RunScript('install-server.ps1',
        '-Kind ' + Kind + ' -Port ' + ParamOr('PORT', '6379'));

    InstallDone := True;
  end;
end;

// ---- Finish: verify + optionally open a sample workbook -------------------
procedure DeinitializeSetup();
var
  HostPort, Args: String;
begin
  if not InstallDone then Exit;
  // "host:port" (strip any credential) for the verifier.
  HostPort := HostString();
  if Pos(',', HostPort) > 0 then
    HostPort := Copy(HostPort, 1, Pos(',', HostPort) - 1);
  Args := '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '"';
  if OpenExcelChecked() then
    Args := Args + ' -OpenExcel'
  else
    Args := Args + ' -Quiet';
  RunScript('verify-connection.ps1', Args);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpFinished) and (not WizardSilent()) then
    WizardForm.FinishedLabel.Caption :=
      'RedisExcel is installed in:' + #13#10 + ExpandConstant('{app}') + #13#10 + #13#10 +
      'To use it: open Excel, go to File > Options > Add-ins > Go... > Browse... ' +
      'and select RedisExcel64-packed.xll (or RedisExcel-packed.xll for 32-bit Office).' + #13#10 + #13#10 +
      'Then try =RedisUDFGet("preco_btc") in a cell.';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    if FileExists(ExpandConstant('{app}\RedisExcel.json')) then
      if MsgBox('Keep your RedisExcel.json settings?', mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
end;
