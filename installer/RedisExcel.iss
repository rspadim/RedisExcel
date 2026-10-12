; RedisExcel installer (Inno Setup 6).
;
; Build:  ISCC .\installer\RedisExcel.iss
; Or:     msbuild /t:Build; the CI builds it on a v* tag (see .github/workflows/build.yml).
;
; What it does:
;   * copies RedisExcel.dll + the runtime XLL (+ NLog.config) into a folder the
;     user picks (default C:\RedisExcel);
;   * detects the Office bitness and installs the matching packed XLL
;     (RedisExcel64-packed.xll for 64-bit Office, RedisExcel-packed.xll for 32-bit);
;   * a small wizard asks for host / port / optional password and writes
;     RedisExcel.json next to the XLL (the add-in folder is searched first);
;   * offers to open the Excel add-ins dialog at the end;
;   * supports a silent "files only" mode:  RedisExcelSetup.exe /VERYSILENT
;     (the wizard and the Excel prompt are skipped; the config uses the given
;     or default host, see GetHost/GetPort below).

#define AppName "RedisExcel"
#define AppPublisher "rspadim"
#define AppUrl "https://github.com/rspadim/RedisExcel"
; Version is injected by the build (ISCC /DAppVersion=x.y.z); falls back to dev.
#ifndef AppVersion
  #define AppVersion "dev"
#endif
; Source folder with the built outputs (bin\Release\net48\publish) and NLog.config.
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
ArchitecturesInstallIn64BitMode=
UninstallDisplayName={#AppName} {#AppVersion}
; The uninstaller keeps RedisExcel.json by default (user settings).

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "addinconfig"; Description: "Create RedisExcel.json in the install folder (host settings)"; GroupDescription: "Configuration:"
Name: "nlogconfig";  Description: "Create NLog.config (write a log file next to the add-in)"; GroupDescription: "Configuration:"

[Files]
; The managed library the XLL loads.
Source: "{#BuildDir}\RedisExcel.dll"; DestDir: "{app}"; Flags: ignoreversion
; Both packed XLLs are shipped; [Code] keeps/removes one after the bitness check.
Source: "{#BuildDir}\RedisExcel64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
; Default log config; a user-customised one is preserved.
Source: "{#RepoRoot}\NLog.config"; DestDir: "{app}"; Flags: onlyifdoesntexist; Tasks: nlogconfig

[Code]
var
  Is64BitOffice: Boolean;
  WizardPage: TInputQueryWizardPage;

// ---- Office bitness -------------------------------------------------------
// Office writes its bitness under HKCU\Software\Microsoft\Office\<ver>\Outlook\Setup
// or ...\Common\General ("x86"/"x64"); ClickToRun exposes Platform. If nothing is
// found we default to 64-bit (the common case) but let the user override on the
// wizard page.
function Detect64BitOffice(): Boolean;
var
  Names: TArrayOfString;
  I: Integer;
  V, P: String;
begin
  Result := True;
  // ClickToRun configuration
  P := 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration';
  if RegQueryStringValue(HKLM, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else if RegQueryStringValue(HKCU, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else
  begin
    // Per-version "Bitness" under HKCU\Software\Microsoft\Office\<ver>\Common\General
    if RegGetSubkeyNames(HKCU, 'SOFTWARE\Microsoft\Office', Names) then
    begin
      for I := 0 to GetArrayLength(Names) - 1 do
      begin
        if RegQueryStringValue(HKCU,
             'SOFTWARE\Microsoft\Office\' + Names[I] + '\Common\General', 'Bitness', V) then
        begin
          Result := (CompareText(V, 'x64') = 0);
          Exit;
        end;
      end;
    end;
  end;
end;

function GetHost(Param: String): String;
begin
  Result := ExpandConstant('{param:HOST|localhost}');
end;

function GetPort(Param: String): String;
begin
  Result := ExpandConstant('{param:PORT|6379}');
end;

function GetPassword(Param: String): String;
begin
  Result := ExpandConstant('{param:PASSWORD|}');
end;

// ---- Wizard page (skipped in silent mode) ---------------------------------
procedure InitializeWizard();
begin
  Is64BitOffice := Detect64BitOffice();
  WizardPage := CreateInputQueryPage(wpSelectTasks,
    'Redis connection', 'Where should RedisExcel connect?',
    'Enter your Redis server. Press Next for the default (localhost:6379).' + #13#10 +
    'You can change this later by editing RedisExcel.json.');
  WizardPage.Add('Host (or host:port):', False);
  WizardPage.Add('Port:', False);
  WizardPage.Add('Password (optional):', True);
  WizardPage.Values[0] := 'localhost';
  WizardPage.Values[1] := '6379';
  WizardPage.Values[2] := '';
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := (PageID = WizardPage.ID) and WizardSilent();
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = WizardPage.ID) and (Trim(WizardPage.Values[0]) = '') then
  begin
    MsgBox('Please enter a Redis host (for example localhost).', mbError, MB_OK);
    Result := False;
  end;
end;

// ---- Output the config + pick the XLL -------------------------------------
function HostString(): String;
var
  H, Port, Pw: String;
begin
  if (not WizardSilent()) and (WizardPage <> nil) and (WizardPage.Values[0] <> '') then
  begin
    H := Trim(WizardPage.Values[0]);
    Port := Trim(WizardPage.Values[1]);
    Pw := Trim(WizardPage.Values[2]);
  end
  else
  begin
    H := GetHost('');
    Port := GetPort('');
    Pw := GetPassword('');
  end;
  // "host:port" when the user typed only a host.
  if (Pos(':', H) = 0) and (Port <> '') then
    Result := H + ':' + Port
  else
    Result := H;
  if Pw <> '' then
    Result := Result + ',password=' + Pw;
end;

function ConfigJson(): String;
var
  HostLine: String;
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

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath, Lines: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Keep only the XLL matching the Office bitness.
    if Is64BitOffice then
      DeleteFile(ExpandConstant('{app}\RedisExcel-packed.xll'))
    else
      DeleteFile(ExpandConstant('{app}\RedisExcel64-packed.xll'));

    // Write RedisExcel.json next to the XLL (searched first by the add-in).
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
  end;
end;

// ---- Finish page: offer to open Excel's add-ins dialog ---------------------
procedure OpenAddInsDialog();
var
  ResultCode: Integer;
begin
  // Best effort: open Excel (the user then browses to the XLL).
  ShellExec('open', 'excel.exe', '', ExpandConstant('{app}'), SW_SHOWNORMAL, ewNoWait, ResultCode);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpFinished) and (not WizardSilent()) then
    WizardForm.FinishedLabel.Caption :=
      'RedisExcel files are installed in:' + #13#10 + ExpandConstant('{app}') + #13#10 + #13#10 +
      'Last step (cannot be automated safely): open Excel, go to' + #13#10 +
      'File > Options > Add-ins > Go... > Browse... and select' + #13#10 +
      'RedisExcel64-packed.xll (or RedisExcel-packed.xll for 32-bit Office).' + #13#10 + #13#10 +
      'Put =RedisUDFGet("preco_btc") in a cell to test the connection.';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    if FileExists(ExpandConstant('{app}\RedisExcel.json')) then
      if MsgBox('Keep your RedisExcel.json settings?',
                mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
  end;
end;
