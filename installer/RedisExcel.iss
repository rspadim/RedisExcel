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
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

; All user-visible text. Order: english (default) / spanish / brazilianportuguese.
[CustomMessages]
english.ServerPageTitle=Redis server
english.ServerPageSubtitle=Do you want to install a Redis-compatible server now?
english.ServerPageDesc=RedisExcel needs a Redis (or Valkey) server. Choose an option if you do not already have one. The server is installed first, then RedisExcel, and the connection is tested at the end. Docker is the simplest.
english.ServerNone=No - I already have a server (or will set it up later)
english.ServerDockerRedis=Yes - Docker: Redis (redis:latest)
english.ServerDockerValkey=Yes - Docker: Valkey (valkey/valkey:latest)
english.ServerWslRedis=Yes - WSL2: Redis (apt install redis-server)
english.ServerWslValkey=Yes - WSL2: Valkey (apt install valkey-server)
english.ServerMemurai=Yes - Memurai (native Windows service, needs admin)
english.ConnPageTitle=Redis connection
english.ConnPageSubtitle=Where should RedisExcel connect?
english.ConnPageDesc=Press Next for the default (localhost:6379). You can change this later by editing RedisExcel.json.
english.ConnHost=Host (or host:port):
english.ConnPort=Port:
english.ConnUser=User (optional, for Redis 6 ACL):
english.ConnPassword=Password (optional):
english.ConnDatabase=Database (0-15, default 0):
english.OptionsPageTitle=Extra checks
english.OptionsPageSubtitle=Almost done
english.OptionsPageDesc=The installer tests the connection when it finishes. Optionally open a sample Excel workbook showing it working.
english.OptionsOpenExcel=Open a sample RedisExcel workbook at the end
english.OptionsUseSsl=Use SSL/TLS for the connection
english.TaskConfigDesc=Create RedisExcel.json in the install folder (host settings)
english.TaskNlogDesc=Create NLog.config (write a log file next to the add-in)
english.TaskAutoLoadDesc=Register the add-in so Excel loads it automatically (recommended)
english.GroupConfig=Configuration:
english.GroupExcel=Excel:
english.IconUninstall=Uninstall {#AppName}
english.HostRequired=Please enter a Redis host (for example localhost).
english.ServerFailed=Could not set up %1 automatically (exit %2).%nYou can install a server later and point RedisExcel.json at it.
english.ConfigWriteFailed=Could not write %1. Create it manually from the sample (RedisExcel.json).
english.TestFailed=Installed, but the Redis at %1 did not answer a PING.%nCheck that the server is running (docker ps) and adjust RedisExcel.json if it is on another host/port.
english.TestOk=All set! Redis answered at %1.%n%nTo use it: open Excel. If the add-in does not load automatically, go to File > Options > Add-ins > Go... > Browse... and pick the .xll in %2.
english.UpgradeConfigKept=An existing RedisExcel.json was found and kept (your settings were not overwritten).%n%nThis is a MAJOR update: the configuration format changed. If that file uses the old format, every cell will show a #CONFIG error until it is migrated.%n%nRun Validate-Config.cmd (or =RedisCheckConfigFile() in Excel) and see the README section "Migrating from the old shape".
english.FinishedText=RedisExcel is installed in:%n%2%n%nIf the add-in does not load automatically, enable it in Excel:%nFile > Options > Add-ins > Go... > Browse... and pick RedisExcel64-packed.xll (or RedisExcel-packed.xll for 32-bit Office).%n%nThen try =RedisUDFGet("preco_btc") in a cell.
english.UninstServersTitle=Servers
english.UninstServersSubtitle=Remove Redis-compatible servers too?
english.UninstServersDesc=These servers were found on this machine. Tick the ones you want to remove with RedisExcel.%nBy default NOTHING is removed.
english.UninstServerFound=%1  [found]
english.UninstKeepConfig=Keep your RedisExcel.json settings?
english.ServerNameDockerRedis=Redis (Docker)
english.ServerNameDockerValkey=Valkey (Docker)
english.ServerNameWslRedis=Redis (WSL)
english.ServerNameWslValkey=Valkey (WSL)
english.ServerNameMemurai=Memurai
english.ServerNameGeneric=the server
english.ModePageTitle=Setup mode
english.ModePageSubtitle=Simple or advanced setup?
english.ModePageDesc=Simple asks only for the server address. Advanced lets you set the database, SSL/TLS, the write mode and the real-time update tuning. You can always edit RedisExcel.json later (or run Validate-Config.cmd).
english.ModeSimple=Simple - just the Redis address (recommended)
english.ModeAdvanced=Advanced - set database, SSL, write mode and update tuning
english.AdvPageTitle=Advanced connection
english.AdvPageDesc=This is your whole RedisExcel.json, ready to edit: every key is pre-filled with its default and a short comment. The add-in accepts // comments, so you can leave them. Change hosts, aliases, write mode, ConflationMs, AsyncWrites - everything is here.
english.AdvPageSubtitle=Database, credentials and SSL
english.AdvDatabase=Database (0-15):
english.AdvUser=User (optional, Redis 6 ACL):
english.AdvPassword=Password (optional):
english.AdvOptPageTitle=Advanced options
english.AdvOptPageSubtitle=Write and update behavior
english.AdvOptPageDesc=Defaults are fine for most users. Tick SSL for a rediss server; the write-mode and update settings are explained in the README.
english.AdvUseSsl=Use SSL/TLS for the connection
english.AdvUpdateCheck=Check for a newer RedisExcel release (background)
english.SummaryPageTitle=Your settings
english.SummaryPageSubtitle=Review the connection that will be written to RedisExcel.json
english.SummaryIntro=These values will be saved to RedisExcel.json in the install folder.
english.SummaryLabel=Connection
english.SummaryMode=Mode
english.SummarySimple=Simple
english.SummaryAdvanced=Advanced

spanish.ServerPageTitle=Servidor Redis
spanish.ServerPageSubtitle=¿Desea instalar ahora un servidor compatible con Redis?
spanish.ServerPageDesc=RedisExcel necesita un servidor Redis (o Valkey). Elija una opción si aún no tiene uno. Primero se instala el servidor, luego RedisExcel, y la conexión se prueba al final. Docker es lo más sencillo.
spanish.ServerNone=No - ya tengo un servidor (o lo instalaré después)
spanish.ServerDockerRedis=Sí - Docker: Redis (redis:latest)
spanish.ServerDockerValkey=Sí - Docker: Valkey (valkey/valkey:latest)
spanish.ServerWslRedis=Sí - WSL2: Redis (apt install redis-server)
spanish.ServerWslValkey=Sí - WSL2: Valkey (apt install valkey-server)
spanish.ServerMemurai=Sí - Memurai (servicio de Windows nativo, requiere admin)
spanish.ConnPageTitle=Conexión a Redis
spanish.ConnPageSubtitle=¿Dónde debe conectarse RedisExcel?
spanish.ConnPageDesc=Pulse Siguiente para el valor predeterminado (localhost:6379). Puede cambiarlo luego editando RedisExcel.json.
spanish.ConnHost=Host (o host:puerto):
spanish.ConnPort=Puerto:
spanish.ConnUser=Usuario (opcional, para ACL de Redis 6):
spanish.ConnPassword=Contraseña (opcional):
spanish.ConnDatabase=Base de datos (0-15, predeterminado 0):
spanish.OptionsPageTitle=Comprobaciones
spanish.OptionsPageSubtitle=Casi listo
spanish.OptionsPageDesc=El instalador prueba la conexión al terminar. Opcionalmente abre un libro de Excel de ejemplo que muestra que funciona.
spanish.OptionsOpenExcel=Abrir un libro de RedisExcel de ejemplo al final
spanish.OptionsUseSsl=Usar SSL/TLS para la conexion
spanish.TaskConfigDesc=Crear RedisExcel.json en la carpeta de instalación (ajustes de host)
spanish.TaskNlogDesc=Crear NLog.config (escribir un archivo de registro junto al complemento)
spanish.TaskAutoLoadDesc=Registrar el complemento para que Excel lo cargue automáticamente (recomendado)
spanish.GroupConfig=Configuración:
spanish.GroupExcel=Excel:
spanish.IconUninstall=Desinstalar {#AppName}
spanish.HostRequired=Introduzca un host de Redis (por ejemplo localhost).
spanish.ServerFailed=No se pudo configurar %1 automáticamente (salida %2).%nPuede instalar un servidor después y apuntar RedisExcel.json a él.
spanish.ConfigWriteFailed=No se pudo escribir %1. Créelo manualmente desde el ejemplo (RedisExcel.json).
spanish.TestFailed=Instalado, pero el Redis en %1 no respondió a PING.%nCompruebe que el servidor esté en ejecución (docker ps) y ajuste RedisExcel.json si está en otro host/puerto.
spanish.TestOk=¡Todo listo! Redis respondió en %1.%n%nPara usarlo: abra Excel. Si el complemento no se carga solo, vaya a Archivo > Opciones > Complementos > Ir... > Examinar... y elija el .xll en %2.
spanish.UpgradeConfigKept=Se encontro un RedisExcel.json existente y se conservo (sus ajustes no se sobrescribieron).%n%nEsta es una actualizacion MAYOR: el formato de configuracion cambio. Si ese archivo usa el formato antiguo, cada celda mostrara un error #CONFIG hasta migrarlo.%n%nEjecute Validate-Config.cmd (o =RedisCheckConfigFile() en Excel) y vea la seccion del README "Migrating from the old shape".
spanish.FinishedText=RedisExcel está instalado en:%n%2%n%nSi el complemento no se carga solo, actívelo en Excel:%nArchivo > Opciones > Complementos > Ir... > Examinar... y elija RedisExcel64-packed.xll (o RedisExcel-packed.xll para Office de 32 bits).%n%nLuego pruebe =RedisUDFGet("preco_btc") en una celda.
spanish.UninstServersTitle=Servidores
spanish.UninstServersSubtitle=¿Eliminar también los servidores compatibles con Redis?
spanish.UninstServersDesc=Se encontraron estos servidores en esta máquina. Marque los que desee eliminar junto con RedisExcel.%nDe forma predeterminada NO se elimina nada.
spanish.UninstServerFound=%1  [encontrado]
spanish.UninstKeepConfig=¿Conservar su configuración RedisExcel.json?
spanish.ServerNameDockerRedis=Redis (Docker)
spanish.ServerNameDockerValkey=Valkey (Docker)
spanish.ServerNameWslRedis=Redis (WSL)
spanish.ServerNameWslValkey=Valkey (WSL)
spanish.ServerNameMemurai=Memurai
spanish.ServerNameGeneric=el servidor
spanish.ModePageTitle=Modo de instalacion
spanish.ModePageSubtitle=Configuracion simple o avanzada?
spanish.ModePageDesc=Simple solo pide la direccion del servidor. Avanzada permite definir la base de datos, SSL/TLS, el modo de escritura y los ajustes de tiempo real. Siempre puede editar RedisExcel.json despues (o ejecutar Validate-Config.cmd).
spanish.ModeSimple=Simple - solo la direccion de Redis (recomendado)
spanish.ModeAdvanced=Avanzada - base de datos, SSL, modo de escritura y ajustes
spanish.AdvPageTitle=Conexion avanzada
spanish.AdvPageDesc=Este es todo su RedisExcel.json, listo para editar: cada clave viene con su valor predeterminado y un comentario corto. El complemento acepta comentarios //, puede dejarlos. Cambie hosts, alias, modo de escritura, ConflationMs, AsyncWrites: todo esta aqui.
spanish.AdvPageSubtitle=Base de datos, credenciales y SSL
spanish.AdvDatabase=Base de datos (0-15):
spanish.AdvUser=Usuario (opcional, ACL de Redis 6):
spanish.AdvPassword=Contrasena (opcional):
spanish.AdvOptPageTitle=Opciones avanzadas
spanish.AdvOptPageSubtitle=Comportamiento de escritura y actualizacion
spanish.AdvOptPageDesc=Los valores predeterminados sirven para casi todos. Marque SSL si su servidor es rediss; el modo de escritura y los ajustes de actualizacion se explican en el README.
spanish.AdvUseSsl=Usar SSL/TLS para la conexion
spanish.AdvUpdateCheck=Buscar una version mas reciente de RedisExcel (en segundo plano)
spanish.SummaryPageTitle=Sus ajustes
spanish.SummaryPageSubtitle=Revise la conexion que se guardara en RedisExcel.json
spanish.SummaryIntro=Estos valores se guardaran en RedisExcel.json en la carpeta de instalacion.
spanish.SummaryLabel=Conexion
spanish.SummaryMode=Modo
spanish.SummarySimple=Simple
spanish.SummaryAdvanced=Avanzada

brazilianportuguese.ServerPageTitle=Servidor Redis
brazilianportuguese.ServerPageSubtitle=Deseja instalar agora um servidor compatível com o Redis?
brazilianportuguese.ServerPageDesc=O RedisExcel precisa de um servidor Redis (ou Valkey). Escolha uma opção se ainda não tiver um. O servidor é instalado primeiro, depois o RedisExcel, e a conexão é testada no final. O Docker é o mais simples.
brazilianportuguese.ServerNone=Não - eu já tenho um servidor (ou vou instalar depois)
brazilianportuguese.ServerDockerRedis=Sim - Docker: Redis (redis:latest)
brazilianportuguese.ServerDockerValkey=Sim - Docker: Valkey (valkey/valkey:latest)
brazilianportuguese.ServerWslRedis=Sim - WSL2: Redis (apt install redis-server)
brazilianportuguese.ServerWslValkey=Sim - WSL2: Valkey (apt install valkey-server)
brazilianportuguese.ServerMemurai=Sim - Memurai (serviço nativo do Windows, precisa de admin)
brazilianportuguese.ConnPageTitle=Conexão com o Redis
brazilianportuguese.ConnPageSubtitle=Onde o RedisExcel deve se conectar?
brazilianportuguese.ConnPageDesc=Clique em Avançar para o padrão (localhost:6379). Você pode mudar depois editando o RedisExcel.json.
brazilianportuguese.ConnHost=Host (ou host:porta):
brazilianportuguese.ConnPort=Porta:
brazilianportuguese.ConnUser=Usuário (opcional, para ACL do Redis 6):
brazilianportuguese.ConnPassword=Senha (opcional):
brazilianportuguese.ConnDatabase=Banco (0-15, padrao 0):
brazilianportuguese.OptionsPageTitle=Verificações
brazilianportuguese.OptionsPageSubtitle=Quase pronto
brazilianportuguese.OptionsPageDesc=O instalador testa a conexão ao terminar. Opcionalmente abre uma planilha de exemplo do RedisExcel mostrando que funcionou.
brazilianportuguese.OptionsOpenExcel=Abrir uma planilha de exemplo do RedisExcel no final
brazilianportuguese.OptionsUseSsl=Usar SSL/TLS na conexao
brazilianportuguese.TaskConfigDesc=Criar RedisExcel.json na pasta de instalação (configuração de host)
brazilianportuguese.TaskNlogDesc=Criar NLog.config (gravar um arquivo de log junto ao add-in)
brazilianportuguese.TaskAutoLoadDesc=Registrar o add-in para o Excel carregá-lo automaticamente (recomendado)
brazilianportuguese.GroupConfig=Configuração:
brazilianportuguese.GroupExcel=Excel:
brazilianportuguese.IconUninstall=Desinstalar {#AppName}
brazilianportuguese.HostRequired=Informe um host do Redis (por exemplo localhost).
brazilianportuguese.ServerFailed=Não foi possível configurar %1 automaticamente (saída %2).%nVocê pode instalar um servidor depois e apontar o RedisExcel.json para ele.
brazilianportuguese.ConfigWriteFailed=Não foi possível gravar %1. Crie manualmente a partir do exemplo (RedisExcel.json).
brazilianportuguese.TestFailed=Instalado, mas o Redis em %1 não respondeu ao PING.%nConfira se o servidor está rodando (docker ps) e ajuste o RedisExcel.json se ele estiver em outro host/porta.
brazilianportuguese.TestOk=Tudo certo! O Redis respondeu em %1.%n%nPara usar: abra o Excel. Se o add-in não carregar sozinho, vá em Arquivo > Opções > Suplementos > Ir... > Procurar... e escolha o .xll em %2.
brazilianportuguese.UpgradeConfigKept=Um RedisExcel.json existente foi encontrado e mantido (suas configuracoes nao foram sobrescritas).%n%nEsta e uma atualizacao MAIOR: o formato de configuracao mudou. Se esse arquivo usa o formato antigo, cada celula mostrara um erro #CONFIG ate ser migrado.%n%nRode o Validate-Config.cmd (ou =RedisCheckConfigFile() no Excel) e veja a secao "Migrating from the old shape" no README.
brazilianportuguese.FinishedText=O RedisExcel foi instalado em:%n%2%n%nSe o add-in não carregar sozinho, habilite no Excel:%nArquivo > Opções > Suplementos > Ir... > Procurar... e escolha RedisExcel64-packed.xll (ou RedisExcel-packed.xll para Office 32 bits).%n%nDepois teste =RedisUDFGet("preco_btc") em uma célula.
brazilianportuguese.UninstServersTitle=Servidores
brazilianportuguese.UninstServersSubtitle=Remover também os servidores compatíveis com o Redis?
brazilianportuguese.UninstServersDesc=Estes servidores foram encontrados na máquina. Marque os que deseja remover junto com o RedisExcel.%nPor padrão NADA é removido.
brazilianportuguese.UninstServerFound=%1  [encontrado]
brazilianportuguese.UninstKeepConfig=Manter suas configurações do RedisExcel.json?
brazilianportuguese.ServerNameDockerRedis=Redis (Docker)
brazilianportuguese.ServerNameDockerValkey=Valkey (Docker)
brazilianportuguese.ServerNameWslRedis=Redis (WSL)
brazilianportuguese.ServerNameWslValkey=Valkey (WSL)
brazilianportuguese.ServerNameMemurai=Memurai
brazilianportuguese.ServerNameGeneric=o servidor
brazilianportuguese.ModePageTitle=Modo de instalacao
brazilianportuguese.ModePageSubtitle=Instalacao simples ou avancada?
brazilianportuguese.ModePageDesc=O Simples pede apenas o endereco do servidor. O Avancado permite definir o banco, SSL/TLS, o modo de escrita e os ajustes de tempo real. Voce pode editar o RedisExcel.json depois (ou rodar o Validate-Config.cmd).
brazilianportuguese.ModeSimple=Simples - apenas o endereco do Redis (recomendado)
brazilianportuguese.ModeAdvanced=Avancado - banco, SSL, modo de escrita e ajustes
brazilianportuguese.AdvPageTitle=Conexao avancada
brazilianportuguese.AdvPageDesc=Este e o seu RedisExcel.json completo, pronto para editar: cada chave ja vem com o valor padrao e um comentario curto. O add-in aceita comentarios //, pode deixa-los. Ajuste hosts, aliases, modo de escrita, ConflationMs, AsyncWrites: tudo esta aqui.
brazilianportuguese.AdvPageSubtitle=Banco, credenciais e SSL
brazilianportuguese.AdvDatabase=Banco (0-15):
brazilianportuguese.AdvUser=Usuario (opcional, ACL do Redis 6):
brazilianportuguese.AdvPassword=Senha (opcional):
brazilianportuguese.AdvOptPageTitle=Opcoes avancadas
brazilianportuguese.AdvOptPageSubtitle=Comportamento de escrita e atualizacao
brazilianportuguese.AdvOptPageDesc=Os padroes servem para a maioria. Marque SSL se o servidor for rediss; o modo de escrita e os ajustes de atualizacao estao explicados no README.
brazilianportuguese.AdvUseSsl=Usar SSL/TLS na conexao
brazilianportuguese.AdvUpdateCheck=Verificar se ha uma versao mais nova do RedisExcel (em segundo plano)
brazilianportuguese.SummaryPageTitle=Suas configuracoes
brazilianportuguese.SummaryPageSubtitle=Confira a conexao que sera gravada no RedisExcel.json
brazilianportuguese.SummaryIntro=Estes valores serao salvos no RedisExcel.json na pasta de instalacao.
brazilianportuguese.SummaryLabel=Conexao
brazilianportuguese.SummaryMode=Modo
brazilianportuguese.SummarySimple=Simples
brazilianportuguese.SummaryAdvanced=Avancado

[Tasks]
Name: "addinconfig"; Description: "{cm:TaskConfigDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "nlogconfig";  Description: "{cm:TaskNlogDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "excelautoload"; Description: "{cm:TaskAutoLoadDesc}"; GroupDescription: "{cm:GroupExcel}"

[Icons]
Name: "{group}\{cm:IconUninstall}"; Filename: "{uninstallexe}"

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
  ModePage: TInputOptionWizardPage;
  AdvConfigPage: TWizardPage;
  AdvMemo: TNewMemo;
  OptionsPage: TInputOptionWizardPage;
  SummaryPage: TOutputMsgMemoWizardPage;
  AdvSeeded: Boolean;
  ServerInstallCode: Integer;
  TestOk: Boolean;
  InstallCompleted: Boolean;
  InstallKeptConfig: Boolean;
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

function BoolToStr(b: Boolean): String;
begin
  if b then Result := 'true' else Result := 'false';
end;

procedure InitializeWizard();
begin
  Is64BitOffice := Detect64BitOffice();

  ModePage := CreateInputOptionPage(wpSelectTasks,
    ExpandConstant('{cm:ModePageTitle}'),
    ExpandConstant('{cm:ModePageSubtitle}'),
    ExpandConstant('{cm:ModePageDesc}'),
    True, False);
  ModePage.Add(ExpandConstant('{cm:ModeSimple}'));
  ModePage.Add(ExpandConstant('{cm:ModeAdvanced}'));
  ModePage.SelectedValueIndex := 0;

  ServerPage := CreateInputOptionPage(ModePage.ID,
    ExpandConstant('{cm:ServerPageTitle}'),
    ExpandConstant('{cm:ServerPageSubtitle}'),
    ExpandConstant('{cm:ServerPageDesc}'),
    True, False);
  ServerPage.Add(ExpandConstant('{cm:ServerNone}'));
  ServerPage.Add(ExpandConstant('{cm:ServerDockerRedis}'));
  ServerPage.Add(ExpandConstant('{cm:ServerDockerValkey}'));
  ServerPage.Add(ExpandConstant('{cm:ServerWslRedis}'));
  ServerPage.Add(ExpandConstant('{cm:ServerWslValkey}'));
  ServerPage.Add(ExpandConstant('{cm:ServerMemurai}'));
  ServerPage.SelectedValueIndex := 0;

  // Basic connection page (both modes): host, port, password.
  ConnPage := CreateInputQueryPage(ServerPage.ID,
    ExpandConstant('{cm:ConnPageTitle}'),
    ExpandConstant('{cm:ConnPageSubtitle}'),
    ExpandConstant('{cm:ConnPageDesc}'));
  ConnPage.Add(ExpandConstant('{cm:ConnHost}'), False);
  ConnPage.Add(ExpandConstant('{cm:ConnPort}'), False);
  ConnPage.Add(ExpandConstant('{cm:ConnPassword}'), True);
  ConnPage.Values[0] := ParamOr('HOST', 'localhost');
  ConnPage.Values[1] := ParamOr('PORT', '6379');
  ConnPage.Values[2] := ParamOr('PASSWORD', '');

  // Advanced mode: a single editable page with the whole RedisExcel.json
  // (pre-filled with comments explaining every key). One screen, everything.
  AdvConfigPage := CreateCustomPage(ConnPage.ID,
    ExpandConstant('{cm:AdvPageTitle}'),
    ExpandConstant('{cm:AdvPageDesc}'));
  AdvMemo := TNewMemo.Create(AdvConfigPage);
  AdvMemo.Parent := AdvConfigPage.Surface;
  AdvMemo.Left := 0;
  AdvMemo.Top := 0;
  AdvMemo.Width := AdvConfigPage.SurfaceWidth;
  AdvMemo.Height := AdvConfigPage.SurfaceHeight;
  AdvMemo.ScrollBars := ssVertical;
  AdvMemo.WordWrap := False;
  AdvMemo.Font.Name := 'Consolas';
  AdvMemo.Font.Size := 9;

  OptionsPage := CreateInputOptionPage(AdvConfigPage.ID,
    ExpandConstant('{cm:OptionsPageTitle}'),
    ExpandConstant('{cm:OptionsPageSubtitle}'),
    ExpandConstant('{cm:OptionsPageDesc}'),
    False, False);
  OptionsPage.Add(ExpandConstant('{cm:OptionsOpenExcel}'));
  OptionsPage.Values[0] := (ParamOr('OPENEXCEL', '0') = '1');

  // Read-only summary of what will be written to RedisExcel.json.
  SummaryPage := CreateOutputMsgMemoPage(OptionsPage.ID,
    ExpandConstant('{cm:SummaryPageTitle}'),
    ExpandConstant('{cm:SummaryPageSubtitle}'),
    ExpandConstant('{cm:SummaryIntro}'), '');
end;

function IsAdvanced(): Boolean;
begin
  if WizardSilent() then
    Result := (Lowercase(ParamOr('MODE', 'simple')) = 'advanced')
  else
    Result := (ModePage <> nil) and (ModePage.SelectedValueIndex = 1);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  if WizardSilent() then
    Result := (PageID = ModePage.ID) or (PageID = ServerPage.ID) or (PageID = ConnPage.ID) or
              (PageID = AdvConfigPage.ID) or (PageID = OptionsPage.ID)
  else
    // In simple mode the advanced config page is hidden.
    Result := (not IsAdvanced()) and (PageID = AdvConfigPage.ID);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = ConnPage.ID) and (Trim(ConnPage.Values[0]) = '') then
  begin
    MsgBox(ExpandConstant('{cm:HostRequired}'), mbError, MB_OK);
    Result := False;
  end;
end;

// ---- Values ---------------------------------------------------------------
// Reads a field from the page that is visible in the current mode.
function Host(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[0])
  else
    Result := ParamOr('HOST', 'localhost');
end;

function PortValue(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[1])
  else
    Result := ParamOr('PORT', '6379');
end;

function PasswordValue(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[2])
  else
    Result := ParamOr('PASSWORD', '');
end;

// ---- Simple-mode host string (advanced mode edits the JSON directly) ------
function HostString(): String;
var
  H, Port, Pw: String;
begin
  H := Host();
  Port := PortValue();
  Pw := PasswordValue();
  if (Pos(':', H) = 0) and (Port <> '') then Result := H + ':' + Port else Result := H;
  Result := Result + ',defaultDatabase=0,ssl=False,abortConnect=False';
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

// The generated config for SIMPLE mode: minimal v1.6 shape (server declared
// once via Default + Servers). Advanced mode uses the editable text.
function SimpleConfigJson(): String;
var HostLine: String;
begin
  HostLine := HostString();
  Result :=
    '{' + #13#10 +
    '  "Default": "local",' + #13#10 +
    '  "Servers": {' + #13#10 +
    '    "local": "' + HostLine + '"' + #13#10 +
    '  },' + #13#10 +
    '  "RTD": {' + #13#10 +
    '    "RedisUpdateRateMs": 1000,' + #13#10 +
    '    "ExcelUpdateRateMs": 100,' + #13#10 +
    '    "MessageCounterThreshold": 10000,' + #13#10 +
    '    "ExcelUpdateStyle": "Automatic",' + #13#10 +
    '    "UseGetMultiple": true,' + #13#10 +
    '    "CoalesceRealtimeUpdates": true' + #13#10 +
    '  },' + #13#10 +
    '  "UpdateCheck": true,' + #13#10 +
    '  "SkipRepeatedMessages": true,' + #13#10 +
    '  "ConflationMs": 200,' + #13#10 +
    '  "PublishDedupCacheSize": 10000,' + #13#10 +
    '  "SyncWrite": "fireforget",' + #13#10 +
    '  "AsyncWrites": false' + #13#10 +
    '}' + #13#10;
end;

// The fully commented template the advanced page starts from. Exposes every key
// the add-in understands, with a one-line explanation each. The loader accepts
// // and /* */ comments, so this stays valid as-is.
function AdvancedTemplate(): String;
var
  H, Port, Pw: String;
begin
  H := Host();
  Port := PortValue();
  Pw := PasswordValue();
  if (Pos(':', H) = 0) and (Port <> '') then H := H + ':' + Port;
  H := H + ',defaultDatabase=0,ssl=False,abortConnect=False';
  if Pw <> '' then H := H + ',password=' + Pw;
  Result :=
    '{' + #13#10 +
    '  // RedisExcel.json - comments (// and /* */) are allowed.' + #13#10 +
    '  // The server is declared once: Default is a name from Servers (or a raw' + #13#10 +
    '  // connection string). Every formula may still pass its own host argument.' + #13#10 +
    '  "Default": "local",' + #13#10 +
    '  // Global timeout defaults (ms) for servers declared as a plain string.' + #13#10 +
    '  "Timeout": { "connectTimeout": 5000, "syncTimeout": 5000 },' + #13#10 +
    '  // Named servers: value is a connection string OR { host, connectTimeout, syncTimeout }.' + #13#10 +
    '  "Servers": {' + #13#10 +
    '    "local": "' + H + '"' + #13#10 +
    '    // "prod": { "host": "sentinel-a:26379,sentinel-b:26379,serviceName=mymaster", "connectTimeout": 8000, "syncTimeout": 8000 },' + #13#10 +
    '    // "cache": "node-a:6379,node-b:6379"' + #13#10 +
    '  },' + #13#10 +
    '  "RTD": {' + #13#10 +
    '    "RedisUpdateRateMs": 1000,              // how often the poll timer reads Redis' + #13#10 +
    '    "ExcelUpdateRateMs": 100,               // how often dirty values are flushed to Excel' + #13#10 +
    '    "ExcelUpdateStyle": "Automatic",        // Timer | Realtime | Automatic' + #13#10 +
    '    "MessageCounterThreshold": 10000,       // Automatic: msgs/sec above which it switches to timer; <=0 disables' + #13#10 +
    '    "UseGetMultiple": true,                 // batch GET topics into one MGET per tick' + #13#10 +
    '    "CoalesceRealtimeUpdates": true         // real-time: at most one push per window (superseded by ConflationMs)' + #13#10 +
    '  },' + #13#10 +
    '  "UpdateCheck": true,                      // background check for a newer release' + #13#10 +
    '  "SkipRepeatedMessages": true,             // skip identical consecutive SUB/GET payloads (ordinal)' + #13#10 +
    '  "ConflationMs": 200,                      // explicit real-time window in ms (0 = off; capped at 3600000)' + #13#10 +
    '  "PublishDedupCacheSize": 10000,           // LRU size for PublishIfChanged' + #13#10 +
    '  "SyncWrite": "fireforget",                // sync | fireforget | fireforget-all' + #13#10 +
    '  "AsyncWrites": false                      // dispatch writes off the Excel thread (Observe)' + #13#10 +
    '}' + #13#10;
end;

function ConfigJson(): String;
begin
  if IsAdvanced() then
  begin
    if (not WizardSilent()) and (AdvMemo <> nil) then
      Result := AdvMemo.Lines.Text
    else
      Result := AdvancedTemplate();
  end
  else
    Result := SimpleConfigJson();
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
  if Kind = 'docker-redis' then Result := ExpandConstant('{cm:ServerNameDockerRedis}')
  else if Kind = 'docker-valkey' then Result := ExpandConstant('{cm:ServerNameDockerValkey}')
  else if Kind = 'wsl-redis' then Result := ExpandConstant('{cm:ServerNameWslRedis}')
  else if Kind = 'wsl-valkey' then Result := ExpandConstant('{cm:ServerNameWslValkey}')
  else if Kind = 'memurai' then Result := ExpandConstant('{cm:ServerNameMemurai}')
  else Result := ExpandConstant('{cm:ServerNameGeneric}');
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
            ' -Port ' + PortValue();
  if Exec('powershell.exe', Params, ExpandConstant('{tmp}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    ServerInstallCode := Code
  else
    ServerInstallCode := -1;

  if (ServerInstallCode <> 0) and (not WizardSilent()) then
    MsgBox(FmtMessage(ExpandConstant('{cm:ServerFailed}'), [ServerName(Kind), IntToStr(ServerInstallCode)]), mbInformation, MB_OK);
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
          MsgBox(FmtMessage(ExpandConstant('{cm:ConfigWriteFailed}'), [ConfigPath]),
                 mbInformation, MB_OK);
      end
      else
        // A pre-1.6 config is preserved verbatim: remember it so the finish step
        // can warn about the migration (the add-in rejects the old shape).
        InstallKeptConfig := True;
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
    Msg := FmtMessage(ExpandConstant('{cm:TestFailed}'), [HostPort])
  else if InstallKeptConfig then
    // An upgrade kept the user's existing RedisExcel.json: warn about the v1.6
    // break (the old shape is rejected -> #CONFIG) instead of claiming all-set.
    Msg := FmtMessage(ExpandConstant('{cm:UpgradeConfigKept}'), [HostPort, ExpandConstant('{app}')])
  else
    Msg := FmtMessage(ExpandConstant('{cm:TestOk}'), [HostPort, ExpandConstant('{app}')]);

  if not WizardSilent() then
    MsgBox(Msg, mbInformation, MB_OK);

  if OpenExcelChecked() then
    RunScript('verify-connection.ps1',
      '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '" -OpenExcel -Quiet');
end;

procedure CurPageChanged(CurPageID: Integer);
var
  Lines, ModeText: String;
begin
  // Seed the advanced editor once, from the template (only the first visit).
  if (CurPageID = AdvConfigPage.ID) and (not WizardSilent()) and (not AdvSeeded) then
  begin
    AdvMemo.Lines.Text := AdvancedTemplate();
    AdvSeeded := True;
  end;

  if (CurPageID = SummaryPage.ID) and (not WizardSilent()) then
  begin
    if IsAdvanced() then ModeText := ExpandConstant('{cm:SummaryAdvanced}')
    else ModeText := ExpandConstant('{cm:SummarySimple}');
    Lines :=
      ExpandConstant('{cm:SummaryLabel}') + ':  ' + HostString() + #13#10 +
      ExpandConstant('{cm:SummaryMode}') + ':  ' + ModeText;
    SummaryPage.RichEditViewer.Lines.Text := Lines;
  end;

  if (CurPageID = wpFinished) and (not WizardSilent()) then
    WizardForm.FinishedLabel.Caption :=
      FmtMessage(ExpandConstant('{cm:FinishedText}'), ['', ExpandConstant('{app}')]);
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
        ExpandConstant('{cm:UninstServersTitle}'),
        ExpandConstant('{cm:UninstServersSubtitle}'),
        ExpandConstant('{cm:UninstServersDesc}'),
        False, True);
      for I := 0 to GetArrayLength(ServerKinds) - 1 do
        UninstallServersPage.Add(FmtMessage(ExpandConstant('{cm:UninstServerFound}'), [ServerLabels[I]]));
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
      if MsgBox(ExpandConstant('{cm:UninstKeepConfig}'), mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
  end;
end;

