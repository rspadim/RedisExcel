; RedisExcel installer + uninstaller (Inno Setup 6).
;
; Build:  ISCC /DAppVersion=<version> .\installer\RedisExcel.iss
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
; english is the default (based on Default.isl); the rest ship with Inno Setup 6
; except the two Chinese files, vendored under installer\languages (the official
; build has no Chinese .isl).
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "languages\ChineseSimplified.isl"
Name: "chinesetraditional"; MessagesFile: "languages\ChineseTraditional.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

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
english.ConnPassword=Password (optional):
english.OptionsPageTitle=Extra checks
english.OptionsPageSubtitle=Almost done
english.OptionsPageDesc=The installer tests the connection when it finishes. Optionally open a sample Excel workbook showing it working.
english.OptionsOpenExcel=Open a sample RedisExcel workbook at the end
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
english.UninstServerFailed=Could not remove %1 automatically (exit %2). Remove it manually if needed.
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
spanish.ConnPassword=Contraseña (opcional):
spanish.OptionsPageTitle=Comprobaciones
spanish.OptionsPageSubtitle=Casi listo
spanish.OptionsPageDesc=El instalador prueba la conexión al terminar. Opcionalmente abre un libro de Excel de ejemplo que muestra que funciona.
spanish.OptionsOpenExcel=Abrir un libro de RedisExcel de ejemplo al final
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
spanish.UninstServerFailed=No se pudo quitar %1 automaticamente (salida %2). Quítelo manualmente si es necesario.
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
brazilianportuguese.ConnPassword=Senha (opcional):
brazilianportuguese.OptionsPageTitle=Verificações
brazilianportuguese.OptionsPageSubtitle=Quase pronto
brazilianportuguese.OptionsPageDesc=O instalador testa a conexão ao terminar. Opcionalmente abre uma planilha de exemplo do RedisExcel mostrando que funcionou.
brazilianportuguese.OptionsOpenExcel=Abrir uma planilha de exemplo do RedisExcel no final
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
brazilianportuguese.UninstServerFailed=Nao foi possivel remover %1 automaticamente (saida %2). Remova manualmente se necessario.
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
brazilianportuguese.SummaryPageTitle=Suas configuracoes
brazilianportuguese.SummaryPageSubtitle=Confira a conexao que sera gravada no RedisExcel.json
brazilianportuguese.SummaryIntro=Estes valores serao salvos no RedisExcel.json na pasta de instalacao.
brazilianportuguese.SummaryLabel=Conexao
brazilianportuguese.SummaryMode=Modo
brazilianportuguese.SummarySimple=Simples
brazilianportuguese.SummaryAdvanced=Avancado


; ---- chinesesimplified ----
chinesesimplified.AdvPageDesc=è¿™å°±æ˜¯å®Œæ•´çš„ RedisExcel.jsonï¼Œå¯ç›´æŽ¥ç¼–è¾‘ï¼šæ¯ä¸ªé”®éƒ½å·²é¢„å¡«é»˜è®¤å€¼å’Œç®€çŸ­æ³¨é‡Šã€‚åŠ è½½é¡¹æŽ¥å— // æ³¨é‡Šï¼Œå› æ­¤å¯ä»¥ä¿ç•™ã€‚å¯åœ¨æ­¤ä¿®æ”¹ä¸»æœºã€åˆ«åã€å†™å…¥æ¨¡å¼ã€ConflationMsã€AsyncWritesâ€”â€”å…¨éƒ¨åœ¨æ­¤ã€‚
chinesesimplified.AdvPageTitle=é«˜çº§è¿žæŽ¥
chinesesimplified.ConfigWriteFailed=æ— æ³•å†™å…¥ %1ã€‚è¯·å‚ç…§ç¤ºä¾‹ï¼ˆRedisExcel.jsonï¼‰æ‰‹åŠ¨åˆ›å»ºã€‚
chinesesimplified.ConnHost=ä¸»æœºï¼ˆæˆ– host:portï¼‰ï¼š
chinesesimplified.ConnPageDesc=ä½¿ç”¨é»˜è®¤å€¼ï¼ˆlocalhost:6379ï¼‰è¯·ç‚¹â€œä¸‹ä¸€æ­¥â€ã€‚ä¹‹åŽå¯ç¼–è¾‘ RedisExcel.json ä¿®æ”¹ã€‚
chinesesimplified.ConnPageSubtitle=RedisExcel åº”è¿žæŽ¥åˆ°ä½•å¤„ï¼Ÿ
chinesesimplified.ConnPageTitle=Redis è¿žæŽ¥
chinesesimplified.ConnPassword=å¯†ç ï¼ˆå¯é€‰ï¼‰ï¼š
chinesesimplified.ConnPort=ç«¯å£ï¼š
chinesesimplified.FinishedText=RedisExcel å·²å®‰è£…åˆ°ï¼š%n%2%n%nå¦‚æžœåŠ è½½é¡¹æœªè‡ªåŠ¨åŠ è½½ï¼Œè¯·åœ¨ Excel ä¸­å¯ç”¨ï¼š%næ–‡ä»¶ > é€‰é¡¹ > åŠ è½½é¡¹ > è½¬åˆ°... > æµè§ˆ...ï¼Œé€‰æ‹© RedisExcel64-packed.xllï¼ˆ32 ä½ Office ç”¨ RedisExcel-packed.xllï¼‰ã€‚%n%nç„¶åŽåœ¨å•å…ƒæ ¼ä¸­è¯•è¯• =RedisUDFGet("preco_btc")ã€‚
chinesesimplified.GroupConfig=é…ç½®ï¼š
chinesesimplified.GroupExcel=Excelï¼š
chinesesimplified.HostRequired=è¯·è¾“å…¥ Redis ä¸»æœºï¼ˆä¾‹å¦‚ localhostï¼‰ã€‚
chinesesimplified.IconUninstall=å¸è½½ {#AppName}
chinesesimplified.ModeAdvanced=é«˜çº§ - è®¾ç½®æ•°æ®åº“ã€SSLã€å†™å…¥æ¨¡å¼å’Œæ›´æ–°è°ƒèŠ‚
chinesesimplified.ModePageDesc=ç®€æ˜“ä»…è¯¢é—®æœåŠ¡å™¨åœ°å€ã€‚é«˜çº§å¯è®¾ç½®æ•°æ®åº“ã€SSL/TLSã€å†™å…¥æ¨¡å¼å’Œå®žæ—¶æ›´æ–°è°ƒèŠ‚ã€‚ä¹‹åŽéšæ—¶å¯ç¼–è¾‘ RedisExcel.jsonï¼ˆæˆ–è¿è¡Œ Validate-Config.cmdï¼‰ã€‚
chinesesimplified.ModePageSubtitle=ç®€æ˜“è¿˜æ˜¯é«˜çº§å®‰è£…ï¼Ÿ
chinesesimplified.ModePageTitle=å®‰è£…æ¨¡å¼
chinesesimplified.ModeSimple=ç®€æ˜“ - ä»… Redis åœ°å€ï¼ˆæŽ¨èï¼‰
chinesesimplified.OptionsOpenExcel=ç»“æŸæ—¶æ‰“å¼€ç¤ºä¾‹ RedisExcel å·¥ä½œç°¿
chinesesimplified.OptionsPageDesc=å®‰è£…ç¨‹åºåœ¨ç»“æŸæ—¶æµ‹è¯•è¿žæŽ¥ã€‚å¯é€‰æ‹©æ‰“å¼€ç¤ºä¾‹ Excel å·¥ä½œç°¿ä»¥æ¼”ç¤ºå…¶å·¥ä½œæ–¹å¼ã€‚
chinesesimplified.OptionsPageSubtitle=å³å°†å®Œæˆ
chinesesimplified.OptionsPageTitle=é¢å¤–æ£€æŸ¥
chinesesimplified.ServerDockerRedis=æ˜¯ - Dockerï¼šRedisï¼ˆredis:latestï¼‰
chinesesimplified.ServerDockerValkey=æ˜¯ - Dockerï¼šValkeyï¼ˆvalkey/valkey:latestï¼‰
chinesesimplified.ServerFailed=æ— æ³•è‡ªåŠ¨è®¾ç½® %1ï¼ˆé€€å‡ºç  %2ï¼‰ã€‚%næ‚¨å¯ä»¥ç¨åŽå®‰è£…æœåŠ¡å™¨ï¼Œå¹¶åœ¨ RedisExcel.json ä¸­æŒ‡å‘å®ƒã€‚
chinesesimplified.ServerMemurai=æ˜¯ - Memuraiï¼ˆWindows æœ¬æœºæœåŠ¡ï¼Œéœ€è¦ç®¡ç†å‘˜æƒé™ï¼‰
chinesesimplified.ServerNameDockerRedis=Redisï¼ˆDockerï¼‰
chinesesimplified.ServerNameDockerValkey=Valkeyï¼ˆDockerï¼‰
chinesesimplified.ServerNameGeneric=æœåŠ¡å™¨
chinesesimplified.ServerNameMemurai=Memurai
chinesesimplified.ServerNameWslRedis=Redisï¼ˆWSLï¼‰
chinesesimplified.ServerNameWslValkey=Valkeyï¼ˆWSLï¼‰
chinesesimplified.ServerNone=å¦ - æˆ‘å·²æœ‰æœåŠ¡å™¨ï¼ˆæˆ–ç¨åŽè®¾ç½®ï¼‰
chinesesimplified.ServerPageDesc=RedisExcel éœ€è¦ Redisï¼ˆæˆ– Valkeyï¼‰æœåŠ¡å™¨ã€‚å¦‚æžœè¿˜æ²¡æœ‰ï¼Œè¯·é€‰æ‹©ä¸€ä¸ªé€‰é¡¹ã€‚å…ˆå®‰è£…æœåŠ¡å™¨ï¼Œå†å®‰è£… RedisExcelï¼Œæœ€åŽæµ‹è¯•è¿žæŽ¥ã€‚Docker æœ€ç®€å•ã€‚
chinesesimplified.ServerPageSubtitle=çŽ°åœ¨å®‰è£…å…¼å®¹ Redis çš„æœåŠ¡å™¨å—ï¼Ÿ
chinesesimplified.ServerPageTitle=Redis æœåŠ¡å™¨
chinesesimplified.ServerWslRedis=æ˜¯ - WSL2ï¼šRedisï¼ˆapt install redis-serverï¼‰
chinesesimplified.ServerWslValkey=æ˜¯ - WSL2ï¼šValkeyï¼ˆapt install valkey-serverï¼‰
chinesesimplified.SummaryAdvanced=é«˜çº§
chinesesimplified.SummaryIntro=è¿™äº›å€¼å°†ä¿å­˜åˆ°å®‰è£…æ–‡ä»¶å¤¹ä¸­çš„ RedisExcel.jsonã€‚
chinesesimplified.SummaryLabel=è¿žæŽ¥
chinesesimplified.SummaryMode=æ¨¡å¼
chinesesimplified.SummaryPageSubtitle=æŸ¥çœ‹å°†å†™å…¥ RedisExcel.json çš„è¿žæŽ¥
chinesesimplified.SummaryPageTitle=æ‚¨çš„è®¾ç½®
chinesesimplified.SummarySimple=ç®€æ˜“
chinesesimplified.TaskAutoLoadDesc=æ³¨å†ŒåŠ è½½é¡¹ï¼Œä½¿ Excel è‡ªåŠ¨åŠ è½½ï¼ˆæŽ¨èï¼‰
chinesesimplified.TaskConfigDesc=åœ¨å®‰è£…æ–‡ä»¶å¤¹ä¸­åˆ›å»º RedisExcel.jsonï¼ˆä¸»æœºè®¾ç½®ï¼‰
chinesesimplified.TaskNlogDesc=åˆ›å»º NLog.configï¼ˆåœ¨åŠ è½½é¡¹æ—å†™å…¥æ—¥å¿—æ–‡ä»¶ï¼‰
chinesesimplified.TestFailed=å·²å®‰è£…ï¼Œä½† %1 å¤„çš„ Redis æœªå“åº” PINGã€‚%nè¯·ç¡®è®¤æœåŠ¡å™¨æ­£åœ¨è¿è¡Œï¼ˆdocker psï¼‰ï¼Œå¦‚æžœå®ƒåœ¨å…¶ä»–ä¸»æœº/ç«¯å£ï¼Œè¯·è°ƒæ•´ RedisExcel.jsonã€‚
chinesesimplified.TestOk=å…¨éƒ¨å°±ç»ªï¼%1 å¤„çš„ Redis å·²å“åº”ã€‚%n%nä½¿ç”¨æ–¹æ³•ï¼šæ‰“å¼€ Excelã€‚å¦‚æžœåŠ è½½é¡¹æœªè‡ªåŠ¨åŠ è½½ï¼Œè¯·è½¬åˆ° æ–‡ä»¶ > é€‰é¡¹ > åŠ è½½é¡¹ > è½¬åˆ°... > æµè§ˆ...ï¼Œé€‰æ‹© %2 ä¸­çš„ .xllã€‚
chinesesimplified.UninstKeepConfig=ä¿ç•™æ‚¨çš„ RedisExcel.json è®¾ç½®å—ï¼Ÿ
chinesesimplified.UninstServerFound=%1  [å·²æ‰¾åˆ°]
chinesesimplified.UninstServerFailed=无法自动移除 %1（退出码 %2）。如有需要请手动移除。
chinesesimplified.UninstServersDesc=åœ¨æ­¤è®¡ç®—æœºä¸Šå‘çŽ°äº†è¿™äº›æœåŠ¡å™¨ã€‚å‹¾é€‰è¦éš RedisExcel ä¸€èµ·ç§»é™¤çš„é¡¹ã€‚%né»˜è®¤ä¸ç§»é™¤ä»»ä½•å†…å®¹ã€‚
chinesesimplified.UninstServersSubtitle=ä¹Ÿç§»é™¤å…¼å®¹ Redis çš„æœåŠ¡å™¨å—ï¼Ÿ
chinesesimplified.UninstServersTitle=æœåŠ¡å™¨
chinesesimplified.UpgradeConfigKept=å‘çŽ°å·²æœ‰çš„ RedisExcel.json å¹¶äºˆä»¥ä¿ç•™ï¼ˆæœªè¦†ç›–æ‚¨çš„è®¾ç½®ï¼‰ã€‚%n%nè¿™æ˜¯ä¸€æ¬¡é‡å¤§æ›´æ–°ï¼šé…ç½®æ ¼å¼å·²æ›´æ”¹ã€‚å¦‚æžœè¯¥æ–‡ä»¶ä½¿ç”¨æ—§æ ¼å¼ï¼Œåœ¨è¿ç§»ä¹‹å‰æ¯ä¸ªå•å…ƒæ ¼éƒ½ä¼šæ˜¾ç¤º #CONFIG é”™è¯¯ã€‚%n%nè¯·è¿è¡Œ Validate-Config.cmdï¼ˆæˆ–åœ¨ Excel ä¸­ä½¿ç”¨ =RedisCheckConfigFile()ï¼‰ï¼Œå¹¶å‚é˜… README çš„ â€œMigrating from the old shapeâ€ ä¸€èŠ‚ã€‚

; ---- chinesetraditional ----
chinesetraditional.AdvPageDesc=é€™å°±æ˜¯å®Œæ•´çš„ RedisExcel.jsonï¼Œå¯ç›´æŽ¥ç·¨è¼¯ï¼šæ¯å€‹éµéƒ½å·²é å¡«é è¨­å€¼èˆ‡ç°¡çŸ­è¨»è§£ã€‚å¢žç›Šé›†æŽ¥å— // è¨»è§£ï¼Œå› æ­¤å¯ä»¥ä¿ç•™ã€‚å¯åœ¨æ­¤ä¿®æ”¹ä¸»æ©Ÿã€åˆ¥åã€å¯«å…¥æ¨¡å¼ã€ConflationMsã€AsyncWritesâ€”â€”å…¨éƒ½åœ¨æ­¤ã€‚
chinesetraditional.AdvPageTitle=é€²éšŽé€£ç·š
chinesetraditional.ConfigWriteFailed=ç„¡æ³•å¯«å…¥ %1ã€‚è«‹ä¾ç¯„ä¾‹ï¼ˆRedisExcel.jsonï¼‰æ‰‹å‹•å»ºç«‹ã€‚
chinesetraditional.ConnHost=ä¸»æ©Ÿï¼ˆæˆ– host:portï¼‰ï¼š
chinesetraditional.ConnPageDesc=ä½¿ç”¨é è¨­å€¼ï¼ˆlocalhost:6379ï¼‰è«‹æŒ‰ã€Œä¸‹ä¸€æ­¥ã€ã€‚ä¹‹å¾Œå¯ç·¨è¼¯ RedisExcel.json ä¿®æ”¹ã€‚
chinesetraditional.ConnPageSubtitle=RedisExcel æ‡‰é€£ç·šåˆ°ä½•è™•ï¼Ÿ
chinesetraditional.ConnPageTitle=Redis é€£ç·š
chinesetraditional.ConnPassword=å¯†ç¢¼ï¼ˆé¸å¡«ï¼‰ï¼š
chinesetraditional.ConnPort=é€£æŽ¥åŸ ï¼š
chinesetraditional.FinishedText=RedisExcel å·²å®‰è£æ–¼ï¼š%n%2%n%nè‹¥å¢žç›Šé›†æœªè‡ªå‹•è¼‰å…¥ï¼Œè«‹åœ¨ Excel ä¸­å•Ÿç”¨ï¼š%næª”æ¡ˆ > é¸é … > å¢žç›Šé›† > å‰å¾€... > ç€è¦½...ï¼Œé¸æ“‡ RedisExcel64-packed.xllï¼ˆ32 ä½å…ƒ Office è«‹ç”¨ RedisExcel-packed.xllï¼‰ã€‚%n%næŽ¥è‘—åœ¨å„²å­˜æ ¼ä¸­è©¦è©¦ =RedisUDFGet("preco_btc")ã€‚
chinesetraditional.GroupConfig=è¨­å®šï¼š
chinesetraditional.GroupExcel=Excelï¼š
chinesetraditional.HostRequired=è«‹è¼¸å…¥ Redis ä¸»æ©Ÿï¼ˆä¾‹å¦‚ localhostï¼‰ã€‚
chinesetraditional.IconUninstall=è§£é™¤å®‰è£ {#AppName}
chinesetraditional.ModeAdvanced=é€²éšŽ - è¨­å®šè³‡æ–™åº«ã€SSLã€å¯«å…¥æ¨¡å¼èˆ‡æ›´æ–°å¾®èª¿
chinesetraditional.ModePageDesc=ç°¡æ˜“åªè©¢å•ä¼ºæœå™¨ä½å€ã€‚é€²éšŽå¯è¨­å®šè³‡æ–™åº«ã€SSL/TLSã€å¯«å…¥æ¨¡å¼èˆ‡å³æ™‚æ›´æ–°å¾®èª¿ã€‚ä¹‹å¾Œéš¨æ™‚å¯ç·¨è¼¯ RedisExcel.jsonï¼ˆæˆ–åŸ·è¡Œ Validate-Config.cmdï¼‰ã€‚
chinesetraditional.ModePageSubtitle=ç°¡æ˜“æˆ–é€²éšŽå®‰è£ï¼Ÿ
chinesetraditional.ModePageTitle=å®‰è£æ¨¡å¼
chinesetraditional.ModeSimple=ç°¡æ˜“ - åƒ… Redis ä½å€ï¼ˆå»ºè­°ï¼‰
chinesetraditional.OptionsOpenExcel=çµæŸæ™‚é–‹å•Ÿç¯„ä¾‹ RedisExcel æ´»é ç°¿
chinesetraditional.OptionsPageDesc=å®‰è£ç¨‹å¼æœƒåœ¨çµæŸæ™‚æ¸¬è©¦é€£ç·šã€‚å¯é¸æ“‡é–‹å•Ÿç¯„ä¾‹ Excel æ´»é ç°¿ä»¥å±•ç¤ºå…¶é‹ä½œæ–¹å¼ã€‚
chinesetraditional.OptionsPageSubtitle=å³å°‡å®Œæˆ
chinesetraditional.OptionsPageTitle=é¡å¤–æª¢æŸ¥
chinesetraditional.ServerDockerRedis=æ˜¯ - Dockerï¼šRedisï¼ˆredis:latestï¼‰
chinesetraditional.ServerDockerValkey=æ˜¯ - Dockerï¼šValkeyï¼ˆvalkey/valkey:latestï¼‰
chinesetraditional.ServerFailed=ç„¡æ³•è‡ªå‹•è¨­å®š %1ï¼ˆçµæŸç¢¼ %2ï¼‰ã€‚%næ‚¨å¯ä»¥ç¨å¾Œå®‰è£ä¼ºæœå™¨ï¼Œä¸¦åœ¨ RedisExcel.json ä¸­æŒ‡å‘å®ƒã€‚
chinesetraditional.ServerMemurai=æ˜¯ - Memuraiï¼ˆWindows åŽŸç”Ÿæœå‹™ï¼Œéœ€è¦ç³»çµ±ç®¡ç†å“¡æ¬Šé™ï¼‰
chinesetraditional.ServerNameDockerRedis=Redisï¼ˆDockerï¼‰
chinesetraditional.ServerNameDockerValkey=Valkeyï¼ˆDockerï¼‰
chinesetraditional.ServerNameGeneric=ä¼ºæœå™¨
chinesetraditional.ServerNameMemurai=Memurai
chinesetraditional.ServerNameWslRedis=Redisï¼ˆWSLï¼‰
chinesetraditional.ServerNameWslValkey=Valkeyï¼ˆWSLï¼‰
chinesetraditional.ServerNone=å¦ - æˆ‘å·²æœ‰ä¼ºæœå™¨ï¼ˆæˆ–ç¨å¾Œè¨­å®šï¼‰
chinesetraditional.ServerPageDesc=RedisExcel éœ€è¦ Redisï¼ˆæˆ– Valkeyï¼‰ä¼ºæœå™¨ã€‚å¦‚æžœé‚„æ²’æœ‰ï¼Œè«‹é¸æ“‡ä¸€å€‹é¸é …ã€‚æœƒå…ˆå®‰è£ä¼ºæœå™¨ï¼Œå†å®‰è£ RedisExcelï¼Œæœ€å¾Œæ¸¬è©¦é€£ç·šã€‚Docker æœ€ç°¡å–®ã€‚
chinesetraditional.ServerPageSubtitle=ç¾åœ¨å®‰è£ç›¸å®¹æ–¼ Redis çš„ä¼ºæœå™¨å—Žï¼Ÿ
chinesetraditional.ServerPageTitle=Redis ä¼ºæœå™¨
chinesetraditional.ServerWslRedis=æ˜¯ - WSL2ï¼šRedisï¼ˆapt install redis-serverï¼‰
chinesetraditional.ServerWslValkey=æ˜¯ - WSL2ï¼šValkeyï¼ˆapt install valkey-serverï¼‰
chinesetraditional.SummaryAdvanced=é€²éšŽ
chinesetraditional.SummaryIntro=é€™äº›å€¼å°‡å„²å­˜åˆ°å®‰è£è³‡æ–™å¤¾ä¸­çš„ RedisExcel.jsonã€‚
chinesetraditional.SummaryLabel=é€£ç·š
chinesetraditional.SummaryMode=æ¨¡å¼
chinesetraditional.SummaryPageSubtitle=æª¢é–±å°‡å¯«å…¥ RedisExcel.json çš„é€£ç·š
chinesetraditional.SummaryPageTitle=æ‚¨çš„è¨­å®š
chinesetraditional.SummarySimple=ç°¡æ˜“
chinesetraditional.TaskAutoLoadDesc=è¨»å†Šå¢žç›Šé›†ï¼Œè®“ Excel è‡ªå‹•è¼‰å…¥ï¼ˆå»ºè­°ï¼‰
chinesetraditional.TaskConfigDesc=åœ¨å®‰è£è³‡æ–™å¤¾ä¸­å»ºç«‹ RedisExcel.jsonï¼ˆä¸»æ©Ÿè¨­å®šï¼‰
chinesetraditional.TaskNlogDesc=å»ºç«‹ NLog.configï¼ˆåœ¨å¢žç›Šé›†æ—å¯«å…¥è¨˜éŒ„æª”ï¼‰
chinesetraditional.TestFailed=å·²å®‰è£ï¼Œä½† %1 çš„ Redis æœªå›žæ‡‰ PINGã€‚%nè«‹ç¢ºèªä¼ºæœå™¨æ­£åœ¨åŸ·è¡Œï¼ˆdocker psï¼‰ï¼Œè‹¥åœ¨å…¶ä»–ä¸»æ©Ÿ/é€£æŽ¥åŸ ï¼Œè«‹èª¿æ•´ RedisExcel.jsonã€‚
chinesetraditional.TestOk=å…¨éƒ¨å°±ç·’ï¼%1 çš„ Redis å·²å›žæ‡‰ã€‚%n%nä½¿ç”¨æ–¹æ³•ï¼šé–‹å•Ÿ Excelã€‚è‹¥å¢žç›Šé›†æœªè‡ªå‹•è¼‰å…¥ï¼Œè«‹å‰å¾€ æª”æ¡ˆ > é¸é … > å¢žç›Šé›† > å‰å¾€... > ç€è¦½...ï¼Œé¸æ“‡ %2 ä¸­çš„ .xllã€‚
chinesetraditional.UninstKeepConfig=ä¿ç•™æ‚¨çš„ RedisExcel.json è¨­å®šå—Žï¼Ÿ
chinesetraditional.UninstServerFound=%1  [å·²æ‰¾åˆ°]
chinesetraditional.UninstServerFailed=無法自動移除 %1（結束碼 %2）。如有需要請手動移除。
chinesetraditional.UninstServersDesc=åœ¨æ­¤é›»è…¦ä¸Šæ‰¾åˆ°é€™äº›ä¼ºæœå™¨ã€‚å‹¾é¸è¦éš¨ RedisExcel ä¸€ä½µç§»é™¤çš„é …ç›®ã€‚%né è¨­ä¸ç§»é™¤ä»»ä½•é …ç›®ã€‚
chinesetraditional.UninstServersSubtitle=ä¹Ÿç§»é™¤ç›¸å®¹æ–¼ Redis çš„ä¼ºæœå™¨å—Žï¼Ÿ
chinesetraditional.UninstServersTitle=ä¼ºæœå™¨
chinesetraditional.UpgradeConfigKept=æ‰¾åˆ°ç¾æœ‰çš„ RedisExcel.json ä¸¦äºˆä»¥ä¿ç•™ï¼ˆæœªè¦†å¯«æ‚¨çš„è¨­å®šï¼‰ã€‚%n%né€™æ˜¯ä¸€æ¬¡é‡å¤§æ›´æ–°ï¼šè¨­å®šæ ¼å¼å·²è®Šæ›´ã€‚è‹¥è©²æª”æ¡ˆä½¿ç”¨èˆŠæ ¼å¼ï¼Œåœ¨ç§»è½‰ä¹‹å‰æ¯å€‹å„²å­˜æ ¼éƒ½æœƒé¡¯ç¤º #CONFIG éŒ¯èª¤ã€‚%n%nè«‹åŸ·è¡Œ Validate-Config.cmdï¼ˆæˆ–åœ¨ Excel ä¸­ä½¿ç”¨ =RedisCheckConfigFile()ï¼‰ï¼Œä¸¦åƒé–± README çš„ã€ŒMigrating from the old shapeã€ä¸€ç¯€ã€‚

; ---- russian ----
russian.AdvPageDesc=Ð­Ñ‚Ð¾ Ð²ÐµÑÑŒ Ð²Ð°Ñˆ RedisExcel.json, Ð³Ð¾Ñ‚Ð¾Ð²Ñ‹Ð¹ Ðº Ð¿Ñ€Ð°Ð²ÐºÐµ: ÐºÐ°Ð¶Ð´Ñ‹Ð¹ ÐºÐ»ÑŽÑ‡ ÑƒÐ¶Ðµ Ð·Ð°Ð¿Ð¾Ð»Ð½ÐµÐ½ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸ÐµÐ¼ Ð¿Ð¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ Ð¸ ÐºÐ¾Ñ€Ð¾Ñ‚ÐºÐ¸Ð¼ ÐºÐ¾Ð¼Ð¼ÐµÐ½Ñ‚Ð°Ñ€Ð¸ÐµÐ¼. ÐÐ´Ð´-Ð¸Ð½ Ð¿Ñ€Ð¸Ð½Ð¸Ð¼Ð°ÐµÑ‚ ÐºÐ¾Ð¼Ð¼ÐµÐ½Ñ‚Ð°Ñ€Ð¸Ð¸ //, Ð¿Ð¾ÑÑ‚Ð¾Ð¼Ñƒ Ð¸Ñ… Ð¼Ð¾Ð¶Ð½Ð¾ Ð¾ÑÑ‚Ð°Ð²Ð¸Ñ‚ÑŒ. ÐœÐµÐ½ÑÐ¹Ñ‚Ðµ Ñ…Ð¾ÑÑ‚Ñ‹, Ð¿ÑÐµÐ²Ð´Ð¾Ð½Ð¸Ð¼Ñ‹, Ñ€ÐµÐ¶Ð¸Ð¼ Ð·Ð°Ð¿Ð¸ÑÐ¸, ConflationMs, AsyncWrites - Ð²ÑÑ‘ Ð·Ð´ÐµÑÑŒ.
russian.AdvPageTitle=Ð”Ð¾Ð¿Ð¾Ð»Ð½Ð¸Ñ‚ÐµÐ»ÑŒÐ½Ð¾Ðµ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ
russian.ConfigWriteFailed=ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð·Ð°Ð¿Ð¸ÑÐ°Ñ‚ÑŒ %1. Ð¡Ð¾Ð·Ð´Ð°Ð¹Ñ‚Ðµ ÐµÐ³Ð¾ Ð²Ñ€ÑƒÑ‡Ð½ÑƒÑŽ Ð¿Ð¾ Ð¾Ð±Ñ€Ð°Ð·Ñ†Ñƒ (RedisExcel.json).
russian.ConnHost=Ð¥Ð¾ÑÑ‚ (Ð¸Ð»Ð¸ host:port):
russian.ConnPageDesc=ÐÐ°Ð¶Ð¼Ð¸Ñ‚Ðµ Â«Ð”Ð°Ð»ÐµÐµÂ» Ð´Ð»Ñ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ñ Ð¿Ð¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ (localhost:6379). ÐŸÐ¾Ð·Ð¶Ðµ ÑÑ‚Ð¾ Ð¼Ð¾Ð¶Ð½Ð¾ Ð¸Ð·Ð¼ÐµÐ½Ð¸Ñ‚ÑŒ Ð² RedisExcel.json.
russian.ConnPageSubtitle=ÐšÑƒÐ´Ð° Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡Ð°Ñ‚ÑŒÑÑ RedisExcel?
russian.ConnPageTitle=ÐŸÐ¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ Ðº Redis
russian.ConnPassword=ÐŸÐ°Ñ€Ð¾Ð»ÑŒ (Ð½ÐµÐ¾Ð±ÑÐ·Ð°Ñ‚ÐµÐ»ÑŒÐ½Ð¾):
russian.ConnPort=ÐŸÐ¾Ñ€Ñ‚:
russian.FinishedText=RedisExcel ÑƒÑÑ‚Ð°Ð½Ð¾Ð²Ð»ÐµÐ½ Ð²:%n%2%n%nÐ•ÑÐ»Ð¸ Ð°Ð´Ð´-Ð¸Ð½ Ð½Ðµ Ð·Ð°Ð³Ñ€ÑƒÐ¶Ð°ÐµÑ‚ÑÑ Ð°Ð²Ñ‚Ð¾Ð¼Ð°Ñ‚Ð¸Ñ‡ÐµÑÐºÐ¸, Ð²ÐºÐ»ÑŽÑ‡Ð¸Ñ‚Ðµ ÐµÐ³Ð¾ Ð² Excel:%nÐ¤Ð°Ð¹Ð» > ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ > ÐÐ°Ð´ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸ > ÐŸÐµÑ€ÐµÐ¹Ñ‚Ð¸... > ÐžÐ±Ð·Ð¾Ñ€... Ð¸ Ð²Ñ‹Ð±ÐµÑ€Ð¸Ñ‚Ðµ RedisExcel64-packed.xll (Ð¸Ð»Ð¸ RedisExcel-packed.xll Ð´Ð»Ñ 32-Ð±Ð¸Ñ‚Ð½Ð¾Ð³Ð¾ Office).%n%nÐ—Ð°Ñ‚ÐµÐ¼ Ð¿Ð¾Ð¿Ñ€Ð¾Ð±ÑƒÐ¹Ñ‚Ðµ =RedisUDFGet("preco_btc") Ð² ÑÑ‡ÐµÐ¹ÐºÐµ.
russian.GroupConfig=ÐšÐ¾Ð½Ñ„Ð¸Ð³ÑƒÑ€Ð°Ñ†Ð¸Ñ:
russian.GroupExcel=Excel:
russian.HostRequired=Ð’Ð²ÐµÐ´Ð¸Ñ‚Ðµ Ñ…Ð¾ÑÑ‚ Redis (Ð½Ð°Ð¿Ñ€Ð¸Ð¼ÐµÑ€, localhost).
russian.IconUninstall=Ð£Ð´Ð°Ð»Ð¸Ñ‚ÑŒ {#AppName}
russian.ModeAdvanced=Ð Ð°ÑÑˆÐ¸Ñ€ÐµÐ½Ð½Ñ‹Ð¹ - Ð±Ð°Ð·Ð° Ð´Ð°Ð½Ð½Ñ‹Ñ…, SSL, Ñ€ÐµÐ¶Ð¸Ð¼ Ð·Ð°Ð¿Ð¸ÑÐ¸ Ð¸ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ° Ð¾Ð±Ð½Ð¾Ð²Ð»ÐµÐ½Ð¸Ð¹
russian.ModePageDesc=ÐŸÑ€Ð¾ÑÑ‚Ð¾Ð¹ Ð·Ð°Ð¿Ñ€Ð°ÑˆÐ¸Ð²Ð°ÐµÑ‚ Ñ‚Ð¾Ð»ÑŒÐºÐ¾ Ð°Ð´Ñ€ÐµÑ ÑÐµÑ€Ð²ÐµÑ€Ð°. Ð Ð°ÑÑˆÐ¸Ñ€ÐµÐ½Ð½Ñ‹Ð¹ Ð¿Ð¾Ð·Ð²Ð¾Ð»ÑÐµÑ‚ Ð·Ð°Ð´Ð°Ñ‚ÑŒ Ð±Ð°Ð·Ñƒ Ð´Ð°Ð½Ð½Ñ‹Ñ…, SSL/TLS, Ñ€ÐµÐ¶Ð¸Ð¼ Ð·Ð°Ð¿Ð¸ÑÐ¸ Ð¸ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÑƒ Ð¾Ð±Ð½Ð¾Ð²Ð»ÐµÐ½Ð¸Ð¹ Ð² Ñ€ÐµÐ°Ð»ÑŒÐ½Ð¾Ð¼ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð¸. RedisExcel.json Ð²ÑÐµÐ³Ð´Ð° Ð¼Ð¾Ð¶Ð½Ð¾ Ð¾Ñ‚Ñ€ÐµÐ´Ð°ÐºÑ‚Ð¸Ñ€Ð¾Ð²Ð°Ñ‚ÑŒ Ð¿Ð¾Ð·Ð¶Ðµ (Ð¸Ð»Ð¸ Ð·Ð°Ð¿ÑƒÑÑ‚Ð¸Ñ‚ÑŒ Validate-Config.cmd).
russian.ModePageSubtitle=ÐŸÑ€Ð¾ÑÑ‚Ð°Ñ Ð¸Ð»Ð¸ Ñ€Ð°ÑÑˆÐ¸Ñ€ÐµÐ½Ð½Ð°Ñ ÑƒÑÑ‚Ð°Ð½Ð¾Ð²ÐºÐ°?
russian.ModePageTitle=Ð ÐµÐ¶Ð¸Ð¼ ÑƒÑÑ‚Ð°Ð½Ð¾Ð²ÐºÐ¸
russian.ModeSimple=ÐŸÑ€Ð¾ÑÑ‚Ð¾Ð¹ - Ñ‚Ð¾Ð»ÑŒÐºÐ¾ Ð°Ð´Ñ€ÐµÑ Redis (Ñ€ÐµÐºÐ¾Ð¼ÐµÐ½Ð´ÑƒÐµÑ‚ÑÑ)
russian.OptionsOpenExcel=ÐžÑ‚ÐºÑ€Ñ‹Ñ‚ÑŒ Ð¿Ñ€Ð¸Ð¼ÐµÑ€ ÐºÐ½Ð¸Ð³Ð¸ RedisExcel Ð² ÐºÐ¾Ð½Ñ†Ðµ
russian.OptionsPageDesc=Ð£ÑÑ‚Ð°Ð½Ð¾Ð²Ñ‰Ð¸Ðº Ð¿Ñ€Ð¾Ð²ÐµÑ€Ð¸Ñ‚ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ Ð¿Ð¾ Ð·Ð°Ð²ÐµÑ€ÑˆÐµÐ½Ð¸Ð¸. ÐŸÑ€Ð¸ Ð¶ÐµÐ»Ð°Ð½Ð¸Ð¸ Ð¾Ñ‚ÐºÑ€Ð¾ÐµÑ‚ÑÑ Ð¿Ñ€Ð¸Ð¼ÐµÑ€ ÐºÐ½Ð¸Ð³Ð¸ Excel Ñ Ð´ÐµÐ¼Ð¾Ð½ÑÑ‚Ñ€Ð°Ñ†Ð¸ÐµÐ¹ Ñ€Ð°Ð±Ð¾Ñ‚Ñ‹.
russian.OptionsPageSubtitle=ÐŸÐ¾Ñ‡Ñ‚Ð¸ Ð³Ð¾Ñ‚Ð¾Ð²Ð¾
russian.OptionsPageTitle=Ð”Ð¾Ð¿Ð¾Ð»Ð½Ð¸Ñ‚ÐµÐ»ÑŒÐ½Ñ‹Ðµ Ð¿Ñ€Ð¾Ð²ÐµÑ€ÐºÐ¸
russian.ServerDockerRedis=Ð”Ð° - Docker: Redis (redis:latest)
russian.ServerDockerValkey=Ð”Ð° - Docker: Valkey (valkey/valkey:latest)
russian.ServerFailed=ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¸Ñ‚ÑŒ %1 Ð°Ð²Ñ‚Ð¾Ð¼Ð°Ñ‚Ð¸Ñ‡ÐµÑÐºÐ¸ (ÐºÐ¾Ð´ %2).%nÐ’Ñ‹ Ð¼Ð¾Ð¶ÐµÑ‚Ðµ ÑƒÑÑ‚Ð°Ð½Ð¾Ð²Ð¸Ñ‚ÑŒ ÑÐµÑ€Ð²ÐµÑ€ Ð¿Ð¾Ð·Ð¶Ðµ Ð¸ ÑƒÐºÐ°Ð·Ð°Ñ‚ÑŒ ÐµÐ³Ð¾ Ð² RedisExcel.json.
russian.ServerMemurai=Ð”Ð° - Memurai (ÑÐ»ÑƒÐ¶Ð±Ð° Windows, Ð½ÑƒÐ¶Ð½Ñ‹ Ð¿Ñ€Ð°Ð²Ð° Ð°Ð´Ð¼Ð¸Ð½Ð¸ÑÑ‚Ñ€Ð°Ñ‚Ð¾Ñ€Ð°)
russian.ServerNameDockerRedis=Redis (Docker)
russian.ServerNameDockerValkey=Valkey (Docker)
russian.ServerNameGeneric=ÑÐµÑ€Ð²ÐµÑ€
russian.ServerNameMemurai=Memurai
russian.ServerNameWslRedis=Redis (WSL)
russian.ServerNameWslValkey=Valkey (WSL)
russian.ServerNone=ÐÐµÑ‚ - ÑÐµÑ€Ð²ÐµÑ€ ÑƒÐ¶Ðµ ÐµÑÑ‚ÑŒ (Ð¸Ð»Ð¸ Ð½Ð°ÑÑ‚Ñ€Ð¾ÑŽ Ð¿Ð¾Ð·Ð¶Ðµ)
russian.ServerPageDesc=RedisExcel Ð½ÑƒÐ¶ÐµÐ½ ÑÐµÑ€Ð²ÐµÑ€ Redis (Ð¸Ð»Ð¸ Valkey). Ð’Ñ‹Ð±ÐµÑ€Ð¸Ñ‚Ðµ Ð²Ð°Ñ€Ð¸Ð°Ð½Ñ‚, ÐµÑÐ»Ð¸ ÐµÐ³Ð¾ ÐµÑ‰Ñ‘ Ð½ÐµÑ‚. Ð¡Ð½Ð°Ñ‡Ð°Ð»Ð° ÑƒÑÑ‚Ð°Ð½Ð°Ð²Ð»Ð¸Ð²Ð°ÐµÑ‚ÑÑ ÑÐµÑ€Ð²ÐµÑ€, Ð·Ð°Ñ‚ÐµÐ¼ RedisExcel, Ð° Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ Ð¿Ñ€Ð¾Ð²ÐµÑ€ÑÐµÑ‚ÑÑ Ð² ÐºÐ¾Ð½Ñ†Ðµ. Docker - ÑÐ°Ð¼Ñ‹Ð¹ Ð¿Ñ€Ð¾ÑÑ‚Ð¾Ð¹ Ð²Ð°Ñ€Ð¸Ð°Ð½Ñ‚.
russian.ServerPageSubtitle=Ð£ÑÑ‚Ð°Ð½Ð¾Ð²Ð¸Ñ‚ÑŒ ÑÐµÐ¹Ñ‡Ð°Ñ ÑÐ¾Ð²Ð¼ÐµÑÑ‚Ð¸Ð¼Ñ‹Ð¹ Ñ Redis ÑÐµÑ€Ð²ÐµÑ€?
russian.ServerPageTitle=Ð¡ÐµÑ€Ð²ÐµÑ€ Redis
russian.ServerWslRedis=Ð”Ð° - WSL2: Redis (apt install redis-server)
russian.ServerWslValkey=Ð”Ð° - WSL2: Valkey (apt install valkey-server)
russian.SummaryAdvanced=Ð Ð°ÑÑˆÐ¸Ñ€ÐµÐ½Ð½Ñ‹Ð¹
russian.SummaryIntro=Ð­Ñ‚Ð¸ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ñ Ð±ÑƒÐ´ÑƒÑ‚ ÑÐ¾Ñ…Ñ€Ð°Ð½ÐµÐ½Ñ‹ Ð² RedisExcel.json Ð² Ð¿Ð°Ð¿ÐºÐµ ÑƒÑÑ‚Ð°Ð½Ð¾Ð²ÐºÐ¸.
russian.SummaryLabel=ÐŸÐ¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ
russian.SummaryMode=Ð ÐµÐ¶Ð¸Ð¼
russian.SummaryPageSubtitle=ÐŸÑ€Ð¾Ð²ÐµÑ€ÑŒÑ‚Ðµ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡ÐµÐ½Ð¸Ðµ, ÐºÐ¾Ñ‚Ð¾Ñ€Ð¾Ðµ Ð±ÑƒÐ´ÐµÑ‚ Ð·Ð°Ð¿Ð¸ÑÐ°Ð½Ð¾ Ð² RedisExcel.json
russian.SummaryPageTitle=Ð’Ð°ÑˆÐ¸ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸
russian.SummarySimple=ÐŸÑ€Ð¾ÑÑ‚Ð¾Ð¹
russian.TaskAutoLoadDesc=Ð—Ð°Ñ€ÐµÐ³Ð¸ÑÑ‚Ñ€Ð¸Ñ€Ð¾Ð²Ð°Ñ‚ÑŒ Ð°Ð´Ð´-Ð¸Ð½, Ñ‡Ñ‚Ð¾Ð±Ñ‹ Excel Ð·Ð°Ð³Ñ€ÑƒÐ¶Ð°Ð» ÐµÐ³Ð¾ Ð°Ð²Ñ‚Ð¾Ð¼Ð°Ñ‚Ð¸Ñ‡ÐµÑÐºÐ¸ (Ñ€ÐµÐºÐ¾Ð¼ÐµÐ½Ð´ÑƒÐµÑ‚ÑÑ)
russian.TaskConfigDesc=Ð¡Ð¾Ð·Ð´Ð°Ñ‚ÑŒ RedisExcel.json Ð² Ð¿Ð°Ð¿ÐºÐµ ÑƒÑÑ‚Ð°Ð½Ð¾Ð²ÐºÐ¸ (Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸ Ñ…Ð¾ÑÑ‚Ð°)
russian.TaskNlogDesc=Ð¡Ð¾Ð·Ð´Ð°Ñ‚ÑŒ NLog.config (Ð¿Ð¸ÑÐ°Ñ‚ÑŒ Ð»Ð¾Ð³-Ñ„Ð°Ð¹Ð» Ñ€ÑÐ´Ð¾Ð¼ Ñ Ð°Ð´Ð´-Ð¸Ð½Ð¾Ð¼)
russian.TestFailed=Ð£ÑÑ‚Ð°Ð½Ð¾Ð²Ð»ÐµÐ½Ð¾, Ð½Ð¾ Redis Ð¿Ð¾ Ð°Ð´Ñ€ÐµÑÑƒ %1 Ð½Ðµ Ð¾Ñ‚Ð²ÐµÑ‚Ð¸Ð» Ð½Ð° PING.%nÐŸÑ€Ð¾Ð²ÐµÑ€ÑŒÑ‚Ðµ, Ñ‡Ñ‚Ð¾ ÑÐµÑ€Ð²ÐµÑ€ Ð·Ð°Ð¿ÑƒÑ‰ÐµÐ½ (docker ps), Ð¸ Ð¿Ð¾Ð¿Ñ€Ð°Ð²ÑŒÑ‚Ðµ RedisExcel.json, ÐµÑÐ»Ð¸ Ð¾Ð½ Ð½Ð° Ð´Ñ€ÑƒÐ³Ð¾Ð¼ Ñ…Ð¾ÑÑ‚Ðµ/Ð¿Ð¾Ñ€Ñ‚Ñƒ.
russian.TestOk=Ð“Ð¾Ñ‚Ð¾Ð²Ð¾! Redis Ð¾Ñ‚Ð²ÐµÑ‚Ð¸Ð» Ð¿Ð¾ Ð°Ð´Ñ€ÐµÑÑƒ %1.%n%nÐ§Ñ‚Ð¾Ð±Ñ‹ Ð½Ð°Ñ‡Ð°Ñ‚ÑŒ: Ð¾Ñ‚ÐºÑ€Ð¾Ð¹Ñ‚Ðµ Excel. Ð•ÑÐ»Ð¸ Ð°Ð´Ð´-Ð¸Ð½ Ð½Ðµ Ð·Ð°Ð³Ñ€ÑƒÐ¶Ð°ÐµÑ‚ÑÑ Ð°Ð²Ñ‚Ð¾Ð¼Ð°Ñ‚Ð¸Ñ‡ÐµÑÐºÐ¸, Ð¿ÐµÑ€ÐµÐ¹Ð´Ð¸Ñ‚Ðµ Ð² Ð¤Ð°Ð¹Ð» > ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ > ÐÐ°Ð´ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸ > ÐŸÐµÑ€ÐµÐ¹Ñ‚Ð¸... > ÐžÐ±Ð·Ð¾Ñ€... Ð¸ Ð²Ñ‹Ð±ÐµÑ€Ð¸Ñ‚Ðµ .xll Ð² %2.
russian.UninstKeepConfig=Ð¡Ð¾Ñ…Ñ€Ð°Ð½Ð¸Ñ‚ÑŒ Ð²Ð°ÑˆÐ¸ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸ RedisExcel.json?
russian.UninstServerFound=%1  [Ð½Ð°Ð¹Ð´ÐµÐ½Ð¾]
russian.UninstServerFailed=Не удалось автоматически удалить %1 (код %2). Удалите вручную при необходимости.
russian.UninstServersDesc=Ð­Ñ‚Ð¸ ÑÐµÑ€Ð²ÐµÑ€Ñ‹ Ð½Ð°Ð¹Ð´ÐµÐ½Ñ‹ Ð½Ð° ÐºÐ¾Ð¼Ð¿ÑŒÑŽÑ‚ÐµÑ€Ðµ. ÐžÑ‚Ð¼ÐµÑ‚ÑŒÑ‚Ðµ Ñ‚Ðµ, ÐºÐ¾Ñ‚Ð¾Ñ€Ñ‹Ðµ Ð½ÑƒÐ¶Ð½Ð¾ ÑƒÐ´Ð°Ð»Ð¸Ñ‚ÑŒ Ð²Ð¼ÐµÑÑ‚Ðµ Ñ RedisExcel.%nÐŸÐ¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ ÐÐ˜Ð§Ð•Ð“Ðž Ð½Ðµ ÑƒÐ´Ð°Ð»ÑÐµÑ‚ÑÑ.
russian.UninstServersSubtitle=Ð£Ð´Ð°Ð»Ð¸Ñ‚ÑŒ Ñ‚Ð°ÐºÐ¶Ðµ ÑÐ¾Ð²Ð¼ÐµÑÑ‚Ð¸Ð¼Ñ‹Ðµ Ñ Redis ÑÐµÑ€Ð²ÐµÑ€Ñ‹?
russian.UninstServersTitle=Ð¡ÐµÑ€Ð²ÐµÑ€Ñ‹
russian.UpgradeConfigKept=ÐÐ°Ð¹Ð´ÐµÐ½ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÑŽÑ‰Ð¸Ð¹ RedisExcel.json, Ð¾Ð½ ÑÐ¾Ñ…Ñ€Ð°Ð½Ñ‘Ð½ (Ð²Ð°ÑˆÐ¸ Ð½Ð°ÑÑ‚Ñ€Ð¾Ð¹ÐºÐ¸ Ð½Ðµ Ð¿ÐµÑ€ÐµÐ·Ð°Ð¿Ð¸ÑÐ°Ð½Ñ‹).%n%nÐ­Ñ‚Ð¾ ÐšÐ Ð£ÐŸÐÐžÐ• Ð¾Ð±Ð½Ð¾Ð²Ð»ÐµÐ½Ð¸Ðµ: Ñ„Ð¾Ñ€Ð¼Ð°Ñ‚ ÐºÐ¾Ð½Ñ„Ð¸Ð³ÑƒÑ€Ð°Ñ†Ð¸Ð¸ Ð¸Ð·Ð¼ÐµÐ½Ð¸Ð»ÑÑ. Ð•ÑÐ»Ð¸ Ñ„Ð°Ð¹Ð» Ð² ÑÑ‚Ð°Ñ€Ð¾Ð¼ Ñ„Ð¾Ñ€Ð¼Ð°Ñ‚Ðµ, Ð² ÐºÐ°Ð¶Ð´Ð¾Ð¹ ÑÑ‡ÐµÐ¹ÐºÐµ Ð±ÑƒÐ´ÐµÑ‚ Ð¾ÑˆÐ¸Ð±ÐºÐ° #CONFIG, Ð¿Ð¾ÐºÐ° ÐµÐ³Ð¾ Ð½Ðµ Ð¿ÐµÑ€ÐµÐ½ÐµÑÑ‚Ð¸.%n%nÐ—Ð°Ð¿ÑƒÑÑ‚Ð¸Ñ‚Ðµ Validate-Config.cmd (Ð¸Ð»Ð¸ =RedisCheckConfigFile() Ð² Excel) Ð¸ ÑÐ¼. Ñ€Ð°Ð·Ð´ÐµÐ» README "Migrating from the old shape".

; ---- french ----
french.AdvPageDesc=Voici tout votre RedisExcel.json, prÃªt Ã  modifier : chaque clÃ© est prÃ©remplie avec sa valeur par dÃ©faut et un court commentaire. Le complÃ©ment accepte les commentaires //, vous pouvez donc les laisser. Modifiez les hÃ´tes, alias, mode d'Ã©criture, ConflationMs, AsyncWrites - tout est ici.
french.AdvPageTitle=Connexion avancÃ©e
french.ConfigWriteFailed=Impossible d'Ã©crire %1. CrÃ©ez-le manuellement Ã  partir de l'exemple (RedisExcel.json).
french.ConnHost=HÃ´te (ou hÃ´te:port) :
french.ConnPageDesc=Cliquez sur Suivant pour la valeur par dÃ©faut (localhost:6379). Vous pourrez la modifier plus tard en Ã©ditant RedisExcel.json.
french.ConnPageSubtitle=OÃ¹ RedisExcel doit-il se connecter ?
french.ConnPageTitle=Connexion Redis
french.ConnPassword=Mot de passe (facultatif) :
french.ConnPort=Port :
french.FinishedText=RedisExcel est installÃ© dans :%n%2%n%nSi le complÃ©ment ne se charge pas automatiquement, activez-le dans Excel :%nFichier > Options > ComplÃ©ments > Atteindre... > Parcourir... et choisissez RedisExcel64-packed.xll (ou RedisExcel-packed.xll pour Office 32 bits).%n%nEssayez ensuite =RedisUDFGet("preco_btc") dans une cellule.
french.GroupConfig=Configuration :
french.GroupExcel=Excel :
french.HostRequired=Veuillez saisir un hÃ´te Redis (par exemple localhost).
french.IconUninstall=DÃ©sinstaller {#AppName}
french.ModeAdvanced=AvancÃ© - dÃ©finir la base de donnÃ©es, SSL, le mode d'Ã©criture et le rÃ©glage des mises Ã  jour
french.ModePageDesc=Simple ne demande que l'adresse du serveur. AvancÃ© permet de dÃ©finir la base de donnÃ©es, SSL/TLS, le mode d'Ã©criture et le rÃ©glage des mises Ã  jour en temps rÃ©el. Vous pourrez toujours modifier RedisExcel.json plus tard (ou exÃ©cuter Validate-Config.cmd).
french.ModePageSubtitle=Installation simple ou avancÃ©e ?
french.ModePageTitle=Mode d'installation
french.ModeSimple=Simple - juste l'adresse Redis (recommandÃ©)
french.OptionsOpenExcel=Ouvrir un classeur RedisExcel d'exemple Ã  la fin
french.OptionsPageDesc=L'installateur teste la connexion Ã  la fin. Ouvrez Ã©ventuellement un classeur Excel d'exemple montrant son fonctionnement.
french.OptionsPageSubtitle=Presque terminÃ©
french.OptionsPageTitle=VÃ©rifications supplÃ©mentaires
french.ServerDockerRedis=Oui - Docker : Redis (redis:latest)
french.ServerDockerValkey=Oui - Docker : Valkey (valkey/valkey:latest)
french.ServerFailed=Impossible de configurer %1 automatiquement (code %2).%nVous pouvez installer un serveur plus tard et le pointer dans RedisExcel.json.
french.ServerMemurai=Oui - Memurai (service Windows natif, admin requis)
french.ServerNameDockerRedis=Redis (Docker)
french.ServerNameDockerValkey=Valkey (Docker)
french.ServerNameGeneric=le serveur
french.ServerNameMemurai=Memurai
french.ServerNameWslRedis=Redis (WSL)
french.ServerNameWslValkey=Valkey (WSL)
french.ServerNone=Non - j'ai dÃ©jÃ  un serveur (ou je le configurerai plus tard)
french.ServerPageDesc=RedisExcel a besoin d'un serveur Redis (ou Valkey). Choisissez une option si vous n'en avez pas encore. Le serveur est installÃ© d'abord, puis RedisExcel, et la connexion est testÃ©e Ã  la fin. Docker est le plus simple.
french.ServerPageSubtitle=Voulez-vous installer un serveur compatible Redis maintenant ?
french.ServerPageTitle=Serveur Redis
french.ServerWslRedis=Oui - WSL2 : Redis (apt install redis-server)
french.ServerWslValkey=Oui - WSL2 : Valkey (apt install valkey-server)
french.SummaryAdvanced=AvancÃ©
french.SummaryIntro=Ces valeurs seront enregistrÃ©es dans RedisExcel.json, dans le dossier d'installation.
french.SummaryLabel=Connexion
french.SummaryMode=Mode
french.SummaryPageSubtitle=VÃ©rifiez la connexion qui sera Ã©crite dans RedisExcel.json
french.SummaryPageTitle=Vos paramÃ¨tres
french.SummarySimple=Simple
french.TaskAutoLoadDesc=Enregistrer le complÃ©ment pour qu'Excel le charge automatiquement (recommandÃ©)
french.TaskConfigDesc=CrÃ©er RedisExcel.json dans le dossier d'installation (paramÃ¨tres de l'hÃ´te)
french.TaskNlogDesc=CrÃ©er NLog.config (Ã©crire un journal Ã  cÃ´tÃ© du complÃ©ment)
french.TestFailed=InstallÃ©, mais le Redis Ã  %1 n'a pas rÃ©pondu Ã  un PING.%nVÃ©rifiez que le serveur est dÃ©marrÃ© (docker ps) et ajustez RedisExcel.json s'il est sur un autre hÃ´te/port.
french.TestOk=Tout est prÃªt ! Redis a rÃ©pondu Ã  %1.%n%nPour l'utiliser : ouvrez Excel. Si le complÃ©ment ne se charge pas automatiquement, allez dans Fichier > Options > ComplÃ©ments > Atteindre... > Parcourir... et choisissez le .xll dans %2.
french.UninstKeepConfig=Conserver vos paramÃ¨tres RedisExcel.json ?
french.UninstServerFound=%1  [trouvÃ©]
french.UninstServerFailed=Impossible de supprimer %1 automatiquement (code %2). Supprimez-le manuellement si besoin.
french.UninstServersDesc=Ces serveurs ont Ã©tÃ© trouvÃ©s sur cette machine. Cochez ceux Ã  supprimer avec RedisExcel.%nPar dÃ©faut, RIEN n'est supprimÃ©.
french.UninstServersSubtitle=Supprimer aussi les serveurs compatibles Redis ?
french.UninstServersTitle=Serveurs
french.UpgradeConfigKept=Un RedisExcel.json existant a Ã©tÃ© trouvÃ© et conservÃ© (vos paramÃ¨tres n'ont pas Ã©tÃ© Ã©crasÃ©s).%n%nIl s'agit d'une MISE Ã€ JOUR MAJEURE : le format de configuration a changÃ©. Si ce fichier utilise l'ancien format, chaque cellule affichera une erreur #CONFIG jusqu'Ã  sa migration.%n%nExÃ©cutez Validate-Config.cmd (ou =RedisCheckConfigFile() dans Excel) et consultez la section README "Migrating from the old shape".

; ---- korean ----
korean.AdvPageDesc=ì´ê²ƒì€ íŽ¸ì§‘í•  ìˆ˜ ìžˆëŠ” ì „ì²´ RedisExcel.jsonìž…ë‹ˆë‹¤. ê° í‚¤ì—ëŠ” ê¸°ë³¸ê°’ê³¼ ì§§ì€ ì„¤ëª…ì´ ë¯¸ë¦¬ ì±„ì›Œì ¸ ìžˆìŠµë‹ˆë‹¤. ì¶”ê°€ ê¸°ëŠ¥ì€ // ì£¼ì„ì„ í—ˆìš©í•˜ë¯€ë¡œ ê·¸ëŒ€ë¡œ ë‘ì–´ë„ ë©ë‹ˆë‹¤. í˜¸ìŠ¤íŠ¸, ë³„ì¹­, ì“°ê¸° ëª¨ë“œ, ConflationMs, AsyncWritesë¥¼ ì—¬ê¸°ì„œ ë³€ê²½í•˜ì„¸ìš”.
korean.AdvPageTitle=ê³ ê¸‰ ì—°ê²°
korean.ConfigWriteFailed=%1ì„(ë¥¼) ì“¸ ìˆ˜ ì—†ìŠµë‹ˆë‹¤. ì˜ˆì œ(RedisExcel.json)ë¥¼ ì°¸ê³ í•´ ìˆ˜ë™ìœ¼ë¡œ ë§Œë“œì„¸ìš”.
korean.ConnHost=í˜¸ìŠ¤íŠ¸(ë˜ëŠ” host:port):
korean.ConnPageDesc=ê¸°ë³¸ê°’(localhost:6379)ì´ë©´ ë‹¤ìŒì„ ëˆ„ë¥´ì„¸ìš”. ë‚˜ì¤‘ì— RedisExcel.jsonì„ íŽ¸ì§‘í•´ ë³€ê²½í•  ìˆ˜ ìžˆìŠµë‹ˆë‹¤.
korean.ConnPageSubtitle=RedisExcelì´ ì–´ë””ì— ì—°ê²°í•´ì•¼ í•©ë‹ˆê¹Œ?
korean.ConnPageTitle=Redis ì—°ê²°
korean.ConnPassword=ë¹„ë°€ë²ˆí˜¸(ì„ íƒ):
korean.ConnPort=í¬íŠ¸:
korean.FinishedText=RedisExcelì´ ë‹¤ìŒ ìœ„ì¹˜ì— ì„¤ì¹˜ë˜ì—ˆìŠµë‹ˆë‹¤:%n%2%n%nì¶”ê°€ ê¸°ëŠ¥ì´ ìžë™ìœ¼ë¡œ ë¡œë“œë˜ì§€ ì•Šìœ¼ë©´ Excelì—ì„œ í™œì„±í™”í•˜ì„¸ìš”:%níŒŒì¼ > ì˜µì…˜ > ì¶”ê°€ ê¸°ëŠ¥ > ì´ë™... > ì°¾ì•„ë³´ê¸°...ì—ì„œ RedisExcel64-packed.xll(32ë¹„íŠ¸ OfficeëŠ” RedisExcel-packed.xll)ì„ ì„ íƒí•©ë‹ˆë‹¤.%n%nê·¸ëŸ° ë‹¤ìŒ ì…€ì— =RedisUDFGet("preco_btc")ë¥¼ ìž…ë ¥í•´ ë³´ì„¸ìš”.
korean.GroupConfig=êµ¬ì„±:
korean.GroupExcel=Excel:
korean.HostRequired=Redis í˜¸ìŠ¤íŠ¸ë¥¼ ìž…ë ¥í•˜ì„¸ìš”(ì˜ˆ: localhost).
korean.IconUninstall={#AppName} ì œê±°
korean.ModeAdvanced=ê³ ê¸‰ - ë°ì´í„°ë² ì´ìŠ¤, SSL, ì“°ê¸° ëª¨ë“œ ë° ì—…ë°ì´íŠ¸ íŠœë‹ ì„¤ì •
korean.ModePageDesc=ê°„ë‹¨ì€ ì„œë²„ ì£¼ì†Œë§Œ ë¬»ìŠµë‹ˆë‹¤. ê³ ê¸‰ì€ ë°ì´í„°ë² ì´ìŠ¤, SSL/TLS, ì“°ê¸° ëª¨ë“œ, ì‹¤ì‹œê°„ ì—…ë°ì´íŠ¸ íŠœë‹ì„ ì„¤ì •í•  ìˆ˜ ìžˆìŠµë‹ˆë‹¤. RedisExcel.jsonì€ ë‚˜ì¤‘ì— ì–¸ì œë“ ì§€ íŽ¸ì§‘í•  ìˆ˜ ìžˆìŠµë‹ˆë‹¤(Validate-Config.cmd ì‹¤í–‰ ê°€ëŠ¥).
korean.ModePageSubtitle=ê°„ë‹¨ ë˜ëŠ” ê³ ê¸‰ ì„¤ì¹˜?
korean.ModePageTitle=ì„¤ì¹˜ ëª¨ë“œ
korean.ModeSimple=ê°„ë‹¨ - Redis ì£¼ì†Œë§Œ(ê¶Œìž¥)
korean.OptionsOpenExcel=ë§ˆì§€ë§‰ì— ìƒ˜í”Œ RedisExcel í†µí•© ë¬¸ì„œ ì—´ê¸°
korean.OptionsPageDesc=ì„¤ì¹˜ í”„ë¡œê·¸ëž¨ì´ ëë‚  ë•Œ ì—°ê²°ì„ í…ŒìŠ¤íŠ¸í•©ë‹ˆë‹¤. ì›í•˜ë©´ ìž‘ë™ì„ ë³´ì—¬ ì£¼ëŠ” ìƒ˜í”Œ Excel í†µí•© ë¬¸ì„œë¥¼ ì—½ë‹ˆë‹¤.
korean.OptionsPageSubtitle=ê±°ì˜ ì™„ë£Œ
korean.OptionsPageTitle=ì¶”ê°€ í™•ì¸
korean.ServerDockerRedis=ì˜ˆ - Docker: Redis(redis:latest)
korean.ServerDockerValkey=ì˜ˆ - Docker: Valkey(valkey/valkey:latest)
korean.ServerFailed=%1ì„(ë¥¼) ìžë™ìœ¼ë¡œ ì„¤ì •í•˜ì§€ ëª»í–ˆìŠµë‹ˆë‹¤(ì¢…ë£Œ ì½”ë“œ %2).%në‚˜ì¤‘ì— ì„œë²„ë¥¼ ì„¤ì¹˜í•˜ê³  RedisExcel.jsonì—ì„œ ì§€ì •í•  ìˆ˜ ìžˆìŠµë‹ˆë‹¤.
korean.ServerMemurai=ì˜ˆ - Memurai(ë„¤ì´í‹°ë¸Œ Windows ì„œë¹„ìŠ¤, ê´€ë¦¬ìž í•„ìš”)
korean.ServerNameDockerRedis=Redis(Docker)
korean.ServerNameDockerValkey=Valkey(Docker)
korean.ServerNameGeneric=ì„œë²„
korean.ServerNameMemurai=Memurai
korean.ServerNameWslRedis=Redis(WSL)
korean.ServerNameWslValkey=Valkey(WSL)
korean.ServerNone=ì•„ë‹ˆìš” - ì´ë¯¸ ì„œë²„ê°€ ìžˆìŠµë‹ˆë‹¤(ë˜ëŠ” ë‚˜ì¤‘ì— ì„¤ì •)
korean.ServerPageDesc=RedisExcelì—ëŠ” Redis(ë˜ëŠ” Valkey) ì„œë²„ê°€ í•„ìš”í•©ë‹ˆë‹¤. ì•„ì§ ì—†ìœ¼ë©´ ì˜µì…˜ì„ ì„ íƒí•˜ì„¸ìš”. ì„œë²„ë¥¼ ë¨¼ì € ì„¤ì¹˜í•œ ë‹¤ìŒ RedisExcelì„ ì„¤ì¹˜í•˜ê³  ë§ˆì§€ë§‰ì— ì—°ê²°ì„ í…ŒìŠ¤íŠ¸í•©ë‹ˆë‹¤. Dockerê°€ ê°€ìž¥ ê°„ë‹¨í•©ë‹ˆë‹¤.
korean.ServerPageSubtitle=ì§€ê¸ˆ Redis í˜¸í™˜ ì„œë²„ë¥¼ ì„¤ì¹˜í•˜ì‹œê² ìŠµë‹ˆê¹Œ?
korean.ServerPageTitle=Redis ì„œë²„
korean.ServerWslRedis=ì˜ˆ - WSL2: Redis(apt install redis-server)
korean.ServerWslValkey=ì˜ˆ - WSL2: Valkey(apt install valkey-server)
korean.SummaryAdvanced=ê³ ê¸‰
korean.SummaryIntro=ì´ ê°’ë“¤ì€ ì„¤ì¹˜ í´ë”ì˜ RedisExcel.jsonì— ì €ìž¥ë©ë‹ˆë‹¤.
korean.SummaryLabel=ì—°ê²°
korean.SummaryMode=ëª¨ë“œ
korean.SummaryPageSubtitle=RedisExcel.jsonì— ê¸°ë¡ë  ì—°ê²°ì„ í™•ì¸í•˜ì„¸ìš”
korean.SummaryPageTitle=ì„¤ì •
korean.SummarySimple=ê°„ë‹¨
korean.TaskAutoLoadDesc=Excelì´ ìžë™ìœ¼ë¡œ ë¡œë“œí•˜ë„ë¡ ì¶”ê°€ ê¸°ëŠ¥ ë“±ë¡(ê¶Œìž¥)
korean.TaskConfigDesc=ì„¤ì¹˜ í´ë”ì— RedisExcel.json ìƒì„±(í˜¸ìŠ¤íŠ¸ ì„¤ì •)
korean.TaskNlogDesc=NLog.config ìƒì„±(ì¶”ê°€ ê¸°ëŠ¥ ì˜†ì— ë¡œê·¸ íŒŒì¼ ê¸°ë¡)
korean.TestFailed=ì„¤ì¹˜ë˜ì—ˆì§€ë§Œ %1ì˜ Redisê°€ PINGì— ì‘ë‹µí•˜ì§€ ì•Šì•˜ìŠµë‹ˆë‹¤.%nì„œë²„ê°€ ì‹¤í–‰ ì¤‘ì¸ì§€ í™•ì¸í•˜ê³ (docker ps) ë‹¤ë¥¸ í˜¸ìŠ¤íŠ¸/í¬íŠ¸ë¼ë©´ RedisExcel.jsonì„ ì¡°ì •í•˜ì„¸ìš”.
korean.TestOk=ì¤€ë¹„ ì™„ë£Œ! %1ì˜ Redisê°€ ì‘ë‹µí–ˆìŠµë‹ˆë‹¤.%n%nì‚¬ìš©ë²•: Excelì„ ì—¬ì„¸ìš”. ì¶”ê°€ ê¸°ëŠ¥ì´ ìžë™ìœ¼ë¡œ ë¡œë“œë˜ì§€ ì•Šìœ¼ë©´ íŒŒì¼ > ì˜µì…˜ > ì¶”ê°€ ê¸°ëŠ¥ > ì´ë™... > ì°¾ì•„ë³´ê¸°...ì—ì„œ %2ì˜ .xllì„ ì„ íƒí•˜ì„¸ìš”.
korean.UninstKeepConfig=RedisExcel.json ì„¤ì •ì„ ìœ ì§€í•˜ì‹œê² ìŠµë‹ˆê¹Œ?
korean.UninstServerFound=%1  [ë°œê²¬]
korean.UninstServerFailed=%1을(를) 자동으로 제거하지 못했습니다(종료 코드 %2). 필요하면 수동으로 제거하세요.
korean.UninstServersDesc=ì´ ì„œë²„ë“¤ì´ ì´ ì»´í“¨í„°ì—ì„œ ë°œê²¬ë˜ì—ˆìŠµë‹ˆë‹¤. RedisExcelê³¼ í•¨ê»˜ ì œê±°í•  í•­ëª©ì„ ì„ íƒí•˜ì„¸ìš”.%nê¸°ë³¸ì ìœ¼ë¡œ ì•„ë¬´ê²ƒë„ ì œê±°ë˜ì§€ ì•ŠìŠµë‹ˆë‹¤.
korean.UninstServersSubtitle=Redis í˜¸í™˜ ì„œë²„ë„ ì œê±°í•˜ì‹œê² ìŠµë‹ˆê¹Œ?
korean.UninstServersTitle=ì„œë²„
korean.UpgradeConfigKept=ê¸°ì¡´ RedisExcel.jsonì„ ì°¾ì•„ ìœ ì§€í–ˆìŠµë‹ˆë‹¤(ì„¤ì •ì´ ë®ì–´ì“°ì´ì§€ ì•Šì•˜ìŠµë‹ˆë‹¤).%n%nì´ê²ƒì€ ì£¼ìš” ì—…ë°ì´íŠ¸ìž…ë‹ˆë‹¤: êµ¬ì„± í˜•ì‹ì´ ë³€ê²½ë˜ì—ˆìŠµë‹ˆë‹¤. í•´ë‹¹ íŒŒì¼ì´ ì´ì „ í˜•ì‹ì´ë©´ ë§ˆì´ê·¸ë ˆì´ì…˜í•  ë•Œê¹Œì§€ ëª¨ë“  ì…€ì— #CONFIG ì˜¤ë¥˜ê°€ í‘œì‹œë©ë‹ˆë‹¤.%n%nValidate-Config.cmd(ë˜ëŠ” Excelì—ì„œ =RedisCheckConfigFile())ë¥¼ ì‹¤í–‰í•˜ê³  READMEì˜ "Migrating from the old shape" ì„¹ì…˜ì„ ì°¸ê³ í•˜ì„¸ìš”.

; ---- german ----
german.AdvPageDesc=Dies ist Ihre gesamte RedisExcel.json, bereit zum Bearbeiten: jeder SchlÃ¼ssel ist mit seinem Standardwert und einem kurzen Kommentar vorbelegt. Das Add-in akzeptiert //-Kommentare, Sie kÃ¶nnen sie also stehen lassen. Ã„ndern Sie Hosts, Aliase, Schreibmodus, ConflationMs, AsyncWrites - alles hier.
german.AdvPageTitle=Erweiterte Verbindung
german.ConfigWriteFailed=%1 konnte nicht geschrieben werden. Erstellen Sie es manuell nach dem Beispiel (RedisExcel.json).
german.ConnHost=Host (oder host:port):
german.ConnPageDesc=Klicken Sie auf Weiter fÃ¼r den Standard (localhost:6379). Sie kÃ¶nnen dies spÃ¤ter in RedisExcel.json Ã¤ndern.
german.ConnPageSubtitle=Womit soll sich RedisExcel verbinden?
german.ConnPageTitle=Redis-Verbindung
german.ConnPassword=Passwort (optional):
german.ConnPort=Port:
german.FinishedText=RedisExcel ist installiert in:%n%2%n%nWenn das Add-in nicht automatisch geladen wird, aktivieren Sie es in Excel:%nDatei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wÃ¤hlen Sie RedisExcel64-packed.xll (oder RedisExcel-packed.xll fÃ¼r 32-Bit-Office).%n%nProbieren Sie dann =RedisUDFGet("preco_btc") in einer Zelle.
german.GroupConfig=Konfiguration:
german.GroupExcel=Excel:
german.HostRequired=Bitte geben Sie einen Redis-Host ein (z. B. localhost).
german.IconUninstall={#AppName} deinstallieren
german.ModeAdvanced=Erweitert - Datenbank, SSL, Schreibmodus und Update-Tuning festlegen
german.ModePageDesc=Einfach fragt nur nach der Serveradresse. Erweitert erlaubt Datenbank, SSL/TLS, den Schreibmodus und das Echtzeit-Update-Tuning. Sie kÃ¶nnen RedisExcel.json spÃ¤ter jederzeit bearbeiten (oder Validate-Config.cmd ausfÃ¼hren).
german.ModePageSubtitle=Einfache oder erweiterte Einrichtung?
german.ModePageTitle=Einrichtungsmodus
german.ModeSimple=Einfach - nur die Redis-Adresse (empfohlen)
german.OptionsOpenExcel=Am Ende eine Beispiel-Arbeitsmappe RedisExcel Ã¶ffnen
german.OptionsPageDesc=Der Installer testet die Verbindung am Ende. Optional wird eine Beispiel-Excel-Arbeitsmappe geÃ¶ffnet, die die Funktion zeigt.
german.OptionsPageSubtitle=Fast fertig
german.OptionsPageTitle=ZusÃ¤tzliche PrÃ¼fungen
german.ServerDockerRedis=Ja - Docker: Redis (redis:latest)
german.ServerDockerValkey=Ja - Docker: Valkey (valkey/valkey:latest)
german.ServerFailed=%1 konnte nicht automatisch eingerichtet werden (Code %2).%nSie kÃ¶nnen spÃ¤ter einen Server installieren und RedisExcel.json darauf verweisen.
german.ServerMemurai=Ja - Memurai (nativer Windows-Dienst, Admin erforderlich)
german.ServerNameDockerRedis=Redis (Docker)
german.ServerNameDockerValkey=Valkey (Docker)
german.ServerNameGeneric=der Server
german.ServerNameMemurai=Memurai
german.ServerNameWslRedis=Redis (WSL)
german.ServerNameWslValkey=Valkey (WSL)
german.ServerNone=Nein - ich habe bereits einen Server (oder richte ihn spÃ¤ter ein)
german.ServerPageDesc=RedisExcel benÃ¶tigt einen Redis- (oder Valkey-)Server. WÃ¤hlen Sie eine Option, wenn Sie noch keinen haben. Zuerst wird der Server installiert, dann RedisExcel, und am Ende wird die Verbindung getestet. Docker ist am einfachsten.
german.ServerPageSubtitle=MÃ¶chten Sie jetzt einen Redis-kompatiblen Server installieren?
german.ServerPageTitle=Redis-Server
german.ServerWslRedis=Ja - WSL2: Redis (apt install redis-server)
german.ServerWslValkey=Ja - WSL2: Valkey (apt install valkey-server)
german.SummaryAdvanced=Erweitert
german.SummaryIntro=Diese Werte werden in RedisExcel.json im Installationsordner gespeichert.
german.SummaryLabel=Verbindung
german.SummaryMode=Modus
german.SummaryPageSubtitle=PrÃ¼fen Sie die Verbindung, die in RedisExcel.json geschrieben wird
german.SummaryPageTitle=Ihre Einstellungen
german.SummarySimple=Einfach
german.TaskAutoLoadDesc=Das Add-in registrieren, damit Excel es automatisch lÃ¤dt (empfohlen)
german.TaskConfigDesc=RedisExcel.json im Installationsordner erstellen (Host-Einstellungen)
german.TaskNlogDesc=NLog.config erstellen (Logdatei neben dem Add-in schreiben)
german.TestFailed=Installiert, aber der Redis unter %1 hat nicht auf einen PING geantwortet.%nPrÃ¼fen Sie, ob der Server lÃ¤uft (docker ps), und passen Sie RedisExcel.json an, wenn er auf einem anderen Host/Port ist.
german.TestOk=Alles bereit! Redis hat unter %1 geantwortet.%n%nSo verwenden Sie es: Excel Ã¶ffnen. Wenn das Add-in nicht automatisch geladen wird, gehen Sie zu Datei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wÃ¤hlen Sie die .xll in %2.
german.UninstKeepConfig=Ihre RedisExcel.json-Einstellungen behalten?
german.UninstServerFound=%1  [gefunden]
german.UninstServerFailed=%1 konnte nicht automatisch entfernt werden (Code %2). Entfernen Sie es bei Bedarf manuell.
german.UninstServersDesc=Diese Server wurden auf diesem Computer gefunden. Markieren Sie die, die mit RedisExcel entfernt werden sollen.%nStandardmÃ¤ÃŸig wird NICHTS entfernt.
german.UninstServersSubtitle=Auch Redis-kompatible Server entfernen?
german.UninstServersTitle=Server
german.UpgradeConfigKept=Eine vorhandene RedisExcel.json wurde gefunden und behalten (Ihre Einstellungen wurden nicht Ã¼berschrieben).%n%nDies ist ein GROSSES Update: Das Konfigurationsformat hat sich geÃ¤ndert. Wenn diese Datei das alte Format verwendet, zeigt jede Zelle einen #CONFIG-Fehler, bis sie migriert ist.%n%nFÃ¼hren Sie Validate-Config.cmd (oder =RedisCheckConfigFile() in Excel) aus und lesen Sie den README-Abschnitt "Migrating from the old shape".

; ---- japanese ----
japanese.AdvPageDesc=ã“ã‚Œã¯ RedisExcel.json ã®å…¨ä½“ã§ã™ã€‚å„ã‚­ãƒ¼ã«ã¯æ—¢å®šå€¤ã¨çŸ­ã„ã‚³ãƒ¡ãƒ³ãƒˆãŒã‚ã‚‰ã‹ã˜ã‚å…¥åŠ›ã•ã‚Œã¦ã„ã¾ã™ã€‚ã‚¢ãƒ‰ã‚¤ãƒ³ã¯ // ã‚³ãƒ¡ãƒ³ãƒˆã‚’å—ã‘ä»˜ã‘ã‚‹ã®ã§ãã®ã¾ã¾æ®‹ã›ã¾ã™ã€‚ãƒ›ã‚¹ãƒˆã€ã‚¨ã‚¤ãƒªã‚¢ã‚¹ã€æ›¸ãè¾¼ã¿ãƒ¢ãƒ¼ãƒ‰ã€ConflationMsã€AsyncWrites ã‚’ã“ã“ã§å¤‰æ›´ã§ãã¾ã™ã€‚
japanese.AdvPageTitle=è©³ç´°æŽ¥ç¶š
japanese.ConfigWriteFailed=%1 ã«æ›¸ãè¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚ã‚µãƒ³ãƒ—ãƒ«ï¼ˆRedisExcel.jsonï¼‰ã‹ã‚‰æ‰‹å‹•ã§ä½œæˆã—ã¦ãã ã•ã„ã€‚
japanese.ConnHost=ãƒ›ã‚¹ãƒˆï¼ˆã¾ãŸã¯ host:portï¼‰:
japanese.ConnPageDesc=æ—¢å®šå€¤ï¼ˆlocalhost:6379ï¼‰ã§ã‚ˆã‘ã‚Œã°ã€Œæ¬¡ã¸ã€ã€‚å¾Œã§ RedisExcel.json ã‚’ç·¨é›†ã—ã¦å¤‰æ›´ã§ãã¾ã™ã€‚
japanese.ConnPageSubtitle=RedisExcel ã¯ã©ã“ã«æŽ¥ç¶šã—ã¾ã™ã‹ï¼Ÿ
japanese.ConnPageTitle=Redis æŽ¥ç¶š
japanese.ConnPassword=ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ï¼ˆä»»æ„ï¼‰:
japanese.ConnPort=ãƒãƒ¼ãƒˆ:
japanese.FinishedText=RedisExcel ã¯æ¬¡ã«ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã•ã‚Œã¾ã—ãŸ:%n%2%n%nã‚¢ãƒ‰ã‚¤ãƒ³ãŒè‡ªå‹•ã§èª­ã¿è¾¼ã¾ã‚Œãªã„å ´åˆã¯ Excel ã§æœ‰åŠ¹åŒ–ã—ã¦ãã ã•ã„:%nãƒ•ã‚¡ã‚¤ãƒ« > ã‚ªãƒ—ã‚·ãƒ§ãƒ³ > ã‚¢ãƒ‰ã‚¤ãƒ³ > è¨­å®š... > å‚ç…§... ã§ RedisExcel64-packed.xllï¼ˆ32 ãƒ“ãƒƒãƒˆ Office ãªã‚‰ RedisExcel-packed.xllï¼‰ã‚’é¸æŠžã—ã¾ã™ã€‚%n%nãã®å¾Œã€ã‚»ãƒ«ã« =RedisUDFGet("preco_btc") ã‚’å…¥åŠ›ã—ã¦ã¿ã¦ãã ã•ã„ã€‚
japanese.GroupConfig=æ§‹æˆ:
japanese.GroupExcel=Excel:
japanese.HostRequired=Redis ãƒ›ã‚¹ãƒˆã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ï¼ˆä¾‹: localhostï¼‰ã€‚
japanese.IconUninstall={#AppName} ã‚’ã‚¢ãƒ³ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«
japanese.ModeAdvanced=è©³ç´° - ãƒ‡ãƒ¼ã‚¿ãƒ™ãƒ¼ã‚¹ã€SSLã€æ›¸ãè¾¼ã¿ãƒ¢ãƒ¼ãƒ‰ã€æ›´æ–°ãƒãƒ¥ãƒ¼ãƒ‹ãƒ³ã‚°ã‚’è¨­å®š
japanese.ModePageDesc=ç°¡æ˜“ã¯ã‚µãƒ¼ãƒãƒ¼ã‚¢ãƒ‰ãƒ¬ã‚¹ã®ã¿å°‹ã­ã¾ã™ã€‚è©³ç´°ã§ã¯ãƒ‡ãƒ¼ã‚¿ãƒ™ãƒ¼ã‚¹ã€SSL/TLSã€æ›¸ãè¾¼ã¿ãƒ¢ãƒ¼ãƒ‰ã€ãƒªã‚¢ãƒ«ã‚¿ã‚¤ãƒ æ›´æ–°ãƒãƒ¥ãƒ¼ãƒ‹ãƒ³ã‚°ã‚’è¨­å®šã§ãã¾ã™ã€‚RedisExcel.json ã¯å¾Œã§ã„ã¤ã§ã‚‚ç·¨é›†ã§ãã¾ã™ï¼ˆValidate-Config.cmd ã®å®Ÿè¡Œã§ã‚‚å¯ï¼‰ã€‚
japanese.ModePageSubtitle=ç°¡æ˜“ã¾ãŸã¯è©³ç´°ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—ï¼Ÿ
japanese.ModePageTitle=ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—ãƒ¢ãƒ¼ãƒ‰
japanese.ModeSimple=ç°¡æ˜“ - Redis ã‚¢ãƒ‰ãƒ¬ã‚¹ã®ã¿ï¼ˆæŽ¨å¥¨ï¼‰
japanese.OptionsOpenExcel=æœ€å¾Œã«ã‚µãƒ³ãƒ—ãƒ«ã® RedisExcel ãƒ–ãƒƒã‚¯ã‚’é–‹ã
japanese.OptionsPageDesc=ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ©ãƒ¼ã¯çµ‚äº†æ™‚ã«æŽ¥ç¶šã‚’ãƒ†ã‚¹ãƒˆã—ã¾ã™ã€‚å¿…è¦ã«å¿œã˜ã¦ã€å‹•ä½œã‚’ç¤ºã™ã‚µãƒ³ãƒ—ãƒ« Excel ãƒ–ãƒƒã‚¯ã‚’é–‹ãã¾ã™ã€‚
japanese.OptionsPageSubtitle=ã¾ã‚‚ãªãå®Œäº†
japanese.OptionsPageTitle=è¿½åŠ ãƒã‚§ãƒƒã‚¯
japanese.ServerDockerRedis=ã¯ã„ - Docker: Redisï¼ˆredis:latestï¼‰
japanese.ServerDockerValkey=ã¯ã„ - Docker: Valkeyï¼ˆvalkey/valkey:latestï¼‰
japanese.ServerFailed=%1 ã‚’è‡ªå‹•ã§ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—ã§ãã¾ã›ã‚“ã§ã—ãŸï¼ˆçµ‚äº†ã‚³ãƒ¼ãƒ‰ %2ï¼‰ã€‚%nå¾Œã§ã‚µãƒ¼ãƒãƒ¼ã‚’ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã—ã€RedisExcel.json ã§æŒ‡å®šã§ãã¾ã™ã€‚
japanese.ServerMemurai=ã¯ã„ - Memuraiï¼ˆWindows ãƒã‚¤ãƒ†ã‚£ãƒ– ã‚µãƒ¼ãƒ“ã‚¹ã€ç®¡ç†è€…æ¨©é™ãŒå¿…è¦ï¼‰
japanese.ServerNameDockerRedis=Redisï¼ˆDockerï¼‰
japanese.ServerNameDockerValkey=Valkeyï¼ˆDockerï¼‰
japanese.ServerNameGeneric=ã‚µãƒ¼ãƒãƒ¼
japanese.ServerNameMemurai=Memurai
japanese.ServerNameWslRedis=Redisï¼ˆWSLï¼‰
japanese.ServerNameWslValkey=Valkeyï¼ˆWSLï¼‰
japanese.ServerNone=ã„ã„ãˆ - æ—¢ã«ã‚µãƒ¼ãƒãƒ¼ãŒã‚ã‚‹ï¼ˆã¾ãŸã¯å¾Œã§ç”¨æ„ã™ã‚‹ï¼‰
japanese.ServerPageDesc=RedisExcel ã«ã¯ Redisï¼ˆã¾ãŸã¯ Valkeyï¼‰ã‚µãƒ¼ãƒãƒ¼ãŒå¿…è¦ã§ã™ã€‚ã¾ã ç„¡ã„å ´åˆã¯ã‚ªãƒ—ã‚·ãƒ§ãƒ³ã‚’é¸æŠžã—ã¦ãã ã•ã„ã€‚å…ˆã«ã‚µãƒ¼ãƒãƒ¼ã‚’ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã—ã€æ¬¡ã« RedisExcel ã‚’ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã—ã€æœ€å¾Œã«æŽ¥ç¶šã‚’ãƒ†ã‚¹ãƒˆã—ã¾ã™ã€‚Docker ãŒæœ€ã‚‚ç°¡å˜ã§ã™ã€‚
japanese.ServerPageSubtitle=ä»Šã€Redis äº’æ›ã‚µãƒ¼ãƒãƒ¼ã‚’ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã—ã¾ã™ã‹ï¼Ÿ
japanese.ServerPageTitle=Redis ã‚µãƒ¼ãƒãƒ¼
japanese.ServerWslRedis=ã¯ã„ - WSL2: Redisï¼ˆapt install redis-serverï¼‰
japanese.ServerWslValkey=ã¯ã„ - WSL2: Valkeyï¼ˆapt install valkey-serverï¼‰
japanese.SummaryAdvanced=è©³ç´°
japanese.SummaryIntro=ã“ã‚Œã‚‰ã®å€¤ã¯ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã® RedisExcel.json ã«ä¿å­˜ã•ã‚Œã¾ã™ã€‚
japanese.SummaryLabel=æŽ¥ç¶š
japanese.SummaryMode=ãƒ¢ãƒ¼ãƒ‰
japanese.SummaryPageSubtitle=RedisExcel.json ã«æ›¸ãè¾¼ã¾ã‚Œã‚‹æŽ¥ç¶šã‚’ç¢ºèªã—ã¦ãã ã•ã„
japanese.SummaryPageTitle=è¨­å®šå†…å®¹
japanese.SummarySimple=ç°¡æ˜“
japanese.TaskAutoLoadDesc=Excel ãŒè‡ªå‹•ã§èª­ã¿è¾¼ã‚€ã‚ˆã†ã‚¢ãƒ‰ã‚¤ãƒ³ã‚’ç™»éŒ²ï¼ˆæŽ¨å¥¨ï¼‰
japanese.TaskConfigDesc=ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã« RedisExcel.json ã‚’ä½œæˆï¼ˆãƒ›ã‚¹ãƒˆè¨­å®šï¼‰
japanese.TaskNlogDesc=NLog.config ã‚’ä½œæˆï¼ˆã‚¢ãƒ‰ã‚¤ãƒ³ã®éš£ã«ãƒ­ã‚° ãƒ•ã‚¡ã‚¤ãƒ«ã‚’æ›¸ãè¾¼ã‚€ï¼‰
japanese.TestFailed=ã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«æ¸ˆã¿ã§ã™ãŒã€%1 ã® Redis ãŒ PING ã«å¿œç­”ã—ã¾ã›ã‚“ã§ã—ãŸã€‚%nã‚µãƒ¼ãƒãƒ¼ãŒèµ·å‹•ã—ã¦ã„ã‚‹ã‹ç¢ºèªã—ï¼ˆdocker psï¼‰ã€åˆ¥ã®ãƒ›ã‚¹ãƒˆ/ãƒãƒ¼ãƒˆãªã‚‰ RedisExcel.json ã‚’èª¿æ•´ã—ã¦ãã ã•ã„ã€‚
japanese.TestOk=æº–å‚™å®Œäº†ï¼ %1 ã® Redis ãŒå¿œç­”ã—ã¾ã—ãŸã€‚%n%nä½¿ã„æ–¹: Excel ã‚’é–‹ãã¾ã™ã€‚ã‚¢ãƒ‰ã‚¤ãƒ³ãŒè‡ªå‹•ã§èª­ã¿è¾¼ã¾ã‚Œãªã„å ´åˆã¯ã€ãƒ•ã‚¡ã‚¤ãƒ« > ã‚ªãƒ—ã‚·ãƒ§ãƒ³ > ã‚¢ãƒ‰ã‚¤ãƒ³ > è¨­å®š... > å‚ç…§... ã§ %2 ã® .xll ã‚’é¸æŠžã—ã¦ãã ã•ã„ã€‚
japanese.UninstKeepConfig=RedisExcel.json ã®è¨­å®šã‚’ä¿æŒã—ã¾ã™ã‹ï¼Ÿ
japanese.UninstServerFound=%1  [æ¤œå‡º]
japanese.UninstServerFailed=%1 を自動で削除できませんでした（終了コード %2）。必要なら手動で削除してください。
japanese.UninstServersDesc=ã“ã‚Œã‚‰ã®ã‚µãƒ¼ãƒãƒ¼ãŒã“ã®ã‚³ãƒ³ãƒ”ãƒ¥ãƒ¼ã‚¿ãƒ¼ã§è¦‹ã¤ã‹ã‚Šã¾ã—ãŸã€‚RedisExcel ã¨ä¸€ç·’ã«å‰Šé™¤ã™ã‚‹ã‚‚ã®ã‚’ãƒã‚§ãƒƒã‚¯ã—ã¦ãã ã•ã„ã€‚%næ—¢å®šã§ã¯ä½•ã‚‚å‰Šé™¤ã•ã‚Œã¾ã›ã‚“ã€‚
japanese.UninstServersSubtitle=Redis äº’æ›ã‚µãƒ¼ãƒãƒ¼ã‚‚å‰Šé™¤ã—ã¾ã™ã‹ï¼Ÿ
japanese.UninstServersTitle=ã‚µãƒ¼ãƒãƒ¼
japanese.UpgradeConfigKept=æ—¢å­˜ã® RedisExcel.json ãŒè¦‹ã¤ã‹ã‚Šã€ä¿æŒã•ã‚Œã¾ã—ãŸï¼ˆè¨­å®šã¯ä¸Šæ›¸ãã•ã‚Œã¦ã„ã¾ã›ã‚“ï¼‰ã€‚%n%nã“ã‚Œã¯å¤§å¹…ãªæ›´æ–°ã§ã™: æ§‹æˆå½¢å¼ãŒå¤‰ã‚ã‚Šã¾ã—ãŸã€‚ãã®ãƒ•ã‚¡ã‚¤ãƒ«ãŒæ—§å½¢å¼ã®å ´åˆã€ç§»è¡Œã™ã‚‹ã¾ã§å„ã‚»ãƒ«ã« #CONFIG ã‚¨ãƒ©ãƒ¼ãŒè¡¨ç¤ºã•ã‚Œã¾ã™ã€‚%n%nValidate-Config.cmdï¼ˆã¾ãŸã¯ Excel ã§ =RedisCheckConfigFile()ï¼‰ã‚’å®Ÿè¡Œã—ã€README ã® "Migrating from the old shape" ã‚’å‚ç…§ã—ã¦ãã ã•ã„ã€‚
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
  RemoveServerKind: TArrayOfString;
  Code: Integer;

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
  // Exec() runs a process directly (no shell), so a literal '> file' would be
  // passed to powershell.exe as a positional argument and the capture file would
  // never be written - detection always returned 0 kinds (audit G3). Route the
  // redirection through cmd.exe so '>' is interpreted.
  Params := '/c ""powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "' + Script + '" > "' + TmpFile + '" 2>nul"';
  if not Exec('cmd.exe', Params, ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code) then
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

procedure ImportRemoveServersParam(); forward;
procedure CaptureTickedServers(); forward;
procedure RunUninstallServers(); forward;

function InitializeUninstall(): Boolean;
var
  I, Present: Integer;
begin
  Result := True;
  UninstallServersPage := nil;
  SetArrayLength(RemoveServerKind, 0);
  try
    DetectServers();
    Present := 0;
    for I := 0 to GetArrayLength(ServerPresent) - 1 do
      if ServerPresent[I] then Present := Present + 1;
    // In a silent uninstall no wizard page can be created (CreateInputOptionPage
    // throws before the wizard exists), so the page-based choice is unavailable.
    // Support an explicit opt-in instead: /REMOVESERVERS=<kind>[,...]. This is
    // harmless in an interactive run (the parameter is usually absent) and lets
    // a silent run drive the removal; import it regardless of WizardSilent().
    ImportRemoveServersParam();
    if WizardSilent() then
      Exit;
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

// Captures the servers to remove from the /REMOVESERVERS=<kind,...> command-line
// parameter (silent uninstall) into RemoveServerKind.
procedure ImportRemoveServersParam();
var
  Raw: String;
  Sep: Integer;
  Part: String;
begin
  Raw := ParamOr('REMOVESERVERS', '');
  while Length(Raw) > 0 do
  begin
    Sep := Pos(',', Raw);
    if Sep > 0 then
    begin
      Part := Copy(Raw, 1, Sep - 1);
      Raw := Copy(Raw, Sep + 1, Length(Raw));
    end
    else
    begin
      Part := Raw;
      Raw := '';
    end;
    Part := Trim(Part);
    if Part <> '' then
    begin
      SetArrayLength(RemoveServerKind, GetArrayLength(RemoveServerKind) + 1);
      RemoveServerKind[GetArrayLength(RemoveServerKind) - 1] := Part;
    end;
  end;
end;

// Snapshots the ticked kinds into RemoveServerKind so the removal no longer
// depends on the page (which a silent run cannot create). When there is no page
// (silent uninstall) the kinds already imported from /REMOVESERVERS are kept.
procedure CaptureTickedServers();
var
  I: Integer;
begin
  if UninstallServersPage = nil then Exit;
  SetArrayLength(RemoveServerKind, 0);
  for I := 0 to GetArrayLength(ServerKinds) - 1 do
    if UninstallServersPage.Values[I] then
    begin
      SetArrayLength(RemoveServerKind, GetArrayLength(RemoveServerKind) + 1);
      RemoveServerKind[GetArrayLength(RemoveServerKind) - 1] := ServerKinds[I];
    end;
end;

procedure RunUninstallServers();
var
  I: Integer;
begin
  for I := 0 to GetArrayLength(RemoveServerKind) - 1 do
  begin
    Code := RunScript('uninstall-server.ps1', '-Kind ' + RemoveServerKind[I]);
    if (Code <> 0) and (not WizardSilent()) then
      MsgBox(FmtMessage(ExpandConstant('{cm:UninstServerFailed}'), [RemoveServerKind[I], IntToStr(Code)]),
             mbInformation, MB_OK);
    Log('UninstallServers: kind=' + RemoveServerKind[I] + ' exit=' + IntToStr(Code));
  end;
end;

// ---- Uninstall: clean Excel + ProgId, ask about the config ----------------
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterExcelAutoLoad();
    // The add-in registers this ProgId under HKCU on load; remove it.
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\RedisRtd');
    // Snapshot the ticked servers while the page still exists, then remove them.
    // This MUST run at usUninstall: the [UninstallDelete] entry removes
    // {app}\scripts (and the helper) BEFORE usPostUninstall, so calling it there
    // silently no-ops (audit F2).
    CaptureTickedServers();
    RunUninstallServers();
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    if FileExists(ExpandConstant('{app}\RedisExcel.json')) then
      if MsgBox(ExpandConstant('{cm:UninstKeepConfig}'), mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
  end;
end;

