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
; The client libraries live in the build output, not in publish\ (which holds
; only what the add-in itself needs at load time).
#ifndef LibDir
  #define LibDir "..\bin\Release\net48"
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
english.TestButton=Test connection
english.TestButtonTesting=Testing... (waits up to 5 seconds for an answer)
english.TestButtonOk=Connected to %1 - the server answered and the login was accepted.
english.TestButtonFailed=Could not connect to %1.
english.TestReasonUnknownHost=Unknown host name - check for a typo.
english.TestReasonTimeout=No answer (timeout) - is the server running and reachable?
english.TestReasonAuth=The server refused the password (or the username). Check the credentials and try again.
english.TestReasonUnexpected=The server answered, but refused the request. Check the username, the password and that this port is really a Redis server.
english.TestReasonRefused=Connection refused - is Redis running on that port?
english.TestCouldNotRun=Could not run the connection test (PowerShell unavailable?). The installer will still check the connection at the end.
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
english.UpgradeConfigKept=An existing RedisExcel.json was found and KEPT (your settings were not overwritten): it is compatible with this version, so nothing was migrated.%n%nRun Validate-Config.cmd any time to check the file yourself.
english.UpgradeConfigMigrated=The RedisExcel.json found in this folder used the OLD (pre-1.6) format, which this add-in rejects - every cell would have shown a #CONFIG error.%n%nIt was saved as %1 and a new RedisExcel.json was written from the settings you entered.
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
spanish.TestButton=Probar conexión
spanish.TestButtonTesting=Probando... (espera hasta 5 segundos la respuesta)
spanish.TestButtonOk=Conectado a %1: el servidor respondió y el inicio de sesión fue aceptado.
spanish.TestButtonFailed=No se pudo conectar a %1.
spanish.TestReasonUnknownHost=Nombre de host desconocido: revise si hay un error de escritura.
spanish.TestReasonTimeout=Sin respuesta (tiempo de espera): ¿el servidor está en marcha y accesible?
spanish.TestReasonAuth=El servidor rechazó la contraseña (o el usuario). Revise las credenciales e inténtelo de nuevo.
spanish.TestReasonUnexpected=El servidor respondió, pero rechazó la solicitud. Revise el usuario, la contraseña y que este puerto sea realmente un servidor Redis.
spanish.TestReasonRefused=Conexión rechazada: ¿Redis está en marcha en ese puerto?
spanish.TestCouldNotRun=No se pudo ejecutar la prueba de conexión (¿PowerShell no disponible?). El instalador comprobará la conexión al final de todos modos.
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
spanish.UpgradeConfigKept=Se encontró un RedisExcel.json existente y se CONSERVÓ (sus ajustes no se sobrescribieron): es compatible con esta versión, así que no se migró nada.%n%nEjecute Validate-Config.cmd cuando quiera para comprobarlo.
spanish.UpgradeConfigMigrated=El RedisExcel.json que había en esta carpeta usaba el formato ANTIGUO (anterior a 1.6), que este complemento rechaza: todas las celdas habrían mostrado el error #CONFIG.%n%nSe guardó como %1 y se escribió un RedisExcel.json nuevo con los ajustes que indicó.
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
brazilianportuguese.TestButton=Testar conexão
brazilianportuguese.TestButtonTesting=Testando... (aguarda até 5 segundos pela resposta)
brazilianportuguese.TestButtonOk=Conectado a %1 - o servidor respondeu e o login foi aceito.
brazilianportuguese.TestButtonFailed=Não foi possível conectar a %1.
brazilianportuguese.TestReasonUnknownHost=Nome de host desconhecido - confira se há um erro de digitação.
brazilianportuguese.TestReasonTimeout=Sem resposta (tempo esgotado) - o servidor está rodando e acessível?
brazilianportuguese.TestReasonAuth=O servidor recusou a senha (ou o usuário). Confira as credenciais e tente de novo.
brazilianportuguese.TestReasonUnexpected=O servidor respondeu, mas recusou a solicitação. Confira o usuário, a senha e se esta porta é mesmo um servidor Redis.
brazilianportuguese.TestReasonRefused=Conexão recusada - o Redis está rodando nessa porta?
brazilianportuguese.TestCouldNotRun=Não foi possível executar o teste de conexão (PowerShell indisponível?). O instalador vai testar a conexão no final de qualquer forma.
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
brazilianportuguese.UpgradeConfigKept=Um RedisExcel.json existente foi encontrado e MANTIDO (suas configurações não foram sobrescritas): ele é compatível com esta versão, então nada foi migrado.%n%nRode o Validate-Config.cmd quando quiser para conferir o arquivo.
brazilianportuguese.UpgradeConfigMigrated=O RedisExcel.json desta pasta usava o formato ANTIGO (pré-1.6), que este add-in rejeita: toda célula mostraria o erro #CONFIG.%n%nEle foi salvo como %1 e um novo RedisExcel.json foi escrito com as configurações que você informou.
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
chinesesimplified.AdvPageDesc=这就是完整的 RedisExcel.json，可直接编辑：每个键都已预填默认值和简短注释。加载项接受 // 注释，因此可以保留。可在此修改主机、别名、写入模式、ConflationMs、AsyncWrites——全部在此。
chinesesimplified.AdvPageTitle=高级连接
chinesesimplified.ConfigWriteFailed=无法写入 %1。请参照示例（RedisExcel.json）手动创建。
chinesesimplified.ConnHost=主机（或 host:port）：
chinesesimplified.ConnPageDesc=使用默认值（localhost:6379）请点“下一步”。之后可编辑 RedisExcel.json 修改。
chinesesimplified.ConnPageSubtitle=RedisExcel 应连接到何处？
chinesesimplified.ConnPageTitle=Redis 连接
chinesesimplified.ConnPassword=密码（可选）：
chinesesimplified.TestButton=测试连接
chinesesimplified.TestButtonTesting=正在测试...（最多等待 5 秒）
chinesesimplified.TestButtonOk=已连接到 %1 - 服务器已响应，登录已被接受。
chinesesimplified.TestButtonFailed=无法连接到 %1。
chinesesimplified.TestReasonUnknownHost=主机名未知 - 请检查是否拼写错误。
chinesesimplified.TestReasonTimeout=无响应（超时） - 服务器是否在运行且可访问？
chinesesimplified.TestReasonAuth=服务器拒绝了密码（或用户名）。请检查凭据后重试。
chinesesimplified.TestReasonUnexpected=服务器有响应，但拒绝了请求。请检查用户名、密码，以及该端口是否真的是 Redis 服务器。
chinesesimplified.TestReasonRefused=连接被拒绝 - 该端口上是否运行着 Redis？
chinesesimplified.TestCouldNotRun=无法运行连接测试（PowerShell 不可用？）。安装程序仍会在结束时测试连接。
chinesesimplified.ConnPort=端口：
chinesesimplified.FinishedText=RedisExcel 已安装到：%n%2%n%n如果加载项未自动加载，请在 Excel 中启用：%n文件 > 选项 > 加载项 > 转到... > 浏览...，选择 RedisExcel64-packed.xll（32 位 Office 用 RedisExcel-packed.xll）。%n%n然后在单元格中试试 =RedisUDFGet("preco_btc")。
chinesesimplified.GroupConfig=配置：
chinesesimplified.GroupExcel=Excel：
chinesesimplified.HostRequired=请输入 Redis 主机（例如 localhost）。
chinesesimplified.IconUninstall=卸载 {#AppName}
chinesesimplified.ModeAdvanced=高级 - 设置数据库、SSL、写入模式和更新调节
chinesesimplified.ModePageDesc=简易仅询问服务器地址。高级可设置数据库、SSL/TLS、写入模式和实时更新调节。之后随时可编辑 RedisExcel.json（或运行 Validate-Config.cmd）。
chinesesimplified.ModePageSubtitle=简易还是高级安装？
chinesesimplified.ModePageTitle=安装模式
chinesesimplified.ModeSimple=简易 - 仅 Redis 地址（推荐）
chinesesimplified.OptionsOpenExcel=结束时打开示例 RedisExcel 工作簿
chinesesimplified.OptionsPageDesc=安装程序在结束时测试连接。可选择打开示例 Excel 工作簿以演示其工作方式。
chinesesimplified.OptionsPageSubtitle=即将完成
chinesesimplified.OptionsPageTitle=额外检查
chinesesimplified.ServerDockerRedis=是 - Docker：Redis（redis:latest）
chinesesimplified.ServerDockerValkey=是 - Docker：Valkey（valkey/valkey:latest）
chinesesimplified.ServerFailed=无法自动设置 %1（退出码 %2）。%n您可以稍后安装服务器，并在 RedisExcel.json 中指向它。
chinesesimplified.ServerMemurai=是 - Memurai（Windows 本机服务，需要管理员权限）
chinesesimplified.ServerNameDockerRedis=Redis（Docker）
chinesesimplified.ServerNameDockerValkey=Valkey（Docker）
chinesesimplified.ServerNameGeneric=服务器
chinesesimplified.ServerNameMemurai=Memurai
chinesesimplified.ServerNameWslRedis=Redis（WSL）
chinesesimplified.ServerNameWslValkey=Valkey（WSL）
chinesesimplified.ServerNone=否 - 我已有服务器（或稍后设置）
chinesesimplified.ServerPageDesc=RedisExcel 需要 Redis（或 Valkey）服务器。如果还没有，请选择一个选项。先安装服务器，再安装 RedisExcel，最后测试连接。Docker 最简单。
chinesesimplified.ServerPageSubtitle=现在安装兼容 Redis 的服务器吗？
chinesesimplified.ServerPageTitle=Redis 服务器
chinesesimplified.ServerWslRedis=是 - WSL2：Redis（apt install redis-server）
chinesesimplified.ServerWslValkey=是 - WSL2：Valkey（apt install valkey-server）
chinesesimplified.SummaryAdvanced=高级
chinesesimplified.SummaryIntro=这些值将保存到安装文件夹中的 RedisExcel.json。
chinesesimplified.SummaryLabel=连接
chinesesimplified.SummaryMode=模式
chinesesimplified.SummaryPageSubtitle=查看将写入 RedisExcel.json 的连接
chinesesimplified.SummaryPageTitle=您的设置
chinesesimplified.SummarySimple=简易
chinesesimplified.TaskAutoLoadDesc=注册加载项，使 Excel 自动加载（推荐）
chinesesimplified.TaskConfigDesc=在安装文件夹中创建 RedisExcel.json（主机设置）
chinesesimplified.TaskNlogDesc=创建 NLog.config（在加载项旁写入日志文件）
chinesesimplified.TestFailed=已安装，但 %1 处的 Redis 未响应 PING。%n请确认服务器正在运行（docker ps），如果它在其他主机/端口，请调整 RedisExcel.json。
chinesesimplified.TestOk=全部就绪！%1 处的 Redis 已响应。%n%n使用方法：打开 Excel。如果加载项未自动加载，请转到 文件 > 选项 > 加载项 > 转到... > 浏览...，选择 %2 中的 .xll。
chinesesimplified.UninstKeepConfig=保留您的 RedisExcel.json 设置吗？
chinesesimplified.UninstServerFound=%1  [已找到]
chinesesimplified.UninstServerFailed=?????? %1(??? %2)???????????
chinesesimplified.UninstServersDesc=在此计算机上发现了这些服务器。勾选要随 RedisExcel 一起移除的项。%n默认不移除任何内容。
chinesesimplified.UninstServersSubtitle=也移除兼容 Redis 的服务器吗？
chinesesimplified.UninstServersTitle=服务器
chinesesimplified.UpgradeConfigKept=发现已有的 RedisExcel.json,已保留(未覆盖您的设置):它与本版本兼容,因此没有迁移任何内容。%n%n可随时运行 Validate-Config.cmd 检查该文件。
chinesesimplified.UpgradeConfigMigrated=该文件夹中的 RedisExcel.json 使用旧格式(1.6 之前),本加载项会拒绝该格式:所有单元格都会显示 #CONFIG 错误。%n%n原文件已保存为 %1,并已按您填写的设置写入新的 RedisExcel.json。

; ---- chinesetraditional ----
chinesetraditional.AdvPageDesc=這就是完整的 RedisExcel.json，可直接編輯：每個鍵都已預填預設值與簡短註解。增益集接受 // 註解，因此可以保留。可在此修改主機、別名、寫入模式、ConflationMs、AsyncWrites——全都在此。
chinesetraditional.AdvPageTitle=進階連線
chinesetraditional.ConfigWriteFailed=無法寫入 %1。請依範例（RedisExcel.json）手動建立。
chinesetraditional.ConnHost=主機（或 host:port）：
chinesetraditional.ConnPageDesc=使用預設值（localhost:6379）請按「下一步」。之後可編輯 RedisExcel.json 修改。
chinesetraditional.ConnPageSubtitle=RedisExcel 應連線到何處？
chinesetraditional.ConnPageTitle=Redis 連線
chinesetraditional.ConnPassword=密碼（選填）：
chinesetraditional.TestButton=測試連線
chinesetraditional.TestButtonTesting=正在測試...（最多等待 5 秒）
chinesetraditional.TestButtonOk=已連線到 %1 - 伺服器已回應，登入已被接受。
chinesetraditional.TestButtonFailed=無法連線到 %1。
chinesetraditional.TestReasonUnknownHost=主機名稱未知 - 請檢查是否拼寫錯誤。
chinesetraditional.TestReasonTimeout=無回應（逾時） - 伺服器是否在執行且可存取？
chinesetraditional.TestReasonAuth=伺服器拒絕了密碼（或使用者名稱）。請檢查認證資訊後重試。
chinesetraditional.TestReasonUnexpected=伺服器有回應，但拒絕了請求。請檢查使用者名稱、密碼，以及該連接埠是否真的是 Redis 伺服器。
chinesetraditional.TestReasonRefused=連線被拒絕 - 該連接埠上是否執行著 Redis？
chinesetraditional.TestCouldNotRun=無法執行連線測試（PowerShell 無法使用？）。安裝程式仍會在結束時測試連線。
chinesetraditional.ConnPort=連接埠：
chinesetraditional.FinishedText=RedisExcel 已安裝於：%n%2%n%n若增益集未自動載入，請在 Excel 中啟用：%n檔案 > 選項 > 增益集 > 前往... > 瀏覽...，選擇 RedisExcel64-packed.xll（32 位元 Office 請用 RedisExcel-packed.xll）。%n%n接著在儲存格中試試 =RedisUDFGet("preco_btc")。
chinesetraditional.GroupConfig=設定：
chinesetraditional.GroupExcel=Excel：
chinesetraditional.HostRequired=請輸入 Redis 主機（例如 localhost）。
chinesetraditional.IconUninstall=解除安裝 {#AppName}
chinesetraditional.ModeAdvanced=進階 - 設定資料庫、SSL、寫入模式與更新微調
chinesetraditional.ModePageDesc=簡易只詢問伺服器位址。進階可設定資料庫、SSL/TLS、寫入模式與即時更新微調。之後隨時可編輯 RedisExcel.json（或執行 Validate-Config.cmd）。
chinesetraditional.ModePageSubtitle=簡易或進階安裝？
chinesetraditional.ModePageTitle=安裝模式
chinesetraditional.ModeSimple=簡易 - 僅 Redis 位址（建議）
chinesetraditional.OptionsOpenExcel=結束時開啟範例 RedisExcel 活頁簿
chinesetraditional.OptionsPageDesc=安裝程式會在結束時測試連線。可選擇開啟範例 Excel 活頁簿以展示其運作方式。
chinesetraditional.OptionsPageSubtitle=即將完成
chinesetraditional.OptionsPageTitle=額外檢查
chinesetraditional.ServerDockerRedis=是 - Docker：Redis（redis:latest）
chinesetraditional.ServerDockerValkey=是 - Docker：Valkey（valkey/valkey:latest）
chinesetraditional.ServerFailed=無法自動設定 %1（結束碼 %2）。%n您可以稍後安裝伺服器，並在 RedisExcel.json 中指向它。
chinesetraditional.ServerMemurai=是 - Memurai（Windows 原生服務，需要系統管理員權限）
chinesetraditional.ServerNameDockerRedis=Redis（Docker）
chinesetraditional.ServerNameDockerValkey=Valkey（Docker）
chinesetraditional.ServerNameGeneric=伺服器
chinesetraditional.ServerNameMemurai=Memurai
chinesetraditional.ServerNameWslRedis=Redis（WSL）
chinesetraditional.ServerNameWslValkey=Valkey（WSL）
chinesetraditional.ServerNone=否 - 我已有伺服器（或稍後設定）
chinesetraditional.ServerPageDesc=RedisExcel 需要 Redis（或 Valkey）伺服器。如果還沒有，請選擇一個選項。會先安裝伺服器，再安裝 RedisExcel，最後測試連線。Docker 最簡單。
chinesetraditional.ServerPageSubtitle=現在安裝相容於 Redis 的伺服器嗎？
chinesetraditional.ServerPageTitle=Redis 伺服器
chinesetraditional.ServerWslRedis=是 - WSL2：Redis（apt install redis-server）
chinesetraditional.ServerWslValkey=是 - WSL2：Valkey（apt install valkey-server）
chinesetraditional.SummaryAdvanced=進階
chinesetraditional.SummaryIntro=這些值將儲存到安裝資料夾中的 RedisExcel.json。
chinesetraditional.SummaryLabel=連線
chinesetraditional.SummaryMode=模式
chinesetraditional.SummaryPageSubtitle=檢閱將寫入 RedisExcel.json 的連線
chinesetraditional.SummaryPageTitle=您的設定
chinesetraditional.SummarySimple=簡易
chinesetraditional.TaskAutoLoadDesc=註冊增益集，讓 Excel 自動載入（建議）
chinesetraditional.TaskConfigDesc=在安裝資料夾中建立 RedisExcel.json（主機設定）
chinesetraditional.TaskNlogDesc=建立 NLog.config（在增益集旁寫入記錄檔）
chinesetraditional.TestFailed=已安裝，但 %1 的 Redis 未回應 PING。%n請確認伺服器正在執行（docker ps），若在其他主機/連接埠，請調整 RedisExcel.json。
chinesetraditional.TestOk=全部就緒！%1 的 Redis 已回應。%n%n使用方法：開啟 Excel。若增益集未自動載入，請前往 檔案 > 選項 > 增益集 > 前往... > 瀏覽...，選擇 %2 中的 .xll。
chinesetraditional.UninstKeepConfig=保留您的 RedisExcel.json 設定嗎？
chinesetraditional.UninstServerFound=%1  [已找到]
chinesetraditional.UninstServerFailed=?????? %1(??? %2)???????????
chinesetraditional.UninstServersDesc=在此電腦上找到這些伺服器。勾選要隨 RedisExcel 一併移除的項目。%n預設不移除任何項目。
chinesetraditional.UninstServersSubtitle=也移除相容於 Redis 的伺服器嗎？
chinesetraditional.UninstServersTitle=伺服器
chinesetraditional.UpgradeConfigKept=發現現有的 RedisExcel.json,已保留(未覆寫您的設定):它與本版本相容,因此未進行任何移轉。%n%n可隨時執行 Validate-Config.cmd 檢查該檔案。
chinesetraditional.UpgradeConfigMigrated=此資料夾中的 RedisExcel.json 使用舊格式(1.6 之前),此外掛會拒絕該格式:所有儲存格都會顯示 #CONFIG 錯誤。%n%n原檔案已儲存為 %1,並已依您填寫的設定寫入新的 RedisExcel.json。

; ---- russian ----
russian.AdvPageDesc=Это весь ваш RedisExcel.json, готовый к правке: каждый ключ уже заполнен значением по умолчанию и коротким комментарием. Адд-ин принимает комментарии //, поэтому их можно оставить. Меняйте хосты, псевдонимы, режим записи, ConflationMs, AsyncWrites - всё здесь.
russian.AdvPageTitle=Дополнительное подключение
russian.ConfigWriteFailed=Не удалось записать %1. Создайте его вручную по образцу (RedisExcel.json).
russian.ConnHost=Хост (или host:port):
russian.ConnPageDesc=Нажмите «Далее» для значения по умолчанию (localhost:6379). Позже это можно изменить в RedisExcel.json.
russian.ConnPageSubtitle=Куда подключаться RedisExcel?
russian.ConnPageTitle=Подключение к Redis
russian.ConnPassword=Пароль (необязательно):
russian.TestButton=Проверить подключение
russian.TestButtonTesting=Проверка... (ответ ожидается до 5 секунд)
russian.TestButtonOk=Подключено к %1 - сервер ответил, вход принят.
russian.TestButtonFailed=Не удалось подключиться к %1.
russian.TestReasonUnknownHost=Неизвестное имя хоста - проверьте, нет ли опечатки.
russian.TestReasonTimeout=Нет ответа (тайм-аут) - сервер запущен и доступен?
russian.TestReasonAuth=Сервер отклонил пароль (или имя пользователя). Проверьте учётные данные и попробуйте снова.
russian.TestReasonUnexpected=Сервер ответил, но отклонил запрос. Проверьте имя пользователя, пароль и то, что на этом порту действительно Redis.
russian.TestReasonRefused=В подключении отказано - запущен ли Redis на этом порту?
russian.TestCouldNotRun=Не удалось запустить проверку подключения (PowerShell недоступен?). Установщик всё равно проверит подключение в конце.
russian.ConnPort=Порт:
russian.FinishedText=RedisExcel установлен в:%n%2%n%nЕсли адд-ин не загружается автоматически, включите его в Excel:%nФайл > Параметры > Надстройки > Перейти... > Обзор... и выберите RedisExcel64-packed.xll (или RedisExcel-packed.xll для 32-битного Office).%n%nЗатем попробуйте =RedisUDFGet("preco_btc") в ячейке.
russian.GroupConfig=Конфигурация:
russian.GroupExcel=Excel:
russian.HostRequired=Введите хост Redis (например, localhost).
russian.IconUninstall=Удалить {#AppName}
russian.ModeAdvanced=Расширенный - база данных, SSL, режим записи и настройка обновлений
russian.ModePageDesc=Простой запрашивает только адрес сервера. Расширенный позволяет задать базу данных, SSL/TLS, режим записи и настройку обновлений в реальном времени. RedisExcel.json всегда можно отредактировать позже (или запустить Validate-Config.cmd).
russian.ModePageSubtitle=Простая или расширенная установка?
russian.ModePageTitle=Режим установки
russian.ModeSimple=Простой - только адрес Redis (рекомендуется)
russian.OptionsOpenExcel=Открыть пример книги RedisExcel в конце
russian.OptionsPageDesc=Установщик проверит подключение по завершении. При желании откроется пример книги Excel с демонстрацией работы.
russian.OptionsPageSubtitle=Почти готово
russian.OptionsPageTitle=Дополнительные проверки
russian.ServerDockerRedis=Да - Docker: Redis (redis:latest)
russian.ServerDockerValkey=Да - Docker: Valkey (valkey/valkey:latest)
russian.ServerFailed=Не удалось настроить %1 автоматически (код %2).%nВы можете установить сервер позже и указать его в RedisExcel.json.
russian.ServerMemurai=Да - Memurai (служба Windows, нужны права администратора)
russian.ServerNameDockerRedis=Redis (Docker)
russian.ServerNameDockerValkey=Valkey (Docker)
russian.ServerNameGeneric=сервер
russian.ServerNameMemurai=Memurai
russian.ServerNameWslRedis=Redis (WSL)
russian.ServerNameWslValkey=Valkey (WSL)
russian.ServerNone=Нет - сервер уже есть (или настрою позже)
russian.ServerPageDesc=RedisExcel нужен сервер Redis (или Valkey). Выберите вариант, если его ещё нет. Сначала устанавливается сервер, затем RedisExcel, а подключение проверяется в конце. Docker - самый простой вариант.
russian.ServerPageSubtitle=Установить сейчас совместимый с Redis сервер?
russian.ServerPageTitle=Сервер Redis
russian.ServerWslRedis=Да - WSL2: Redis (apt install redis-server)
russian.ServerWslValkey=Да - WSL2: Valkey (apt install valkey-server)
russian.SummaryAdvanced=Расширенный
russian.SummaryIntro=Эти значения будут сохранены в RedisExcel.json в папке установки.
russian.SummaryLabel=Подключение
russian.SummaryMode=Режим
russian.SummaryPageSubtitle=Проверьте подключение, которое будет записано в RedisExcel.json
russian.SummaryPageTitle=Ваши настройки
russian.SummarySimple=Простой
russian.TaskAutoLoadDesc=Зарегистрировать адд-ин, чтобы Excel загружал его автоматически (рекомендуется)
russian.TaskConfigDesc=Создать RedisExcel.json в папке установки (настройки хоста)
russian.TaskNlogDesc=Создать NLog.config (писать лог-файл рядом с адд-ином)
russian.TestFailed=Установлено, но Redis по адресу %1 не ответил на PING.%nПроверьте, что сервер запущен (docker ps), и поправьте RedisExcel.json, если он на другом хосте/порту.
russian.TestOk=Готово! Redis ответил по адресу %1.%n%nЧтобы начать: откройте Excel. Если адд-ин не загружается автоматически, перейдите в Файл > Параметры > Надстройки > Перейти... > Обзор... и выберите .xll в %2.
russian.UninstKeepConfig=Сохранить ваши настройки RedisExcel.json?
russian.UninstServerFound=%1  [найдено]
russian.UninstServerFailed=?? ??????? ????????????? ??????? %1 (??? %2). ??????? ??????? ??? ?????????????.
russian.UninstServersDesc=Эти серверы найдены на компьютере. Отметьте те, которые нужно удалить вместе с RedisExcel.%nПо умолчанию НИЧЕГО не удаляется.
russian.UninstServersSubtitle=Удалить также совместимые с Redis серверы?
russian.UninstServersTitle=Серверы
russian.UpgradeConfigKept=Найден существующий RedisExcel.json — он СОХРАНЁН (ваши настройки не перезаписаны) и совместим с этой версией, поэтому ничего не переносилось.%n%nПроверить файл можно через Validate-Config.cmd.
russian.UpgradeConfigMigrated=Файл RedisExcel.json в этой папке использовал СТАРЫЙ формат (до 1.6), который это дополнение отклоняет: все ячейки показывали бы ошибку #CONFIG.%n%nОн сохранён как %1, а новый RedisExcel.json записан с вашими параметрами.

; ---- french ----
french.AdvPageDesc=Voici tout votre RedisExcel.json, prêt à modifier : chaque clé est préremplie avec sa valeur par défaut et un court commentaire. Le complément accepte les commentaires //, vous pouvez donc les laisser. Modifiez les hôtes, alias, mode d'écriture, ConflationMs, AsyncWrites - tout est ici.
french.AdvPageTitle=Connexion avancée
french.ConfigWriteFailed=Impossible d'écrire %1. Créez-le manuellement à partir de l'exemple (RedisExcel.json).
french.ConnHost=Hôte (ou hôte:port) :
french.ConnPageDesc=Cliquez sur Suivant pour la valeur par défaut (localhost:6379). Vous pourrez la modifier plus tard en éditant RedisExcel.json.
french.ConnPageSubtitle=Où RedisExcel doit-il se connecter ?
french.ConnPageTitle=Connexion Redis
french.ConnPassword=Mot de passe (facultatif) :
french.TestButton=Tester la connexion
french.TestButtonTesting=Test en cours... (attente de la réponse : jusqu'à 5 secondes)
french.TestButtonOk=Connecté à %1 - le serveur a répondu et la connexion a été acceptée.
french.TestButtonFailed=Impossible de se connecter à %1.
french.TestReasonUnknownHost=Nom d'hôte inconnu - vérifiez s'il y a une faute de frappe.
french.TestReasonTimeout=Aucune réponse (délai dépassé) - le serveur est-il démarré et joignable ?
french.TestReasonAuth=Le serveur a refusé le mot de passe (ou l'utilisateur). Vérifiez les identifiants et réessayez.
french.TestReasonUnexpected=Le serveur a répondu, mais a refusé la demande. Vérifiez l'utilisateur, le mot de passe et que ce port est bien un serveur Redis.
french.TestReasonRefused=Connexion refusée - Redis est-il démarré sur ce port ?
french.TestCouldNotRun=Impossible d'exécuter le test de connexion (PowerShell indisponible ?). L'installateur vérifiera quand même la connexion à la fin.
french.ConnPort=Port :
french.FinishedText=RedisExcel est installé dans :%n%2%n%nSi le complément ne se charge pas automatiquement, activez-le dans Excel :%nFichier > Options > Compléments > Atteindre... > Parcourir... et choisissez RedisExcel64-packed.xll (ou RedisExcel-packed.xll pour Office 32 bits).%n%nEssayez ensuite =RedisUDFGet("preco_btc") dans une cellule.
french.GroupConfig=Configuration :
french.GroupExcel=Excel :
french.HostRequired=Veuillez saisir un hôte Redis (par exemple localhost).
french.IconUninstall=Désinstaller {#AppName}
french.ModeAdvanced=Avancé - définir la base de données, SSL, le mode d'écriture et le réglage des mises à jour
french.ModePageDesc=Simple ne demande que l'adresse du serveur. Avancé permet de définir la base de données, SSL/TLS, le mode d'écriture et le réglage des mises à jour en temps réel. Vous pourrez toujours modifier RedisExcel.json plus tard (ou exécuter Validate-Config.cmd).
french.ModePageSubtitle=Installation simple ou avancée ?
french.ModePageTitle=Mode d'installation
french.ModeSimple=Simple - juste l'adresse Redis (recommandé)
french.OptionsOpenExcel=Ouvrir un classeur RedisExcel d'exemple à la fin
french.OptionsPageDesc=L'installateur teste la connexion à la fin. Ouvrez éventuellement un classeur Excel d'exemple montrant son fonctionnement.
french.OptionsPageSubtitle=Presque terminé
french.OptionsPageTitle=Vérifications supplémentaires
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
french.ServerNone=Non - j'ai déjà un serveur (ou je le configurerai plus tard)
french.ServerPageDesc=RedisExcel a besoin d'un serveur Redis (ou Valkey). Choisissez une option si vous n'en avez pas encore. Le serveur est installé d'abord, puis RedisExcel, et la connexion est testée à la fin. Docker est le plus simple.
french.ServerPageSubtitle=Voulez-vous installer un serveur compatible Redis maintenant ?
french.ServerPageTitle=Serveur Redis
french.ServerWslRedis=Oui - WSL2 : Redis (apt install redis-server)
french.ServerWslValkey=Oui - WSL2 : Valkey (apt install valkey-server)
french.SummaryAdvanced=Avancé
french.SummaryIntro=Ces valeurs seront enregistrées dans RedisExcel.json, dans le dossier d'installation.
french.SummaryLabel=Connexion
french.SummaryMode=Mode
french.SummaryPageSubtitle=Vérifiez la connexion qui sera écrite dans RedisExcel.json
french.SummaryPageTitle=Vos paramètres
french.SummarySimple=Simple
french.TaskAutoLoadDesc=Enregistrer le complément pour qu'Excel le charge automatiquement (recommandé)
french.TaskConfigDesc=Créer RedisExcel.json dans le dossier d'installation (paramètres de l'hôte)
french.TaskNlogDesc=Créer NLog.config (écrire un journal à côté du complément)
french.TestFailed=Installé, mais le Redis à %1 n'a pas répondu à un PING.%nVérifiez que le serveur est démarré (docker ps) et ajustez RedisExcel.json s'il est sur un autre hôte/port.
french.TestOk=Tout est prêt ! Redis a répondu à %1.%n%nPour l'utiliser : ouvrez Excel. Si le complément ne se charge pas automatiquement, allez dans Fichier > Options > Compléments > Atteindre... > Parcourir... et choisissez le .xll dans %2.
french.UninstKeepConfig=Conserver vos paramètres RedisExcel.json ?
french.UninstServerFound=%1  [trouvé]
french.UninstServerFailed=Impossible de supprimer %1 automatiquement (code %2). Supprimez-le manuellement si besoin.
french.UninstServersDesc=Ces serveurs ont été trouvés sur cette machine. Cochez ceux à supprimer avec RedisExcel.%nPar défaut, RIEN n'est supprimé.
french.UninstServersSubtitle=Supprimer aussi les serveurs compatibles Redis ?
french.UninstServersTitle=Serveurs
french.UpgradeConfigKept=Un RedisExcel.json existant a été trouvé et CONSERVÉ (vos paramètres n'ont pas été écrasés) : il est compatible avec cette version, rien n'a donc été migré.%n%nLancez Validate-Config.cmd quand vous voulez pour le vérifier.
french.UpgradeConfigMigrated=Le RedisExcel.json de ce dossier utilisait l'ANCIEN format (pré-1.6), que ce complément refuse : toutes les cellules auraient affiché l'erreur #CONFIG.%n%nIl a été enregistré sous %1 et un nouveau RedisExcel.json a été écrit avec les paramètres saisis.

; ---- korean ----
korean.AdvPageDesc=이것은 편집할 수 있는 전체 RedisExcel.json입니다. 각 키에는 기본값과 짧은 설명이 미리 채워져 있습니다. 추가 기능은 // 주석을 허용하므로 그대로 두어도 됩니다. 호스트, 별칭, 쓰기 모드, ConflationMs, AsyncWrites를 여기서 변경하세요.
korean.AdvPageTitle=고급 연결
korean.ConfigWriteFailed=%1을(를) 쓸 수 없습니다. 예제(RedisExcel.json)를 참고해 수동으로 만드세요.
korean.ConnHost=호스트(또는 host:port):
korean.ConnPageDesc=기본값(localhost:6379)이면 다음을 누르세요. 나중에 RedisExcel.json을 편집해 변경할 수 있습니다.
korean.ConnPageSubtitle=RedisExcel이 어디에 연결해야 합니까?
korean.ConnPageTitle=Redis 연결
korean.ConnPassword=비밀번호(선택):
korean.TestButton=연결 테스트
korean.TestButtonTesting=테스트 중... (응답을 최대 5초 기다립니다)
korean.TestButtonOk=%1에 연결됨 - 서버가 응답했고 로그인이 승인되었습니다.
korean.TestButtonFailed=%1에 연결할 수 없습니다.
korean.TestReasonUnknownHost=알 수 없는 호스트 이름 - 오타가 있는지 확인하세요.
korean.TestReasonTimeout=응답 없음(시간 초과) - 서버가 실행 중이며 접근 가능한가요?
korean.TestReasonAuth=서버가 비밀번호(또는 사용자 이름)를 거부했습니다. 자격 증명을 확인하고 다시 시도하세요.
korean.TestReasonUnexpected=서버가 응답했지만 요청을 거부했습니다. 사용자 이름, 비밀번호, 그리고 이 포트가 실제 Redis 서버인지 확인하세요.
korean.TestReasonRefused=연결이 거부됨 - 해당 포트에서 Redis가 실행 중인가요?
korean.TestCouldNotRun=연결 테스트를 실행할 수 없습니다(PowerShell을 사용할 수 없음?). 설치 프로그램이 끝에서 연결을 확인합니다.
korean.ConnPort=포트:
korean.FinishedText=RedisExcel이 다음 위치에 설치되었습니다:%n%2%n%n추가 기능이 자동으로 로드되지 않으면 Excel에서 활성화하세요:%n파일 > 옵션 > 추가 기능 > 이동... > 찾아보기...에서 RedisExcel64-packed.xll(32비트 Office는 RedisExcel-packed.xll)을 선택합니다.%n%n그런 다음 셀에 =RedisUDFGet("preco_btc")를 입력해 보세요.
korean.GroupConfig=구성:
korean.GroupExcel=Excel:
korean.HostRequired=Redis 호스트를 입력하세요(예: localhost).
korean.IconUninstall={#AppName} 제거
korean.ModeAdvanced=고급 - 데이터베이스, SSL, 쓰기 모드 및 업데이트 튜닝 설정
korean.ModePageDesc=간단은 서버 주소만 묻습니다. 고급은 데이터베이스, SSL/TLS, 쓰기 모드, 실시간 업데이트 튜닝을 설정할 수 있습니다. RedisExcel.json은 나중에 언제든지 편집할 수 있습니다(Validate-Config.cmd 실행 가능).
korean.ModePageSubtitle=간단 또는 고급 설치?
korean.ModePageTitle=설치 모드
korean.ModeSimple=간단 - Redis 주소만(권장)
korean.OptionsOpenExcel=마지막에 샘플 RedisExcel 통합 문서 열기
korean.OptionsPageDesc=설치 프로그램이 끝날 때 연결을 테스트합니다. 원하면 작동을 보여 주는 샘플 Excel 통합 문서를 엽니다.
korean.OptionsPageSubtitle=거의 완료
korean.OptionsPageTitle=추가 확인
korean.ServerDockerRedis=예 - Docker: Redis(redis:latest)
korean.ServerDockerValkey=예 - Docker: Valkey(valkey/valkey:latest)
korean.ServerFailed=%1을(를) 자동으로 설정하지 못했습니다(종료 코드 %2).%n나중에 서버를 설치하고 RedisExcel.json에서 지정할 수 있습니다.
korean.ServerMemurai=예 - Memurai(네이티브 Windows 서비스, 관리자 필요)
korean.ServerNameDockerRedis=Redis(Docker)
korean.ServerNameDockerValkey=Valkey(Docker)
korean.ServerNameGeneric=서버
korean.ServerNameMemurai=Memurai
korean.ServerNameWslRedis=Redis(WSL)
korean.ServerNameWslValkey=Valkey(WSL)
korean.ServerNone=아니요 - 이미 서버가 있습니다(또는 나중에 설정)
korean.ServerPageDesc=RedisExcel에는 Redis(또는 Valkey) 서버가 필요합니다. 아직 없으면 옵션을 선택하세요. 서버를 먼저 설치한 다음 RedisExcel을 설치하고 마지막에 연결을 테스트합니다. Docker가 가장 간단합니다.
korean.ServerPageSubtitle=지금 Redis 호환 서버를 설치하시겠습니까?
korean.ServerPageTitle=Redis 서버
korean.ServerWslRedis=예 - WSL2: Redis(apt install redis-server)
korean.ServerWslValkey=예 - WSL2: Valkey(apt install valkey-server)
korean.SummaryAdvanced=고급
korean.SummaryIntro=이 값들은 설치 폴더의 RedisExcel.json에 저장됩니다.
korean.SummaryLabel=연결
korean.SummaryMode=모드
korean.SummaryPageSubtitle=RedisExcel.json에 기록될 연결을 확인하세요
korean.SummaryPageTitle=설정
korean.SummarySimple=간단
korean.TaskAutoLoadDesc=Excel이 자동으로 로드하도록 추가 기능 등록(권장)
korean.TaskConfigDesc=설치 폴더에 RedisExcel.json 생성(호스트 설정)
korean.TaskNlogDesc=NLog.config 생성(추가 기능 옆에 로그 파일 기록)
korean.TestFailed=설치되었지만 %1의 Redis가 PING에 응답하지 않았습니다.%n서버가 실행 중인지 확인하고(docker ps) 다른 호스트/포트라면 RedisExcel.json을 조정하세요.
korean.TestOk=준비 완료! %1의 Redis가 응답했습니다.%n%n사용법: Excel을 여세요. 추가 기능이 자동으로 로드되지 않으면 파일 > 옵션 > 추가 기능 > 이동... > 찾아보기...에서 %2의 .xll을 선택하세요.
korean.UninstKeepConfig=RedisExcel.json 설정을 유지하시겠습니까?
korean.UninstServerFound=%1  [발견]
korean.UninstServerFailed=%1?(?) ???? ???? ?????(?? ?? %2). ???? ???? ?????.
korean.UninstServersDesc=이 서버들이 이 컴퓨터에서 발견되었습니다. RedisExcel과 함께 제거할 항목을 선택하세요.%n기본적으로 아무것도 제거되지 않습니다.
korean.UninstServersSubtitle=Redis 호환 서버도 제거하시겠습니까?
korean.UninstServersTitle=서버
korean.UpgradeConfigKept=기존 RedisExcel.json을 찾아 유지했습니다(설정이 덮어쓰이지 않았습니다). 이 버전과 호환되므로 마이그레이션하지 않았습니다.%n%nValidate-Config.cmd로 언제든지 확인할 수 있습니다.
korean.UpgradeConfigMigrated=이 폴더의 RedisExcel.json은 이 추가 기능이 거부하는 예전(1.6 이전) 형식이었습니다. 모든 셀에 #CONFIG 오류가 표시될 상황이었습니다.%n%n원본은 %1(으)로 저장하고, 입력한 설정으로 새 RedisExcel.json을 작성했습니다.

; ---- german ----
german.AdvPageDesc=Dies ist Ihre gesamte RedisExcel.json, bereit zum Bearbeiten: jeder Schlüssel ist mit seinem Standardwert und einem kurzen Kommentar vorbelegt. Das Add-in akzeptiert //-Kommentare, Sie können sie also stehen lassen. Ändern Sie Hosts, Aliase, Schreibmodus, ConflationMs, AsyncWrites - alles hier.
german.AdvPageTitle=Erweiterte Verbindung
german.ConfigWriteFailed=%1 konnte nicht geschrieben werden. Erstellen Sie es manuell nach dem Beispiel (RedisExcel.json).
german.ConnHost=Host (oder host:port):
german.ConnPageDesc=Klicken Sie auf Weiter für den Standard (localhost:6379). Sie können dies später in RedisExcel.json ändern.
german.ConnPageSubtitle=Womit soll sich RedisExcel verbinden?
german.ConnPageTitle=Redis-Verbindung
german.ConnPassword=Passwort (optional):
german.TestButton=Verbindung testen
german.TestButtonTesting=Wird getestet... (wartet bis zu 5 Sekunden auf eine Antwort)
german.TestButtonOk=Verbunden mit %1 - der Server antwortete und die Anmeldung wurde akzeptiert.
german.TestButtonFailed=Verbindung zu %1 nicht möglich.
german.TestReasonUnknownHost=Unbekannter Hostname - auf Tippfehler prüfen.
german.TestReasonTimeout=Keine Antwort (Zeitüberschreitung) - läuft der Server und ist er erreichbar?
german.TestReasonAuth=Der Server hat das Passwort (oder den Benutzer) abgelehnt. Prüfen Sie die Anmeldedaten und versuchen Sie es erneut.
german.TestReasonUnexpected=Der Server antwortete, lehnte die Anfrage aber ab. Prüfen Sie Benutzername, Passwort und ob dieser Port wirklich ein Redis-Server ist.
german.TestReasonRefused=Verbindung abgelehnt - läuft Redis auf diesem Port?
german.TestCouldNotRun=Der Verbindungstest konnte nicht ausgeführt werden (PowerShell nicht verfügbar?). Der Installer prüft die Verbindung am Ende trotzdem.
german.ConnPort=Port:
german.FinishedText=RedisExcel ist installiert in:%n%2%n%nWenn das Add-in nicht automatisch geladen wird, aktivieren Sie es in Excel:%nDatei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wählen Sie RedisExcel64-packed.xll (oder RedisExcel-packed.xll für 32-Bit-Office).%n%nProbieren Sie dann =RedisUDFGet("preco_btc") in einer Zelle.
german.GroupConfig=Konfiguration:
german.GroupExcel=Excel:
german.HostRequired=Bitte geben Sie einen Redis-Host ein (z. B. localhost).
german.IconUninstall={#AppName} deinstallieren
german.ModeAdvanced=Erweitert - Datenbank, SSL, Schreibmodus und Update-Tuning festlegen
german.ModePageDesc=Einfach fragt nur nach der Serveradresse. Erweitert erlaubt Datenbank, SSL/TLS, den Schreibmodus und das Echtzeit-Update-Tuning. Sie können RedisExcel.json später jederzeit bearbeiten (oder Validate-Config.cmd ausführen).
german.ModePageSubtitle=Einfache oder erweiterte Einrichtung?
german.ModePageTitle=Einrichtungsmodus
german.ModeSimple=Einfach - nur die Redis-Adresse (empfohlen)
german.OptionsOpenExcel=Am Ende eine Beispiel-Arbeitsmappe RedisExcel öffnen
german.OptionsPageDesc=Der Installer testet die Verbindung am Ende. Optional wird eine Beispiel-Excel-Arbeitsmappe geöffnet, die die Funktion zeigt.
german.OptionsPageSubtitle=Fast fertig
german.OptionsPageTitle=Zusätzliche Prüfungen
german.ServerDockerRedis=Ja - Docker: Redis (redis:latest)
german.ServerDockerValkey=Ja - Docker: Valkey (valkey/valkey:latest)
german.ServerFailed=%1 konnte nicht automatisch eingerichtet werden (Code %2).%nSie können später einen Server installieren und RedisExcel.json darauf verweisen.
german.ServerMemurai=Ja - Memurai (nativer Windows-Dienst, Admin erforderlich)
german.ServerNameDockerRedis=Redis (Docker)
german.ServerNameDockerValkey=Valkey (Docker)
german.ServerNameGeneric=der Server
german.ServerNameMemurai=Memurai
german.ServerNameWslRedis=Redis (WSL)
german.ServerNameWslValkey=Valkey (WSL)
german.ServerNone=Nein - ich habe bereits einen Server (oder richte ihn später ein)
german.ServerPageDesc=RedisExcel benötigt einen Redis- (oder Valkey-)Server. Wählen Sie eine Option, wenn Sie noch keinen haben. Zuerst wird der Server installiert, dann RedisExcel, und am Ende wird die Verbindung getestet. Docker ist am einfachsten.
german.ServerPageSubtitle=Möchten Sie jetzt einen Redis-kompatiblen Server installieren?
german.ServerPageTitle=Redis-Server
german.ServerWslRedis=Ja - WSL2: Redis (apt install redis-server)
german.ServerWslValkey=Ja - WSL2: Valkey (apt install valkey-server)
german.SummaryAdvanced=Erweitert
german.SummaryIntro=Diese Werte werden in RedisExcel.json im Installationsordner gespeichert.
german.SummaryLabel=Verbindung
german.SummaryMode=Modus
german.SummaryPageSubtitle=Prüfen Sie die Verbindung, die in RedisExcel.json geschrieben wird
german.SummaryPageTitle=Ihre Einstellungen
german.SummarySimple=Einfach
german.TaskAutoLoadDesc=Das Add-in registrieren, damit Excel es automatisch lädt (empfohlen)
german.TaskConfigDesc=RedisExcel.json im Installationsordner erstellen (Host-Einstellungen)
german.TaskNlogDesc=NLog.config erstellen (Logdatei neben dem Add-in schreiben)
german.TestFailed=Installiert, aber der Redis unter %1 hat nicht auf einen PING geantwortet.%nPrüfen Sie, ob der Server läuft (docker ps), und passen Sie RedisExcel.json an, wenn er auf einem anderen Host/Port ist.
german.TestOk=Alles bereit! Redis hat unter %1 geantwortet.%n%nSo verwenden Sie es: Excel öffnen. Wenn das Add-in nicht automatisch geladen wird, gehen Sie zu Datei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wählen Sie die .xll in %2.
german.UninstKeepConfig=Ihre RedisExcel.json-Einstellungen behalten?
german.UninstServerFound=%1  [gefunden]
german.UninstServerFailed=%1 konnte nicht automatisch entfernt werden (Code %2). Entfernen Sie es bei Bedarf manuell.
german.UninstServersDesc=Diese Server wurden auf diesem Computer gefunden. Markieren Sie die, die mit RedisExcel entfernt werden sollen.%nStandardmäßig wird NICHTS entfernt.
german.UninstServersSubtitle=Auch Redis-kompatible Server entfernen?
german.UninstServersTitle=Server
german.UpgradeConfigKept=Eine vorhandene RedisExcel.json wurde gefunden und BEHALTEN (Ihre Einstellungen wurden nicht überschrieben): sie ist mit dieser Version kompatibel, es wurde also nichts migriert.%n%nValidate-Config.cmd prüft die Datei jederzeit.
german.UpgradeConfigMigrated=Die RedisExcel.json in diesem Ordner verwendete das ALTE Format (vor 1.6), das dieses Add-In ablehnt - jede Zelle hätte #CONFIG gezeigt.%n%nSie wurde als %1 gespeichert und eine neue RedisExcel.json mit Ihren Angaben geschrieben.

; ---- japanese ----
japanese.AdvPageDesc=これは RedisExcel.json の全体です。各キーには既定値と短いコメントがあらかじめ入力されています。アドインは // コメントを受け付けるのでそのまま残せます。ホスト、エイリアス、書き込みモード、ConflationMs、AsyncWrites をここで変更できます。
japanese.AdvPageTitle=詳細接続
japanese.ConfigWriteFailed=%1 に書き込めませんでした。サンプル（RedisExcel.json）から手動で作成してください。
japanese.ConnHost=ホスト（または host:port）:
japanese.ConnPageDesc=既定値（localhost:6379）でよければ「次へ」。後で RedisExcel.json を編集して変更できます。
japanese.ConnPageSubtitle=RedisExcel はどこに接続しますか？
japanese.ConnPageTitle=Redis 接続
japanese.ConnPassword=パスワード（任意）:
japanese.TestButton=接続をテスト
japanese.TestButtonTesting=テスト中...（最大 5 秒待機します）
japanese.TestButtonOk=%1 に接続しました - サーバーが応答し、ログインが承認されました。
japanese.TestButtonFailed=%1 に接続できませんでした。
japanese.TestReasonUnknownHost=ホスト名が不明です - 入力ミスがないか確認してください。
japanese.TestReasonTimeout=応答なし（タイムアウト） - サーバーは起動して到達可能ですか？
japanese.TestReasonAuth=サーバーがパスワード（またはユーザー名）を拒否しました。認証情報を確認して再試行してください。
japanese.TestReasonUnexpected=サーバーは応答しましたが、要求を拒否しました。ユーザー名、パスワード、およびこのポートが本当に Redis サーバーか確認してください。
japanese.TestReasonRefused=接続が拒否されました - そのポートで Redis は起動していますか？
japanese.TestCouldNotRun=接続テストを実行できませんでした（PowerShell が使用できない可能性があります）。インストーラーは最後に接続を確認します。
japanese.ConnPort=ポート:
japanese.FinishedText=RedisExcel は次にインストールされました:%n%2%n%nアドインが自動で読み込まれない場合は Excel で有効化してください:%nファイル > オプション > アドイン > 設定... > 参照... で RedisExcel64-packed.xll（32 ビット Office なら RedisExcel-packed.xll）を選択します。%n%nその後、セルに =RedisUDFGet("preco_btc") を入力してみてください。
japanese.GroupConfig=構成:
japanese.GroupExcel=Excel:
japanese.HostRequired=Redis ホストを入力してください（例: localhost）。
japanese.IconUninstall={#AppName} をアンインストール
japanese.ModeAdvanced=詳細 - データベース、SSL、書き込みモード、更新チューニングを設定
japanese.ModePageDesc=簡易はサーバーアドレスのみ尋ねます。詳細ではデータベース、SSL/TLS、書き込みモード、リアルタイム更新チューニングを設定できます。RedisExcel.json は後でいつでも編集できます（Validate-Config.cmd の実行でも可）。
japanese.ModePageSubtitle=簡易または詳細セットアップ？
japanese.ModePageTitle=セットアップモード
japanese.ModeSimple=簡易 - Redis アドレスのみ（推奨）
japanese.OptionsOpenExcel=最後にサンプルの RedisExcel ブックを開く
japanese.OptionsPageDesc=インストーラーは終了時に接続をテストします。必要に応じて、動作を示すサンプル Excel ブックを開きます。
japanese.OptionsPageSubtitle=まもなく完了
japanese.OptionsPageTitle=追加チェック
japanese.ServerDockerRedis=はい - Docker: Redis（redis:latest）
japanese.ServerDockerValkey=はい - Docker: Valkey（valkey/valkey:latest）
japanese.ServerFailed=%1 を自動でセットアップできませんでした（終了コード %2）。%n後でサーバーをインストールし、RedisExcel.json で指定できます。
japanese.ServerMemurai=はい - Memurai（Windows ネイティブ サービス、管理者権限が必要）
japanese.ServerNameDockerRedis=Redis（Docker）
japanese.ServerNameDockerValkey=Valkey（Docker）
japanese.ServerNameGeneric=サーバー
japanese.ServerNameMemurai=Memurai
japanese.ServerNameWslRedis=Redis（WSL）
japanese.ServerNameWslValkey=Valkey（WSL）
japanese.ServerNone=いいえ - 既にサーバーがある（または後で用意する）
japanese.ServerPageDesc=RedisExcel には Redis（または Valkey）サーバーが必要です。まだ無い場合はオプションを選択してください。先にサーバーをインストールし、次に RedisExcel をインストールし、最後に接続をテストします。Docker が最も簡単です。
japanese.ServerPageSubtitle=今、Redis 互換サーバーをインストールしますか？
japanese.ServerPageTitle=Redis サーバー
japanese.ServerWslRedis=はい - WSL2: Redis（apt install redis-server）
japanese.ServerWslValkey=はい - WSL2: Valkey（apt install valkey-server）
japanese.SummaryAdvanced=詳細
japanese.SummaryIntro=これらの値はインストールフォルダーの RedisExcel.json に保存されます。
japanese.SummaryLabel=接続
japanese.SummaryMode=モード
japanese.SummaryPageSubtitle=RedisExcel.json に書き込まれる接続を確認してください
japanese.SummaryPageTitle=設定内容
japanese.SummarySimple=簡易
japanese.TaskAutoLoadDesc=Excel が自動で読み込むようアドインを登録（推奨）
japanese.TaskConfigDesc=インストールフォルダーに RedisExcel.json を作成（ホスト設定）
japanese.TaskNlogDesc=NLog.config を作成（アドインの隣にログ ファイルを書き込む）
japanese.TestFailed=インストール済みですが、%1 の Redis が PING に応答しませんでした。%nサーバーが起動しているか確認し（docker ps）、別のホスト/ポートなら RedisExcel.json を調整してください。
japanese.TestOk=準備完了！ %1 の Redis が応答しました。%n%n使い方: Excel を開きます。アドインが自動で読み込まれない場合は、ファイル > オプション > アドイン > 設定... > 参照... で %2 の .xll を選択してください。
japanese.UninstKeepConfig=RedisExcel.json の設定を保持しますか？
japanese.UninstServerFound=%1  [検出]
japanese.UninstServerFailed=%1 ??????????????(????? %2)?????????????????
japanese.UninstServersDesc=これらのサーバーがこのコンピューターで見つかりました。RedisExcel と一緒に削除するものをチェックしてください。%n既定では何も削除されません。
japanese.UninstServersSubtitle=Redis 互換サーバーも削除しますか？
japanese.UninstServersTitle=サーバー
japanese.UpgradeConfigKept=既存の RedisExcel.json が見つかったため保持しました(設定は上書きされていません)。このバージョンと互換性があるため、移行は行っていません。%n%nValidate-Config.cmd でいつでも確認できます。
japanese.UpgradeConfigMigrated=このフォルダーの RedisExcel.json は、このアドインが受け付けない古い(1.6 以前の)形式でした。すべてのセルに #CONFIG エラーが表示されるところでした。%n%n元のファイルは %1 として保存し、入力された設定で新しい RedisExcel.json を作成しました。
[Tasks]
Name: "addinconfig"; Description: "{cm:TaskConfigDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "nlogconfig";  Description: "{cm:TaskNlogDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "excelautoload"; Description: "{cm:TaskAutoLoadDesc}"; GroupDescription: "{cm:GroupExcel}"

[Icons]
; The shortcut FILE name must be language-independent: using the localized
; {cm:IconUninstall} made a reinstall in another language create a SECOND
; shortcut and leave the old one orphaned (audit J1/C2). The AppName is stable.
Name: "{group}\{#AppName}"; Filename: "{uninstallexe}"; Comment: "{cm:IconUninstall}"

[Files]
Source: "{#BuildDir}\RedisExcel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\RedisExcel-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
; StackExchange.Redis and its own dependencies, ALSO unpacked next to the XLL.
; The XLL already carries them packed, but the installer's "Test connection"
; button drives the REAL client from PowerShell (test-connection.ps1), so the
; check exercises exactly the connection-string parser, TLS/ACL/cluster support
; and client version the add-in itself uses - a raw PING proves none of that.
; ~1.5 MB, and harmless next to the packed XLL (Excel-DNA loads from the pack).
Source: "{#LibDir}\StackExchange.Redis.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\Pipelines.Sockets.Unofficial.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.IO.Pipelines.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.Threading.Tasks.Extensions.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.Runtime.CompilerServices.Unsafe.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.Memory.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.Buffers.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\System.Numerics.Vectors.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\Microsoft.Extensions.Logging.Abstractions.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#LibDir}\Microsoft.Bcl.AsyncInterfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\NLog.config"; DestDir: "{app}"; Flags: onlyifdoesntexist; Tasks: nlogconfig
; Helper scripts: server setup, connection test, end-of-install verification,
; server detection/removal for the uninstaller, and the config validator.
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\test-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\verify-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
; The console header (ASCII art) printed by the helpers above. Dot-sourced with
; a try/catch, so its absence degrades to a plain header instead of failing.
Source: "{#RepoRoot}\installer\banner.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\detect-servers.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\uninstall-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\validate-config.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\read-config-host.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
; The full JSON Schema (the same one the add-in embeds) and a double-click tool.
Source: "{#RepoRoot}\RedisExcel.schema.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\Validate-Config.cmd"; DestDir: "{app}"; Flags: ignoreversion
; A temporary copy of the server helpers, so the server can be installed
; BEFORE the add-in files are extracted (see PrepareToInstall), and of the host
; reader, which the WIZARD needs before any file is installed. banner.ps1 comes
; along because the helpers dot-source it from their own folder.
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "{#RepoRoot}\installer\test-connection.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "{#RepoRoot}\installer\read-config-host.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "{#RepoRoot}\installer\banner.ps1"; DestDir: "{tmp}"; Flags: dontcopy

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
  TestButton: TNewButton;
  TestResultLabel: TNewStaticText;
  AdvSeeded: Boolean;
  ServerInstallCode: Integer;
  TestOk: Boolean;
  InstallCompleted: Boolean;
  InstallKeptConfig: Boolean;
  InstallMigratedConfig: Boolean;
  StaleEntriesRemoved: Integer;
  ConfigBackupPath: String;
  ExistingHostValue: String;
  ExistingHostLoaded: Boolean;
  ConnHostPrefilled: Boolean;
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

// ---- Remove the auto-load entry of a PREVIOUS RedisExcel install ----------
// An upgrade into another folder (or a copy registered by hand) leaves its own
// OPEN value behind, so Excel loads TWO copies of the add-in: every RedisExcel
// function is then duplicated in the function list, and each copy reads its own
// folder's RedisExcel.json - one of them usually shows #CONFIG or answers from
// the wrong server. Only entries that point at a RedisExcel XLL and are NOT the
// one being installed are removed; other add-ins in the key are left alone.
function XllNameIsRedisExcel(const Value: String): Boolean;
var
  FileName: String;
begin
  FileName := ExtractFileName(RemoveQuotes(Trim(Value)));
  Result := (CompareText(FileName, 'RedisExcel64-packed.xll') = 0)
    or (CompareText(FileName, 'RedisExcel-packed.xll') = 0)
    or (CompareText(FileName, 'RedisExcel64.xll') = 0)
    or (CompareText(FileName, 'RedisExcel.xll') = 0);
end;

// TRUE when the value names a RedisExcel add-in other than KeepPath.
function IsStaleRedisExcelEntry(const Value, KeepPath: String): Boolean;
var
  Target: String;
begin
  Result := False;
  if (Trim(Value) = '') or (not XllNameIsRedisExcel(Value)) then Exit;
  Target := RemoveQuotes(Trim(Value));
  if Pos('\', Target) = 0 then
    // A bare file name: keep it only when it is this install's own XLL.
    Result := CompareText(Target, ExtractFileName(KeepPath)) <> 0
  else
    Result := CompareText(Target, KeepPath) <> 0;
end;

// Removes OPEN / OPEN1..OPENn values that would load ANOTHER RedisExcel copy.
// Returns how many entries were removed (logged by the caller).
function RemoveStaleExcelEntries(const KeepPath: String): Integer;
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
begin
  Result := 0;
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and
       IsStaleRedisExcelEntry(Cur, KeepPath) then
    begin
      if RegDeleteValue(HKCU, BaseKey, 'OPEN') then
      begin
        Result := Result + 1;
        Log('Stale add-in removed: ' + Keys[I] + '\Excel\Options\OPEN = ' + Cur);
      end;
    end;
    N := 1;
    while N < 100 do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if RegQueryStringValue(HKCU, BaseKey, ValName, Cur) and
         IsStaleRedisExcelEntry(Cur, KeepPath) then
      begin
        if RegDeleteValue(HKCU, BaseKey, ValName) then
        begin
          Result := Result + 1;
          Log('Stale add-in removed: ' + Keys[I] + '\Excel\Options\' + ValName + ' = ' + Cur);
        end;
      end;
      N := N + 1;
    end;
  end;
end;

// ---- Remove a stale "Add-in Manager" entry --------------------------------
// Excel still lists add-ins registered through the LEGACY dialog: the value NAME
// is the label the user sees and the data is the path. A renamed or moved old
// copy leaves a label pointing at a file that no longer exists, which shows up
// as a second, dead "RedisExcel"/"redisd" line in File > Options > Add-ins.
// Only values whose path is a RedisExcel XLL that no longer exists are removed.
function RemoveStaleAddInManagerEntries(const KeepPath: String): Integer;
var
  Keys: TArrayOfString;
  Names: TArrayOfString;
  I, J: Integer;
  BaseKey, Cur, Target: String;
begin
  Result := 0;
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Add-in Manager';
    if RegGetValueNames(HKCU, BaseKey, Names) then
      for J := 0 to GetArrayLength(Names) - 1 do
        if RegQueryStringValue(HKCU, BaseKey, Names[J], Cur) then
        begin
          Target := RemoveQuotes(Trim(Cur));
          if XllNameIsRedisExcel(Cur) and (not FileExists(Target)) and
             (CompareText(Target, KeepPath) <> 0) then
          begin
            if RegDeleteValue(HKCU, BaseKey, Names[J]) then
            begin
              Result := Result + 1;
              Log('Stale Add-in Manager entry removed: ' + Names[J] + ' = ' + Cur);
            end;
          end;
        end;
  end;
end;

// ---- Wizard pages ---------------------------------------------------------
// Forward declarations: the wizard (created first) installs this handler, and
// it reads the values through helpers defined further down.
procedure TestButtonClick(Sender: TObject); forward;
function TestCurrentConnection(): Boolean; forward;
function HostString(): String; forward;
function PasswordValue(): String; forward;

function ParamOr(Param, Default: String): String;
begin
  Result := ExpandConstant('{param:' + Param + '|' + Default + '}');
end;

// ---- "Test connection" button (host page) ---------------------------------
// PINGs the address currently in the fields, WITHOUT leaving the page and
// without touching RedisExcel.json: nobody should have to walk through the
// whole wizard just to learn that the address is wrong. The wizard already
// ships test-connection.ps1 (the end-of-install check extracts it), so the same
// authenticated probe is reused; a password travels through a private temp file
// instead of the command line.
//
// Only the helper's STABLE exit code is used for the message: its own text
// comes through a console redirect in the OEM codepage, so accented characters
// would arrive mangled in a label.
function TestCurrentConnection(): Boolean;
var
  Script, OutFile, PwFile, Args, HostPort, Pw, Reason: String;
  Code: Integer;
begin
  Result := False;
  HostPort := HostString();
  if Pos(',', HostPort) > 0 then HostPort := Copy(HostPort, 1, Pos(',', HostPort) - 1);

  ExtractTemporaryFile('test-connection.ps1');
  Script := ExpandConstant('{tmp}\test-connection.ps1');
  OutFile := ExpandConstant('{tmp}\redisexcel-test.txt');
  DeleteFile(OutFile);

  Pw := PasswordValue();
  PwFile := '';
  if Pw <> '' then
  begin
    PwFile := ExpandConstant('{tmp}\redisexcel-test-pw.txt');
    if not SaveStringToFile(PwFile, Pw, False) then
      PwFile := '';
  end;

  Args := '-ServerHost "' + HostPort + '"';
  if PwFile <> '' then
    Args := Args + ' -PasswordFile "' + PwFile + '"';
  // -Quiet: this output is redirected to a file and only the exit code is
  // used, so no banner is printed (it would just be captured as noise).
  Args := Args + ' -Quiet';

  if not Exec('cmd.exe',
          '/c ""powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "' + Script + '" ' +
          Args + ' > "' + OutFile + '" 2>&1"',
          ExpandConstant('{tmp}'), SW_HIDE, ewWaitUntilTerminated, Code) then
  begin
    TestResultLabel.Caption := ExpandConstant('{cm:TestCouldNotRun}');
    Exit;
  end;
  if PwFile <> '' then DeleteFile(PwFile);

  if Code = 0 then
  begin
    TestResultLabel.Caption := FmtMessage(ExpandConstant('{cm:TestButtonOk}'), [HostPort]);
    Result := True;
    Exit;
  end;

  case Code of
    2: Reason := ExpandConstant('{cm:TestReasonUnknownHost}');
    3: Reason := ExpandConstant('{cm:TestReasonTimeout}');
    5: Reason := ExpandConstant('{cm:TestReasonAuth}');
    6: Reason := ExpandConstant('{cm:TestReasonUnexpected}');
  else
    Reason := ExpandConstant('{cm:TestReasonRefused}');
  end;
  TestResultLabel.Caption := FmtMessage(ExpandConstant('{cm:TestButtonFailed}'), [HostPort])
    + ' ' + Reason;
  Log('TestCurrentConnection: host=' + HostPort + ' exit=' + IntToStr(Code));
end;

procedure TestButtonClick(Sender: TObject);
begin
  // Immediate feedback: a dead host can take the whole 5s TCP timeout.
  TestResultLabel.Font.Color := clGrayText;
  TestResultLabel.Caption := ExpandConstant('{cm:TestButtonTesting}');
  TestResultLabel.Refresh;
  TestButton.Enabled := False;
  try
    if TestCurrentConnection() then
      TestResultLabel.Font.Color := clGreen
    else
      TestResultLabel.Font.Color := clRed;
  finally
    TestButton.Enabled := True;
  end;
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

  // "Test connection" on the host page: the answer is shown next to the button
  // without leaving the page. Next still behaves exactly as before.
  TestButton := TNewButton.Create(ConnPage);
  TestButton.Parent := ConnPage.Surface;
  TestButton.Width := ScaleX(150);
  TestButton.Height := ScaleY(25);
  TestButton.Left := 0;
  TestButton.Top := ConnPage.SurfaceHeight - TestButton.Height - ScaleY(4);
  TestButton.Caption := ExpandConstant('{cm:TestButton}');
  TestButton.OnClick := @TestButtonClick;

  // The result sits ABOVE the button and spans the page width: a failure is two
  // lines (address + reason) and a 1-line label next to the button clipped it.
  TestResultLabel := TNewStaticText.Create(ConnPage);
  TestResultLabel.Parent := ConnPage.Surface;
  TestResultLabel.Left := 0;
  TestResultLabel.Top := TestButton.Top - ScaleY(42);
  TestResultLabel.Width := ConnPage.SurfaceWidth;
  TestResultLabel.Height := ScaleY(38);
  TestResultLabel.AutoSize := False;
  TestResultLabel.WordWrap := True;
  TestResultLabel.Caption := '';

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
// The XLL this installation actually leaves on disk (the non-matching one is
// deleted in CurStepChanged), used to tell "our" auto-load entry from an older
// copy's during the stale-entry cleanup.
function CurrentXllPath(): String;
begin
  if Is64BitOffice then
    Result := ExpandConstant('{app}\RedisExcel64-packed.xll')
  else
    Result := ExpandConstant('{app}\RedisExcel-packed.xll');
end;

// Reads a field from the page that is visible in the current mode.
// The address an existing RedisExcel.json points at, in either shape (v1.6
// Default/Servers or a pre-1.6 RTD.host/UDF.host), read by the bundled helper so
// an UPGRADE keeps the user's server instead of silently writing localhost into
// the replacement file. Cached: the wizard asks for it several times.
function ExistingConfigHost(): String;
var
  Script, Target, OutFile: String;
  Lines: TArrayOfString;
  Code: Integer;
begin
  Result := '';
  if ExistingHostLoaded then
  begin
    Result := ExistingHostValue;
    Exit;
  end;
  ExistingHostLoaded := True;

  Target := AddBackslash(WizardDirValue()) + 'RedisExcel.json';
  if not FileExists(Target) then Exit;

  ExtractTemporaryFile('read-config-host.ps1');
  Script := ExpandConstant('{tmp}\read-config-host.ps1');
  OutFile := ExpandConstant('{tmp}\redisexcel-host.txt');
  DeleteFile(OutFile);
  if Exec('powershell.exe',
          '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" -Config "' + Target +
          '" -Out "' + OutFile + '"',
          ExpandConstant('{tmp}'), SW_HIDE, ewWaitUntilTerminated, Code) then
  begin
    if LoadStringsFromFile(OutFile, Lines) and (GetArrayLength(Lines) > 0) then
    begin
      if Copy(Lines[0], 1, 5) = 'HOST=' then
        ExistingHostValue := Copy(Lines[0], 6, Length(Lines[0]));
    end;
  end;
  Result := ExistingHostValue;
end;

// 0 = the add-in accepts this file, 1 = it REJECTS it (pre-1.6 shape or a wrong
// value: every cell would show #CONFIG), 2 = the check could not run (no schema
// next to the helper, powershell missing): keep the file, never guess.
function ValidateExistingConfig(const ConfigPath: String): Integer;
var
  Script: String;
  Code: Integer;
begin
  Script := ExpandConstant('{app}\scripts\validate-config.ps1');
  if not FileExists(Script) then
  begin
    Result := 2;
    Exit;
  end;
  if Exec('powershell.exe',
          '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" -Config "' + ConfigPath + '"',
          ExpandConstant('{app}\scripts'), SW_HIDE, ewWaitUntilTerminated, Code) then
    Result := Code
  else
    Result := 2;
end;

function Host(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[0])
  else
  begin
    Result := ParamOr('HOST', '');
    if Result = '' then Result := ExistingConfigHost();
    if Result = '' then Result := 'localhost';
  end;
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
// Escapes a value for a JSON string literal (backslash and double quote), so a
// password like "pa\ss" does not produce invalid JSON (audit H2/D3).
function JsonEscape(S: String): String;
begin
  StringChangeEx(S, '\', '\\', True);
  StringChangeEx(S, '"', '\"', True);
  Result := S;
end;

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
  HostLine := JsonEscape(HostString());
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
  H := JsonEscape(H); // the template embeds H in a JSON string literal
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

  // Fail closed on an unknown /SERVER value. install-server.ps1 validates -Kind
  // with a ValidateSet, so a typo aborts its own run, and the "server failed"
  // message below is suppressed in silent mode: a /VERYSILENT install with
  // /SERVER=docker (typo) would otherwise finish with exit code 0 and no server,
  // which automation cannot detect. Raising here aborts BEFORE any file is
  // installed and writes the reason to the setup log. CompareText keeps the
  // case-insensitive acceptance the ValidateSet always had.
  if (CompareText(Kind, 'docker-redis') <> 0) and (CompareText(Kind, 'docker-valkey') <> 0)
     and (CompareText(Kind, 'wsl-redis') <> 0) and (CompareText(Kind, 'wsl-valkey') <> 0)
     and (CompareText(Kind, 'memurai') <> 0) then
    RaiseException('Unknown /SERVER value "' + Kind +
      '": use none, docker-redis, docker-valkey, wsl-redis, wsl-valkey or memurai. Nothing was installed.');

  ExtractTemporaryFile('install-server.ps1');
  // The helper dot-sources banner.ps1 from its own folder ({tmp} here); extract
  // it too so the visible console shows the RedisExcel header, not an
  // unidentified window printing technical lines.
  ExtractTemporaryFile('banner.ps1');
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
      begin
        // An existing config is KEPT when the add-in accepts it (the user's tuned
        // settings win). A pre-1.6 file - or any file with a wrong value - is
        // rejected WHOLESALE by the add-in, so keeping it would leave the upgrade
        // showing #CONFIG in every cell: back it up and write a working file
        // instead. The wizard already carries the old address (ExistingConfigHost),
        // so the replacement does not move the user to localhost.
        if ValidateExistingConfig(ConfigPath) = 1 then
        begin
          ConfigBackupPath := ConfigPath + '.pre-1.6.bak';
          if FileCopy(ConfigPath, ConfigBackupPath, True) then
          begin
            Lines := ConfigJson();
            if SaveStringToFile(ConfigPath, Lines, False) then
              InstallMigratedConfig := True
            else
            begin
              // The write failed: put the original back and keep the old behavior.
              FileCopy(ConfigBackupPath, ConfigPath, True);
              InstallKeptConfig := True;
            end;
          end
          else
            // Could not back it up: never overwrite something we cannot preserve.
            InstallKeptConfig := True;
        end
        else
          // Accepted (0), or undecidable (2): keep it and mention the validator,
          // which is what the finish message does for InstallKeptConfig.
          InstallKeptConfig := True;
      end;
    end;

    if WizardIsTaskSelected('excelautoload') then
    begin
      if Is64BitOffice then
        RegisterExcelAutoLoad(ExpandConstant('{app}\RedisExcel64-packed.xll'))
      else
        RegisterExcelAutoLoad(ExpandConstant('{app}\RedisExcel-packed.xll'));
    end;

    // An earlier copy (another folder, or one registered by hand) makes Excel
    // load TWO add-ins: the functions are duplicated and each copy reads its own
    // folder's RedisExcel.json - one of them usually shows #CONFIG. Drop the
    // other copy's auto-load entry and its dead legacy dialog entry; this
    // install's own entry (just written above) is never touched.
    StaleEntriesRemoved := RemoveStaleExcelEntries(CurrentXllPath());
    StaleEntriesRemoved := StaleEntriesRemoved + RemoveStaleAddInManagerEntries(CurrentXllPath());
    if StaleEntriesRemoved > 0 then
      Log('Stale older installs unregistered: ' + IntToStr(StaleEntriesRemoved));

    InstallCompleted := True;
  end;
end;

// ---- Finish: TEST the configuration, then optionally open a sample sheet ---
procedure DeinitializeSetup();
var
  HostPort, Args, PwFile, Msg: String;
begin
  // Only on a successful install (never during uninstall/cancel).
  if not InstallCompleted then Exit;

  HostPort := HostString();
  if Pos(',', HostPort) > 0 then HostPort := Copy(HostPort, 1, Pos(',', HostPort) - 1);

  // The end-of-install checks must AUTH: a server configured with a password
  // answers -NOAUTH to a bare PING, and the wizard would then report a healthy
  // install as "did not answer a PING" (and the sample sheet's status cell as
  // FALHOU). The secret travels through a private temp file rather than the
  // argument list: a password containing a quote or a space cannot be passed as a
  // single argument safely, and a process command line is readable by other local
  // users. (A silent install still receives /PASSWORD= on ITS command line, which
  // the setup log records - that one is the operator's own choice.)
  PwFile := '';
  if PasswordValue() <> '' then
  begin
    PwFile := ExpandConstant('{tmp}\redisexcel-pw.txt');
    if not SaveStringToFile(PwFile, PasswordValue(), False) then
      PwFile := ''; // unreadable temp: fall back to the unauthenticated probe
  end;

  Args := '-ServerHost "' + HostPort + '"';
  if PwFile <> '' then
    Args := Args + ' -PasswordFile "' + PwFile + '"';
  TestOk := (RunScript('test-connection.ps1', Args) = 0);

  // The migration note is INDEPENDENT of the connection test: a migrated config
  // can point at a server that is down right now, and the user must still learn
  // that the old file was backed up and replaced (its own box, before the result).
  if InstallMigratedConfig then
  begin
    Log('ConfigMigrated: backup=' + ConfigBackupPath);
    if not WizardSilent() then
      MsgBox(FmtMessage(ExpandConstant('{cm:UpgradeConfigMigrated}'), [ConfigBackupPath]),
             mbInformation, MB_OK);
  end;

  if not TestOk then
    Msg := FmtMessage(ExpandConstant('{cm:TestFailed}'), [HostPort])
  else if InstallKeptConfig then
    // An existing, compatible RedisExcel.json was kept verbatim.
    Msg := FmtMessage(ExpandConstant('{cm:UpgradeConfigKept}'), [HostPort, ExpandConstant('{app}')])
  else
    Msg := FmtMessage(ExpandConstant('{cm:TestOk}'), [HostPort, ExpandConstant('{app}')]);

  if not WizardSilent() then
    MsgBox(Msg, mbInformation, MB_OK);

  if OpenExcelChecked() then
  begin
    Args := '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '" -OpenExcel -Quiet';
    if PwFile <> '' then
      Args := Args + ' -PasswordFile "' + PwFile + '"';
    RunScript('verify-connection.ps1', Args);
  end;

  if PwFile <> '' then
    DeleteFile(PwFile);
end;

procedure CurPageChanged(CurPageID: Integer);
var
  Lines, ModeText, Existing: String;
begin
  // Upgrade: start the connection page from the address the EXISTING config
  // points at, so the replacement file cannot silently move to localhost.
  if (CurPageID = ConnPage.ID) and (not WizardSilent()) and (not ConnHostPrefilled) then
  begin
    ConnHostPrefilled := True;
    if ParamOr('HOST', '') = '' then
    begin
      Existing := ExistingConfigHost();
      if Existing <> '' then
      begin
        if Pos(':', Existing) > 0 then
        begin
          ConnPage.Values[0] := Copy(Existing, 1, Pos(':', Existing) - 1);
          ConnPage.Values[1] := Copy(Existing, Pos(':', Existing) + 1, Length(Existing));
        end
        else
          ConnPage.Values[0] := Existing;
      end;
    end;
  end;

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
    // a silent run drive the removal; import it regardless of the mode. NOTE:
    // this runs during UNINSTALL, where WizardSilent() is illegal - use
    // UninstallSilent() (audit H2/D2).
    ImportRemoveServersParam();
    if UninstallSilent() then
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
    if (Code <> 0) and (not UninstallSilent()) then
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
    // Excel-DNA's ComServer registers the CLSID -> InProcServer32 mapping under
    // HKCU too (RedisRtd.cs AutoOpen); the uninstaller deletes files directly so
    // that path never runs, leaving an orphaned key pointing at the deleted XLL
    // (audit J1/C1). Remove it explicitly. Excel-DNA (64-bit Office) writes it in
    // the 64-BIT registry view, and the Inno uninstaller is 32-bit, so the plain
    // HKCU root would look in the wrong view - use HKCU64. The braces are built
    // with Chr(123)/Chr(125) so Inno does not read '{...}' as a constant.
    RegDeleteKeyIncludingSubkeys(HKCU64,
      'Software\Classes\CLSID\' + Chr(123) + '7067CF61-FE6B-3C78-A522-F5602308B418' + Chr(125));
    // Snapshot the ticked servers while the page still exists, then remove them.
    // This MUST run at usUninstall: the [UninstallDelete] entry removes
    // {app}\scripts (and the helper) BEFORE usPostUninstall, so calling it there
    // silently no-ops (audit F2).
    CaptureTickedServers();
    RunUninstallServers();
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    // Never block an unattended uninstall with a dialog: keep the config by
    // default in a silent run (audit H2/D4).
    if FileExists(ExpandConstant('{app}\RedisExcel.json')) and (not UninstallSilent()) then
      if MsgBox(ExpandConstant('{cm:UninstKeepConfig}'), mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\RedisExcel.json'));
  end;
end;

