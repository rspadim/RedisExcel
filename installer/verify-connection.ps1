<#
  verify-connection.ps1 - end-of-install check for RedisExcel.

  * Tests that a Redis-compatible server answers on the given host/port.
  * With -OpenExcel, opens a visible Excel workbook that loads the installed XLL
    and shows the connection status with a friendly message.

  Exit codes: 0 = Redis answered (and the sheet opened if asked), 1 = no answer.

  Examples:
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\RedisExcel" -ServerHost "localhost" -Port 6379
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\RedisExcel" -ServerHost "localhost:6379" -OpenExcel
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$AddInDir,
  [string]$ServerHost = 'localhost',
  [int]$Port = 6379,
  [switch]$OpenExcel,
  [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Show([string]$text, [string]$title = 'RedisExcel') {
  if (-not $Quiet) { [System.Windows.Forms.MessageBox]::Show($text, $title) | Out-Null }
  Write-Host $text
}

# Host may already be "host:port" (or a connection string "host:port,password=..").
$target = $ServerHost.Trim()
if ($target -match '^\s*([^,:;]+)(?::(\d+))?') {
  $serverHost = $Matches[1]
  $serverPort = $Port
  if ($Matches[2]) { $serverPort = [int]$Matches[2] }
} else {
  $serverHost = $ServerHost
  $serverPort = $Port
}
$target = "$serverHost`:$serverPort"

function Test-Redis([string]$h, [int]$p) {
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $client.Connect($h, $p)
    $stream = $client.GetStream()
    # Real protocol PING, not just a socket connect.
    $payload = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
    $stream.Write($payload, 0, $payload.Length)
    $buffer = New-Object byte[] 32
    $read = $stream.Read($buffer, 0, $buffer.Length)
    $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $read)
    $client.Close()
    return ($response -match 'PONG')
  } catch {
    return $false
  }
}

$target = "$serverHost`:$serverPort"
Write-Host "[verify] testing Redis at $target ..."
$ok = Test-Redis $serverHost $serverPort

if (-not $ok) {
  Show ("Nao consegui conectar no Redis em $target.`r`n`r`n" +
        "- Confira se o servidor esta rodando (docker ps, ou o servico).`r`n" +
        "- Se ele esta em outro host/porta, ajuste o RedisExcel.json.") "RedisExcel - conexao"
  exit 1
}

Write-Host "[verify] Redis answered PONG at $target"

if (-not $OpenExcel) {
  Show ("Tudo certo!`r`n`r`nO Redis respondeu em $target.`r`n" +
        "Abra o Excel, registre o XLL (File > Options > Add-ins > Go... > Browse...) " +
        "e use, por exemplo, =RedisUDFGet(`"chave`").") "RedisExcel - instalado"
  exit 0
}

# ---- Open a sample sheet in Excel ------------------------------------------
$xll64 = Join-Path $AddInDir 'RedisExcel64-packed.xll'
$xll32 = Join-Path $AddInDir 'RedisExcel-packed.xll'
$xll = if (Test-Path $xll64) { $xll64 } elseif (Test-Path $xll32) { $xll32 } else { $null }
if (-not $xll) {
  Show ("Nao encontrei o XLL em $AddInDir.") "RedisExcel"
  exit 0
}

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $registered = $excel.RegisterXLL($xll)
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'RedisExcel'

  # Nice title block.
  $ws.Range('A1').Value2 = 'RedisExcel instalado com sucesso!'
  $ws.Range('A1').Font.Size = 16
  $ws.Range('A1').Font.Bold = $true
  $ws.Range('A2').Value2 = 'Teste de conexao com o Redis - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  $ws.Range('A2').Font.Italic = $true

  $ws.Range('A4').Value2 = 'Servidor:'
  $ws.Range('B4').Value2 = $target
  $ws.Range('A5').Value2 = 'Status:'
  # ServerTime is a real round trip; if it errors the cell shows Error:.
  $ws.Range('B5').Formula = '=IF(ISERROR(RedisUDFServerTime("' + $target + '")),"FALHOU","OK - Redis respondeu")'
  $ws.Range('A6').Value2 = 'Servidor diz:'
  $ws.Range('B6').Formula = '=RedisUDFServerTime("' + $target + '")'

  $ws.Range('A8').Value2 = 'Teste leitura/escrita (pressione F9 para atualizar):'
  $ws.Range('A9').Value2 = 'Set:'
  $ws.Range('B9').Formula = '=RedisUDFSet("redisexcel:install_test","ola do instalador")'
  $ws.Range('A10').Value2 = 'Get:'
  $ws.Range('B10').Formula = '=RedisUDFGet("redisexcel:install_test")'

  $ws.Range('A12').Value2 = 'Exemplo: =RedisUDFGet("preco_btc", "' + $target + '")'
  $ws.Range('A4:A6').Font.Bold = $true
  $ws.Range('A9:A10').Font.Bold = $true
  $ws.Columns.Item('A').ColumnWidth = 42
  $ws.Columns.Item('B').ColumnWidth = 40
  $ws.Range('B5').Font.Bold = $true

  $excel.Application.Calculate() | Out-Null
  Write-Host "[verify] sample workbook opened (XLL registered: $registered)"
  exit 0
} catch {
  Show ("Nao consegui abrir o Excel: " + $_.Exception.Message) "RedisExcel"
  exit 0
}
