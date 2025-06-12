# RedisRTD

Integração em tempo real entre Redis e Excel via RTD (Real-Time Data) e UDF (User Defined Functions), utilizando C# (.NET Framework 4.8) com [ExcelDna](https://github.com/Excel-DNA/ExcelDna). Esta biblioteca permite a atualização contínua de valores no Excel através de funções RTD, com suporte a polling (GET, HGET, HGETALL) e Pub/Sub (SUB, PSUB), e funções de UDF do Redis.

---

## 🚀 Exemplo: publicar no Redis via Python

```python
import redis

r = redis.StrictRedis(host='localhost', port=6379, db=0)

# Publicar um valor
r.set("preco_btc", "67000.50")

# Enviar mensagem para um canal
r.publish("canal_alerta", "ALTA")
```

---

## 🛠️ Como compilar o RedisRTD

### Requisitos

- Visual Studio 2022 ou superior
- .NET Framework 4.8
- NuGet packages:
  - `ExcelDna.AddIn`
  - `ExcelDna.Integration`
  - `ExcelDna.IntelliSense`
  - `StackExchange.Redis`
  - `NLog`
  - `Newtonsoft.Json`

### Passos

1. Clone o projeto:

```bash
git clone https://github.com/rspadim/RedisExcel.git
```

2. Abra no Visual Studio.

3. Restaure os pacotes NuGet.

4. Compile em modo `Release`. O `.xll` será gerado em `bin\Release`.

---

## 🧩 Como usar no Excel

### Opção 1: Adicionar manualmente

1. Copie os arquivos de `bin\Release` para uma pasta local.

2. No Excel:
    - Vá em `Arquivo > Opções > Suplementos`.
    - Clique em **Ir...**, depois em **Procurar...**, e selecione o `.xll` gerado.

### Opção 2: Registrar COM (`regasm`)

Se estiver usando `.dll` com COM (para RTD):

```bat
cd "C:\Caminho\Para\Release"
regasm RedisExcel.dll /codebase
```

> Execute como administrador no `Developer Command Prompt`.

---

## 🧪 Testando no Excel

### Pub/Sub com canal

```excel
=RTD("RedisRtd", , "SUB", "canal_alerta", "localhost:6379")
```

### Polling de chave

```excel
=RTD("RedisRtd", , "GET", "preco_btc", "localhost:6379")
```

### Polling de hash

```excel
=RTD("RedisRtd", , "HGET", "user:42", "saldo", "localhost:6379")
```

---

## ⚙️ Configurações Internas

Estas funções podem ser usadas para inspecionar o status do RTD server diretamente no Excel:

| Função                         | Descrição                                                |
|--------------------------------|----------------------------------------------------------|
| `RedisRTDConnectionCount()`    | Número de conexões Redis ativas                          |
| `RedisRTDSubscriptionCount()`  | Número de subscrições Redis ativas                       |
| `RedisRTDTopicCount()`         | Total de tópicos RTD registrados                         |
| `RedisRTDChannelCount()`       | Número de canais distintos com subscrição ativa          |
| `RedisRTDDefaultHost()`        | Host Redis padrão configurado                            |
| `RedisRTDExcelUpdateInterval()`| Intervalo de atualização do Excel em milissegundos       |
| `RedisRTDRedisUpdateInterval()`| Intervalo de polling do Redis em milissegundos           |
| `RedisRTDRealTimeUpdates()`    | Indica se está em modo de atualização contínua (boolean) |

---

## 📄 Exemplo de configuração do Redis (opcional)

```json
{
  "UDF": {
    "host": "localhost:6379,password=senha123,defaultDatabase=0",
    "timeout": "800"
  },
  "RTD": {
   "host": "localhost:6379",
   "timeout": 1000,
   "RedisUpdateRateMs": 1000,
   "ExcelUpdateRateMS": 100,
   "RealTimeUpdates": false,
   "UseGetMultiple": true,
  }
}
```

Coloque o arquivo `RedisExcel.json` no mesmo diretório do `.xll` para carregar configurações padrão, ou no diretório do usuário, ou no diretório do excel, ou no diretório c:\windows\.

---

## 🔄 Atualização Automática no Excel

Para forçar recálculo:

- Pressione `F9` manualmente.
- Ou use VBA:

```vba
Dim proximaAtualizacao As Date

Sub AtualizarRTD()
    Sheet1.Calculate
    proximaAtualizacao = Now + TimeValue("00:00:01")
    Application.OnTime proximaAtualizacao, "AtualizarRTD"
End Sub

Sub PararAtualizacao()
    On Error Resume Next
    Application.OnTime proximaAtualizacao, "AtualizarRTD", , False
End Sub
```

---

## 📬 Contato

Dúvidas, sugestões ou problemas? Abra uma *issue* no repositório ou entre em contato pelo e-mail associado ao projeto.
