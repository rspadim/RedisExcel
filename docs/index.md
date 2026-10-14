---
description: Add-in do Excel (XLL) que conecta o Excel ao Redis via RTD e funções de planilha
og_image: assets/social-card.png
og_locale: pt_BR
---

# RedisExcel

![RedisExcel - Excel e Redis, via UDF, RTD e Pub/Sub](assets/social-card.png){: .rdx-hero }

**Add-in do Excel (XLL) que conecta o Excel ao [Redis](https://redis.io/)** por
meio de RTD (Real-Time Data) e UDF (funções definidas pelo usuário): você digita
uma fórmula, o dado do Redis chega na célula e — nas funções de streaming — fica
ao vivo. E o caminho inverso também: dá para publicar no Redis direto de uma
célula.

Suporta polling (`GET`, `HGET`, `HGETALL`) e Pub/Sub (`SUB`, `PSUB`), operações
de chaves, hashes, listas, conjuntos e contadores, conversão JSON ↔ matriz e
Sentinel/Cluster de ponta a ponta. Construído com Excel-DNA sobre o .NET
Framework 4.8, usando StackExchange.Redis, NLog e Newtonsoft.Json.

[![Licença: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/RedisExcel/blob/main/LICENSE)
[![Última versão](https://img.shields.io/github/v/release/rspadim/RedisExcel?sort=semver)](https://github.com/rspadim/RedisExcel/releases/latest)


Requisitos
----------

- **Windows** com Microsoft Office (Excel), 64 bits ou 32 bits.
- **.NET Framework 4.8**, que já vem no Windows 10 e no Windows 11.
- Um **servidor compatível com Redis** — Redis, Valkey ou Memurai. O instalador
  consegue subir um para você (Docker/WSL); Docker é opcional e nunca obrigatório.
- Nenhuma conta, cadastro ou serviço em nuvem: o add-in fala direto com o
  servidor que você configurar.


O que ele faz
-------------

| Área | Cobertura |
| --- | --- |
| Células ao vivo (RTD) | `GET`, `HGET`, `HGETALL` por polling e `SUB`/`PSUB` em tempo real, com retry automático e coalescing para manter a planilha fluida em feeds rápidos |
| Funções de planilha (UDF) | leitura e escrita de chaves, hashes, listas, conjuntos, contadores e TTL, `SCAN` por padrão e conversão JSON ↔ matriz |
| Escrita controlada | 24 funções `...NonVolatile`: escrevem uma vez (ao entrar na célula ou quando o argumento muda), sem repetir a cada recálculo |
| Pub/Sub | `ChannelPublish` (texto e JSON), `ChannelLatest`, `ChannelUnsubscribe` e `PubSubChannelsInfo` |
| Alta disponibilidade | Sentinel e Cluster nativos via StackExchange.Redis: segue MOVED/ASK e volta a assinar após failover |
| Diagnóstico | contadores RTD, `RedisUDFConnectionCount`, `RedisUDFUpdateAvailable` e `RedisCheckConfigFile` (valida a configuração na própria planilha) |


Download
--------

Baixe a versão mais recente na
**[página de releases](https://github.com/rspadim/RedisExcel/releases/latest)**:

| Arquivo | Para que serve |
| --- | --- |
| `RedisExcelSetup-<versão>.exe` | Instalador de um clique (por usuário; assistente em 10 idiomas) |
| `RedisExcel64-packed.xll` | Add-in para Office 64 bits |
| `RedisExcel-packed.xll` | Add-in para Office 32 bits |
| `RedisExcel.dll` | Assembly gerenciado do add-in (vai ao lado do XLL) |
| `NLog.config` | Configuração de log |
| `RedisExcel.json` | Configuração inicial (exemplo) |

Ainda não tem um servidor Redis? O instalador consegue subir um para você
(Redis ou Valkey via Docker/WSL, ou Memurai) e testa a conexão no final.


Instalação
----------

1. Execute `RedisExcelSetup-<versão>.exe`. O assistente copia o add-in para
   `C:\RedisExcel`, mantém só o XLL da arquitetura do seu Office, escreve um
   `RedisExcel.json` inicial e pode testar a conexão e abrir uma planilha de
   exemplo.
2. Abra o Excel. Se o add-in não carregar sozinho, habilite em
   **Arquivo → Opções → Suplementos → Gerenciar: Suplementos do Excel →
   Procurar…** e escolha `RedisExcel64-packed.xll` (ou `RedisExcel-packed.xll`
   no Office 32 bits).
3. Aponte para o seu servidor no `RedisExcel.json` — o padrão é
   `localhost:6379`; dá para usar nomes (aliases) e connection strings
   completas, inclusive com senha (que nunca aparece em erros ou no log).

!!! tip "Ainda não tem um servidor?"
    O guia **“No Redis server yet? Run one on Windows”** no
    [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md) tem o
    passo a passo (Docker, Memurai ou WSL2). Todos escutam em `localhost:6379`
    por padrão, que é o que a configuração inicial espera.

!!! warning "Aviso do SmartScreen"
    O Windows pode avisar sobre o XLL na primeira carga: clique em
    **Mais informações → Executar assim mesmo**. Se um arquivo baixado ficar
    bloqueado, clique com o botão direito → **Propriedades** → **Desbloquear**
    antes de carregar o add-in.

O assistente do instalador está disponível em 10 idiomas; dentro do Excel o
add-in é sempre em inglês, para que os resultados das fórmulas não mudem com o
idioma do Windows.


Exemplos
--------

```excel
=RedisUDFSet("preco_btc", A1)                          ' grava uma chave com o valor da célula A1
=RedisUDFGet("preco_btc")                              ' lê a chave (sem limite de caracteres)
=RedisUDFHashGetAll("carteira")                        ' hash inteiro como matriz de 2 colunas
=RedisUDFChannelPublish("alertas", "ALTA")             ' publica em um canal
=RedisUDFChannelLatest("alertas")                      ' última mensagem do canal, ao vivo
=RTD("RedisRtd", , "SUB", "alertas")                   ' célula ao vivo via RTD
=RedisUDFJSONToMatrix(RedisUDFChannelLatest("range"))  ' payload JSON → matriz do Excel
```

E de fora do Excel, com Python dá para ler, assinar e publicar:

```python
import redis

r = redis.StrictRedis(host="localhost", port=6379, db=0)

# Python → Excel: grava uma chave e publica num canal ao vivo
r.set("preco_btc", "67000.50")     # o Excel lê com  =RedisUDFGet("preco_btc")
r.publish("alertas", "ALTA")       # aparece na célula  =RedisUDFChannelLatest("alertas")

# Excel → Python: lê o que a planilha gravou com  =RedisUDFSet("preco_btc", ...)
print(r.get("preco_btc"))

# Excel → Python: assina o que a planilha publica com  =RedisUDFChannelPublish("alertas", ...)
pubsub = r.pubsub()
pubsub.subscribe("alertas")
for message in pubsub.listen():
    print(message)
```


Documentação
------------

- **Guia completo** (instalação, RTD, UDF, configuração, testes):
  [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md)
- **Histórico de mudanças:**
  [CHANGELOG.md](https://github.com/rspadim/RedisExcel/blob/main/CHANGELOG.md)
- **Configuração:** exemplo de
  [`RedisExcel.json`](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.json),
  [schema](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.schema.json)
  e [NLog.config](https://github.com/rspadim/RedisExcel/blob/main/NLog.config)
- **Instalador:**
  [installer/README.md](https://github.com/rspadim/RedisExcel/blob/main/installer/README.md)


Privacidade
-----------

O RedisExcel não coleta nenhum dado do usuário e não envia nada para nós. A única
requisição que ele faz por conta própria é uma checagem em segundo plano na API
de releases do GitHub, para saber se existe versão mais nova (desligue com
`"UpdateCheck": false` no `RedisExcel.json`); falhas são ignoradas. Ele grava um
log local (`RedisExcel.log`) ao lado dos arquivos do add-in. O tráfego Redis vai
apenas para os servidores que você configurar, e credenciais em connection
strings são mascaradas nos erros e no log.


Suporte e contribuição
----------------------

Dúvidas, relatos de bug e pull requests são bem-vindos em
**[github.com/rspadim/RedisExcel](https://github.com/rspadim/RedisExcel/issues)**.
Ao abrir um bug, inclua a versão instalada, a arquitetura do seu Office e a
mensagem de erro (da célula e/ou do `RedisExcel.log`).

**Traduções são bem-vindas.** Quer manter esta página no seu idioma? Abra um
pull request mexendo em `docs/` (o conteúdo do site) — hoje temos português,
inglês e espanhol.


Licença
-------

O RedisExcel é distribuído sob a **Licença MIT** — veja o
[LICENSE](https://github.com/rspadim/RedisExcel/blob/main/LICENSE).
Componentes de terceiros embarcados: Excel-DNA (Zlib), StackExchange.Redis (MIT),
Newtonsoft.Json (MIT), NLog (BSD-3-Clause) e NJsonSchema/Namotion.Reflection
(MIT).

> O RedisExcel é um projeto independente e não é afiliado nem endossado pela
> Redis Ltd.
