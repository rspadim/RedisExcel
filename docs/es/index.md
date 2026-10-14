---
description: Complemento de Excel (XLL) que conecta Excel a Redis mediante RTD y funciones de hoja de cálculo
---

# RedisExcel

![RedisExcel — Excel y Redis, vía UDF, RTD y Pub/Sub](../assets/redisexcel.jpg){: .rdx-hero }

**Complemento de Excel (XLL) que conecta Excel a [Redis](https://redis.io/)**
mediante RTD (Real-Time Data) y UDF (funciones definidas por el usuario):
escribes una fórmula, el dato de Redis llega a la celda y — en las funciones de
streaming — se mantiene en vivo. Y también al revés: puedes publicar en Redis
directamente desde una celda.

Admite polling (`GET`, `HGET`, `HGETALL`) y Pub/Sub (`SUB`, `PSUB`), operaciones
de claves, hashes, listas, conjuntos y contadores, conversión JSON ↔ matriz y
Sentinel/Cluster de punta a punta. Construido con Excel-DNA sobre .NET Framework
4.8, usando StackExchange.Redis, NLog y Newtonsoft.Json.

[![Licencia: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/RedisExcel/blob/main/LICENSE)
[![Última versión](https://img.shields.io/github/v/release/rspadim/RedisExcel?sort=semver)](https://github.com/rspadim/RedisExcel/releases/latest)


Requisitos
----------

- **Windows** con Microsoft Office (Excel), 64 o 32 bits.
- **.NET Framework 4.8**, incluido en Windows 10 y Windows 11.
- Un **servidor compatible con Redis** — Redis, Valkey o Memurai. El instalador
  puede levantar uno por ti (Docker/WSL); Docker es opcional y nunca obligatorio.
- Sin cuenta, registro ni servicio en la nube: el complemento habla directamente
  con el servidor que configures.


Qué hace
--------

| Área | Cobertura |
| --- | --- |
| Celdas en vivo (RTD) | `GET`, `HGET`, `HGETALL` por polling y `SUB`/`PSUB` en tiempo real, con reintento automático y coalescing para mantener la hoja fluida en feeds rápidos |
| Funciones de hoja (UDF) | lectura y escritura de claves, hashes, listas, conjuntos, contadores y TTL, `SCAN` por patrón y conversión JSON ↔ matriz |
| Escritura controlada | 24 funciones `...NonVolatile`: escriben una vez (al entrar en la celda o cuando cambia el argumento), sin repetir en cada recálculo |
| Pub/Sub | `ChannelPublish` (texto y JSON), `ChannelLatest`, `ChannelUnsubscribe` y `PubSubChannelsInfo` |
| Alta disponibilidad | Sentinel y Cluster nativos vía StackExchange.Redis: sigue MOVED/ASK y vuelve a suscribirse tras un failover |
| Diagnóstico | contadores RTD, `RedisUDFConnectionCount`, `RedisUDFUpdateAvailable` y `RedisCheckConfigFile` (valida la configuración en la propia hoja) |


Descarga
--------

Descarga la última versión en la
**[página de releases](https://github.com/rspadim/RedisExcel/releases/latest)**:

| Archivo | Para qué sirve |
| --- | --- |
| `RedisExcelSetup-<versión>.exe` | Instalador de un clic (por usuario; asistente en 10 idiomas) |
| `RedisExcel64-packed.xll` | Complemento para Office de 64 bits |
| `RedisExcel-packed.xll` | Complemento para Office de 32 bits |
| `RedisExcel.dll` | Ensamblado gestionado del complemento (va junto al XLL) |
| `NLog.config` | Configuración del log |
| `RedisExcel.json` | Configuración inicial (ejemplo) |

¿Todavía no tienes un servidor Redis? El instalador puede levantar uno por ti
(Redis o Valkey vía Docker/WSL, o Memurai) y prueba la conexión al final.


Instalación
-----------

1. Ejecuta `RedisExcelSetup-<versión>.exe`. El asistente copia el complemento en
   `C:\RedisExcel`, conserva solo el XLL de la arquitectura de tu Office,
   escribe un `RedisExcel.json` inicial y puede probar la conexión y abrir una
   hoja de ejemplo.
2. Abre Excel. Si el complemento no se carga solo, habilítalo en
   **Archivo → Opciones → Complementos → Administrar: Complementos de Excel →
   Examinar…** y elige `RedisExcel64-packed.xll` (`RedisExcel-packed.xll` en
   Office de 32 bits).
3. Apunta `RedisExcel.json` a tu servidor — el valor por defecto es
   `localhost:6379`; puedes usar alias y cadenas de conexión completas, incluso
   con contraseña (que nunca aparece en los errores ni en el log).

!!! tip "¿Todavía sin servidor?"
    La guía **“No Redis server yet? Run one on Windows”** del
    [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md)
    explica el paso a paso (Docker, Memurai o WSL2). Todos escuchan en
    `localhost:6379` por defecto, que es lo que espera la configuración inicial.

!!! warning "Aviso de SmartScreen"
    Windows puede avisar sobre el XLL al cargarlo por primera vez: haz clic en
    **Más información → Ejecutar de todas formas**. Si un archivo descargado
    queda bloqueado, haz clic derecho → **Propiedades** → **Desbloquear** antes
    de cargar el complemento.

El asistente del instalador está disponible en 10 idiomas; dentro de Excel el
complemento siempre está en inglés, para que los resultados de las fórmulas no
cambien con el idioma de Windows.


Ejemplos
--------

```excel
=RedisUDFSet("precio_btc", A1)                        ' escribe una clave con el valor de la celda A1
=RedisUDFGet("precio_btc")                            ' lee la clave (sin límite de caracteres)
=RedisUDFHashGetAll("cartera")                        ' el hash completo como matriz de 2 columnas
=RedisUDFChannelPublish("alertas", "ALTA")            ' publica en un canal
=RedisUDFChannelLatest("alertas")                     ' último mensaje del canal, en vivo
=RTD("RedisRtd", , "SUB", "alertas")                  ' celda en vivo vía RTD
=RedisUDFJSONToMatrix(RedisUDFChannelLatest("rango")) ' payload JSON → matriz de la hoja
```

Y desde fuera de Excel, con Python puedes leer, suscribirte y publicar:

```python
import redis

r = redis.StrictRedis(host="localhost", port=6379, db=0)

# Python -> Excel: escribe una clave y publica en un canal en vivo
r.set("precio_btc", "67000.50")   # Excel lo lee con  =RedisUDFGet("precio_btc")
r.publish("alertas", "ALTA")      # aparece en  =RedisUDFChannelLatest("alertas")

# Excel -> Python: lee lo que escribió la hoja con  =RedisUDFSet("precio_btc", ...)
print(r.get("precio_btc"))

# Excel -> Python: suscríbete a lo que publica la hoja con  =RedisUDFChannelPublish("alertas", ...)
pubsub = r.pubsub()
pubsub.subscribe("alertas")
for message in pubsub.listen():
    print(message)
```


Documentación
-------------

- **Guía completa** (instalación, RTD, UDF, configuración, pruebas):
  [Readme.md](https://github.com/rspadim/RedisExcel/blob/main/Readme.md)
- **Historial de cambios:**
  [CHANGELOG.md](https://github.com/rspadim/RedisExcel/blob/main/CHANGELOG.md)
- **Configuración:** ejemplo de
  [`RedisExcel.json`](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.json),
  [schema](https://github.com/rspadim/RedisExcel/blob/main/RedisExcel.schema.json)
  y [NLog.config](https://github.com/rspadim/RedisExcel/blob/main/NLog.config)
- **Instalador:**
  [installer/README.md](https://github.com/rspadim/RedisExcel/blob/main/installer/README.md)


Privacidad
----------

RedisExcel no recopila ningún dato del usuario y no nos envía nada. La única
petición que hace por su cuenta es una comprobación en segundo plano contra la
API de releases de GitHub para saber si existe una versión más nueva (desactívala
con `"UpdateCheck": false` en `RedisExcel.json`); los fallos se ignoran. Escribe
un log local (`RedisExcel.log`) junto a los archivos del complemento. El tráfico
de Redis va solo a los servidores que configures, y las credenciales de las
cadenas de conexión se enmascaran en los errores y en el log.


Soporte y contribución
----------------------

Las preguntas, los informes de errores y los pull requests son bienvenidos en
**[github.com/rspadim/RedisExcel](https://github.com/rspadim/RedisExcel/issues)**.
Al abrir un error, incluye la versión instalada, la arquitectura de tu Office y
el mensaje de error (de la celda y/o de `RedisExcel.log`).

**Las traducciones son bienvenidas.** ¿Quieres mantener esta página en tu
idioma? Abre un pull request tocando `docs/` (el contenido del sitio) — hoy hay
portugués, inglés y español.


Licencia
--------

RedisExcel se distribuye bajo la **Licencia MIT** — consulta el
[LICENSE](https://github.com/rspadim/RedisExcel/blob/main/LICENSE).
Componentes de terceros incluidos: Excel-DNA (Zlib), StackExchange.Redis (MIT),
Newtonsoft.Json (MIT), NLog (BSD-3-Clause) y NJsonSchema/Namotion.Reflection
(MIT).

> RedisExcel es un proyecto independiente y no está afiliado ni respaldado por
> Redis Ltd.
