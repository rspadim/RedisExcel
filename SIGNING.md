# Code signing

RedisExcel releases will be code-signed. This file records the plan; the signing
step is **not wired into the CI yet** because the certificate does not exist at
the time of writing.

## What gets signed

| Artifact | Why |
| --- | --- |
| `RedisExcel-packed.xll`, `RedisExcel64-packed.xll` | the add-ins the user loads; an unsigned XLL is a classic SmartScreen/AV false positive |
| `RedisExcel.dll` | the managed assembly shipped next to the XLLs (manual installs use it); every binary a user receives should be verifiable |
| `RedisExcelSetup-<version>.exe` (Inno Setup) | the installer is the first file a new user runs and the #1 false-positive target |

The packed XLLs embed their own copies of the managed dependencies; the loose
`RedisExcel.dll` is the one a manual install puts next to the XLL.

## Where the step enters the pipeline

1. `build.yml` produces the release payload (the packed XLLs, `RedisExcel.dll`
   and the installer) - sign each binary **after** it is built and **before**
   the artifact upload.
2. Use an RFC 3161 timestamp server (the SignPath action does this by default),
   so the signatures stay valid after the certificate expires.
3. The SignPath **artifact configuration** lists all four PE files: the two
   packed XLLs, `RedisExcel.dll` and `RedisExcelSetup-<version>.exe`.

## Service

Free code signing for open-source projects is provided by the
[SignPath Foundation](https://signpath.org/); the landing page mentions it in the
download section (required for the application). After approval, the remaining
steps (organization, signing policy and artifact configuration on signpath.io, an
API token in GitHub Secrets, then the workflow step) follow the SignPath
documentation at <https://docs.signpath.io/>.

## Never

No script in this repository disables or reconfigures antivirus or Defender
(that is MITRE T1562.001 and any EDR will - correctly - block it). Test tooling
may only **read** the AV state.
