# Code signing

RedisExcel releases will be code-signed. This file records the plan; the signing
step is **not wired into the CI yet** because the certificate does not exist at
the time of writing.

## What gets signed

| Artifact | Why |
| --- | --- |
| `RedisExcel-packed.xll`, `RedisExcel64-packed.xll` | the add-ins the user loads; an unsigned XLL is a classic SmartScreen/AV false positive |
| `RedisExcelSetup-<version>.exe` (Inno Setup) | the installer is the first file a new user runs and the #1 false-positive target |

`RedisExcel.dll` ships next to the XLLs; sign it too if the signing policy
covers loose DLLs (the packed XLLs already embed their dependencies).

## Where the step enters the pipeline

1. `build.yml` packs the XLLs - sign **after** packing.
2. The Inno Setup compile produces `RedisExcelSetup-<version>.exe` - sign
   **after** it exists and **before** the artifact upload.
3. Use an RFC 3161 timestamp server (the SignPath action does this by default),
   so the signatures stay valid after the certificate expires.

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
