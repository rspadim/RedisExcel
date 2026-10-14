# Code signing

Status: **inactive**. RedisExcel releases are **not signed** and nothing in the
CI signs them. Signing is parked; when a certificate exists it will be applied
**manually** first (outside the pipeline), and only wired into the CI later if
it proves worth it.

## What gets signed (when a certificate exists)

| Artifact | Why |
| --- | --- |
| `RedisExcel-packed.xll`, `RedisExcel64-packed.xll` | the add-ins the user loads; an unsigned XLL is a classic SmartScreen/AV false positive |
| `RedisExcel.dll` | the managed assembly shipped next to the XLLs (manual installs use it); every binary a user receives should be verifiable |
| `RedisExcelSetup-<version>.exe` (Inno Setup) | the installer is the first file a new user runs and the #1 false-positive target |

The packed XLLs embed their own copies of the managed dependencies; the loose
`RedisExcel.dll` is the one a manual install puts next to the XLL.

## Manual procedure (the current plan)

1. Produce the release payload (the packed XLLs, `RedisExcel.dll` and the
   installer) - the `Build & Release` artifacts are enough.
2. Sign the four PE files with an Authenticode certificate and an **RFC 3161
   timestamp** (`signtool sign /fd SHA256 /tr <timestamp-url> /td SHA256
   <file>` or SignPath's own tooling), then `signtool verify /pa` each one.
3. Upload the signed files to the release.

Only after that manual flow is proven should a CI step be considered; when it
is, the landing page must mention the signing service again in its download
section (the SignPath application requires it).

## Service

Free code signing for open-source projects is provided by the
[SignPath Foundation](https://signpath.org/); the service is not active for this
project yet, so the landing page does not advertise signing. If it is activated
later, the remaining steps (organization, signing policy and artifact
configuration on signpath.io, an API token in GitHub Secrets, then the workflow
step) follow the SignPath documentation at <https://docs.signpath.io/>.

## Never

No script in this repository disables or reconfigures antivirus or Defender
(that is MITRE T1562.001 and any EDR will - correctly - block it). Test tooling
may only **read** the AV state.
