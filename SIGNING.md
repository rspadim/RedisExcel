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

Free Windows code signing for open-source projects is being pursued through two
programs; neither is active for this project yet, so the landing page does not
advertise signing.

- **Necessary Code Signing** (<https://sign.necessary.nu/>) - the certificate is
  issued to *Necessary Innovations AB (Sweden)*, so a signed binary shows that
  publisher; identity and address are verified before a token is granted, and the
  service signs established projects (the application covers this project and its
  sibling add-in NATSExcel). Signing is three `osslsigncode` steps, which is what
  makes it usable from CI:

  ```
  osslsigncode extract-data -in app.exe -out tosign.bin
  curl -X POST -H "Authorization: Bearer $SIGNING_TOKEN" \
    --data-binary @tosign.bin https://sign.necessary.nu/windows/sign -o signed.bin
  osslsigncode attach-signature -sigin signed.bin -in app.exe -out app-signed.exe
  ```

- **SignPath Foundation** (<https://signpath.org/>) - the other free lane; the
  certificate is issued to the foundation, and it additionally requires a published
  code-signing policy on the project's home page plus a manually approved signing
  request per release.

### When a token exists

1. Sign manually first, in the release order: the payload (XLLs + dll), then the
   installer built from the signed files.
2. Verify every file (`Get-AuthenticodeSignature`, or `osslsigncode verify` on what
   the service returns): the status must be `Valid` and the publisher must be the
   signing service's, not a local certificate.
3. Only then wire it into the CI as a guarded step - a fork or a PR has no token and
   must not fail - and keep the post-signing verification as the gate: a
   misconfigured signing step must fail the job, never ship a half-signed release.

## Never

No script in this repository disables or reconfigures antivirus or Defender
(that is MITRE T1562.001 and any EDR will - correctly - block it). Test tooling
may only **read** the AV state.
