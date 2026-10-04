# KSJ Anything Maintenance Runbook

## Principle

Anything remains an independent application.

This repository exists only to preserve and validate the minimal corporate-DRM compatibility patch. Do not turn it into KSJ Nexus or FileOrbit.

## Known-good state

- Anything baseline: v3.8.12
- Patch branch: `fix/xlsx-wincom-timeout`
- Verified behavior: authorized Fasoo-protected XLSX is indexed through the existing Microsoft Excel COM fallback.
- Minimal functional patch: Excel COM fallback timeout 30s -> 120s.
- Native XLS/XLSX parser timeout remains 15s.
- Test installer uses updater artifacts disabled; it does not publish a release.

## Update policy

Do not overwrite the known-good installed build just because upstream publishes a newer version.

For every upstream update:

1. Read upstream release notes and inspect the current XLS/XLSX COM fallback implementation.
2. Check whether upstream already fixed the slow Office-COM fallback problem.
3. If upstream fixed it, test the official build first. Prefer returning to the official build.
4. If not fixed, create a fresh update branch from the new upstream version and re-apply only the minimal COM-timeout change.
5. Build with the dedicated DRM test workflow.
6. Validate with an authorized non-confidential test document first when possible.
7. Validate the required corporate DRM XLSX:
   - indexing completes without final XLS/XLSX(COM) timeout;
   - known internal text is searchable;
   - preview contains expected extracted text;
   - no duplicate/unwanted test folders are indexed.
8. Only after validation replace the installed known-good build.

## Rollback

Keep the last known-good installer until the replacement has passed the DRM validation.

If a new build fails:
- uninstall/reinstall the last known-good Anything build;
- keep user index/settings when the installer supports it;
- do not alter Fasoo DRM, Office security, or company policy to make the test pass.

## Repository rules

- `main`: upstream baseline snapshot. Do not add KSJ product features here.
- `fix/xlsx-wincom-timeout`: verified v3.8.12 DRM compatibility patch.
- Future updates: use a new branch such as `fix/vX.Y.Z-xlsx-wincom-timeout`.
- Do not merge Nexus or FileOrbit code into this repository.
- Do not commit confidential documents, extracted document content, logs containing business text, credentials, or company data.
- Do not publish test installers as public GitHub Releases.
- Do not enable inherited upstream publish/mirror workflows in KSJ test branches.

## Boundary

Allowed:
- authorized Office COM fallback;
- parser orchestration/timeouts;
- local indexing/search/preview;
- normal upstream maintenance.

Not part of this patch:
- DRM bypass/decryption;
- key extraction;
- copying decrypted temporary data;
- disabling Fasoo controls;
- screen/accessibility scraping to defeat copy restrictions.

## Future Nexus connection

Anything stays independent. If KSJ Nexus integration is added later, use an explicit user-selected handoff/connector. Do not make Nexus own Anything's database or silently export entire DRM documents.
