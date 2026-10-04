# Anything DRM Excel Test Patch

## Purpose
Minimal validation patch for Anything v3.8.12 on an authorized corporate Windows/Office environment.

This repository does **not** bypass, decrypt, disable, or alter Fasoo DRM controls. It only extends the wait time for Anything's existing Microsoft Excel COM fallback when Office is already authorized to open the document.

## Verified baseline
- Upstream: chrisryugj/Docufinder
- Baseline tag: v3.8.12
- Baseline commit: 3545a833d686c1db4088a9e446f75a3bc038d6ed
- Test branch: fix/xlsx-wincom-timeout

## Patch
`src-tauri/src/parsers/mod.rs`

- Excel COM fallback timeout: 30 seconds -> 120 seconds
- Log label: `XLS/XLSX` -> `XLS/XLSX(COM)`
- Native XLS/XLSX parser timeout remains unchanged at 15 seconds.

## Validation result — 2026-10-04
A Fasoo-protected XLSX that the authorized Microsoft Excel installation could open was successfully indexed after applying the patch.

Observed result in Anything:
- document: `대기인허가 검토_261001.xlsx`
- content matches: 14
- preview displayed extracted worksheet text
- known text `GATEDRYER HOT AIR GENERATOR` was visible/searchable
- known value `126.72 ton/year` was visible in extracted content

Before the patch, logs showed the COM worker eventually completing with 14 chunks / 5,185 characters, but the outer 30-second timeout had already returned an indexing error.

## Build
Use the GitHub Actions workflow:
`.github/workflows/build-drm-test.yml`

The test workflow:
- builds on `windows-latest`
- disables updater artifact signing for the test installer
- uploads the NSIS installer as a short-lived GitHub Actions artifact
- does not publish a GitHub Release

## Architecture decision
- Anything: local document indexing/search/preview
- FileOrbit: file organization/move/approval/undo
- KSJ AI Bridge: pass selected search evidence to AI

Avoid duplicating Anything's indexing/search engine inside FileOrbit or KSJ Nexus.

## Next engineering step
Create a thin, explicit handoff from a user-selected Anything search result to KSJ AI Bridge. Prefer passing only:
1. file path / document identity
2. matched chunks
3. location hints (sheet/page/cell range when available)
4. query/context

Do not automatically export entire DRM-protected documents.
