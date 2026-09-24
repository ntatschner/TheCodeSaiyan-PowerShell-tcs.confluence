# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-09-24

### Breaking

- Requires tcs.core 0.3.0 or later.
- `Set-ConfluenceContext` no longer creates the global variable `$ConfluenceContext` (which held the
  Authorization header in plain text). The context is module-private; use the new
  `Get-ConfluenceContext` to read it. The token is kept in memory as a `PSCredential` only.
- `Get-ConfluenceSpaces` is renamed `Get-ConfluenceSpace` and `Remove-DuplicateConfluencePages` is
  renamed `Remove-DuplicateConfluencePage` (singular nouns). The old names still work as aliases.
- `New-ConfluencePage -Force` only updates the page with exactly the same title. It no longer falls
  back to overwriting a page whose title merely starts with the title, or any page under the same
  parent.
- `Invoke-ConfluenceRequest -Method` only accepts GET, POST, PUT and DELETE.
- `New-ConfluenceContentTOC`: `-HorizontalList` now sets the macro type to `flat` and
  `-IncludeSectionNumbers` sets `outline`; `-BulletPointStyle` maps to the CSS list styles Confluence
  expects (`Mixed` uses the Confluence default). `-HeadersFromLevel`/`-HeadersToLevel` are integers.
- `Update-ConfluencePage` and `New-ConfluencePage` support `-WhatIf`/`-Confirm`;
  `Remove-DuplicateConfluencePage` now uses `ConfirmImpact = 'High'` and reports failed deletions
  as errors instead of warnings.
- The module loads through `RootModule` instead of `NestedModules`.

### Added

- `Get-ConfluenceContext`.
- `Set-ConfluenceContext -Credential` (PSCredential) and `-WhatIf` support.
- `Remove-ConfluencePage` accepts page objects from the pipeline (`id` property).
- Comment-based help with every parameter and examples for every exported function.
- Pester tests for every function (REST calls mocked), `tests/Module.Tests.ps1` and a CI smoke test.
- CI: Pester 5.7.1 on Ubuntu, macOS, Windows (PowerShell 7) and Windows PowerShell 5.1;
  PSScriptAnalyzer 1.23.0 lint that fails on warnings.
- Repository standards: README, CONTRIBUTING, SECURITY, `.editorconfig`, `.gitattributes`,
  `.gitignore`, CODEOWNERS, Dependabot, issue and pull request templates.
- Every exported command reports anonymous usage telemetry through tcs.core
  (`Invoke-TelemetryCollection` Start/End: command name, duration, success and exception type).
  It honours `TCS_TELEMETRY_OPTOUT` and `Set-ModuleConfig -Telemetry $false`, and never changes
  a command's output, errors or `-WhatIf` behaviour. `tests/Telemetry.Tests.ps1` checks that
  every exported command reports telemetry.
- `about_tcs.confluence` help topic (`Get-Help about_tcs.confluence`).
- Publishing workflows from tcs-shared-workflows: `create-version-tag.yml` (tags `v<version>`
  when `ModuleVersion` increases on `main`), `generate-docs.yml` (PlatyPS help in `docs/` and
  `en-GB/`) and `publish-to-psgallery.yml`, with `.github/PUBLISHING.md` describing the release
  steps.

### Changed

- The CI workflow is named `CI Validate` (was `CI - Validate Module`), the name the tag and docs
  workflows wait for.

### Fixed

- Windows PowerShell 5.1 support, which the manifest claimed: `Invoke-ConfluenceRequest` and
  `New-HtmlTable` failed to load (PowerShell 7 ternary operator `? :`), so every REST command failed;
  `-SkipHttpErrorCheck` (PowerShell 7 only) was used for the REST calls; `Join-ConfluenceContent`
  used `[ValidateNotNullOrWhiteSpace()]` (PowerShell 7.4+); `New-ConfluenceContentLink -TextBlock`
  passed a script block to `-replace` (PowerShell 6+).
- `Remove-ConfluencePage` always failed (it passed `pages/<id>` to `-Resource`) and was not exported.
- Pagination never followed Confluence's `_links.next`; relative v1 links are now resolved against
  `/wiki`, and links to other hosts are not followed (so the credential is not sent elsewhere).
- Resolving `-SpaceKey` to a space ID sent `?<key>=` instead of `?keys=<key>` (a hashtable key named
  `keys` hid the `Keys` property).
- DELETE requests failed when Confluence returned an empty body (204 No Content).
- When the context was missing or the URL invalid, `Invoke-ConfluenceRequest` still ran its request
  loop (a `return` in `begin` does not stop `process`), reusing the previous call's endpoint.
- `Get-HtmlTableRowData` wrote an error for every table whose header had more than one column, and
  failed with `-ErrorAction Stop` (`IsNullOrWhiteSpace` received two arguments);
  `-RowHeaderColumnName` is now validated before parsing.
- `New-HtmlTable -MergeColumns` without `-MergeRows` threw instead of warning.
- `New-ConfluenceContentTable`: empty collections were rejected; numbers in the first column were
  dropped; first-cell formatting tags were never closed and header tags were closed in the wrong
  order; dates were rendered as nested tables; validation errors did not stop the output.
- `New-ConfluenceContentCodeBlock`: code containing `]]>` broke the macro.
- `ConvertTo-ConfluenceHTML`: fenced code blocks spanning several lines were not converted, and
  their content was changed by the bold/italic rules; code is now HTML-encoded.
- `Get-ConfluencePageContent -ContentType` was mandatory, so its default was never used.
- The module-load telemetry passed the placeholder URI `https://NOTYETDEFINED.com` and the update
  check wrote its result to the pipeline on import.
- Removed `Config.ps1`, which would write `Config.psd1` into the module folder (fails for AllUsers
  installs) and was not used.

## [0.0.31] and earlier

Releases 0.0.6 to 0.0.31 (2025) added the REST client and content builders; their notes were kept
in the module manifest and are summarised here:

- 0.0.30: fixed syntax errors in Remove-DuplicateConfluencePages that prevented module import.
- 0.0.26 - 0.0.27: legacy path and full URL parsing in Invoke-ConfluenceRequest; the page and space
  functions use Invoke-ConfluenceRequest; fixed Set-ConfluenceContext header creation.
- 0.0.16 - 0.0.25: Invoke-ConfluenceRequest gained -Resource/-ApiVersion/-Id/-RawPath, wildcard
  (CQL) search fallback, spaceKey to spaceId resolution and robust endpoint building.
- 0.0.7 - 0.0.12: New-ConfluencePage conflict handling; added Remove-DuplicateConfluencePages.
- 0.0.6: added Get-ConfluencePageContent.
