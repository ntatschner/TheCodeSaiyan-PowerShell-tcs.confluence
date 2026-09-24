# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-09-24

### Breaking

- `Get-ConfluencePage` and `Get-ConfluencePageContent` write the page objects to the pipeline
  instead of an object with `Results` and `MultiPage` properties. Replace
  `(Get-ConfluencePage ...).Results` with `Get-ConfluencePage ...`. The pages now pipe straight
  into `Remove-ConfluencePage`.
- The content builders escape their text and attributes (`&`, `<`, `>` and quotes). Text that
  relied on being inserted as markup must use the new `-Raw` switch (`New-ConfluenceContentHeader`,
  `New-ConfluenceContentTable`, `New-ConfluenceContentInfo`).
- `New-ConfluenceContentInfo` returns the storage-format panel macro
  (`<ac:structured-macro ac:name="info|tip|note|warning">` with a rich-text body) instead of an AUI
  `div`, which Confluence did not render as a panel. `-Type error` produces the warning panel
  (Confluence has no error panel). The content is wrapped in a paragraph unless `-Raw` is given.
- `New-ConfluenceContentInternalLink -PageId` returns a storage-format page link
  (`<ac:link><ri:page ri:content-title=... ri:space-key=.../></ac:link>`); use `-AsUrl` for the
  previous `<a href>` output. `-InternalLinkURL` still returns an `<a href>` link. The new default
  parameter set takes `-PageTitle` and `-SpaceKey`.
- `New-ConfluencePageLayout` returns the layout markup as a string (was an object with
  `LayoutType`, `LayoutXml` and `ContentSections`). The section parameters are ordinary parameters;
  passing the wrong number of columns for the layout is an error.
- `New-ConfluenceContentDivider` returns `<hr/>` (self-closed) and `<p>&#160;</p>`;
  `Join-ConfluenceContent -Separator Space|Tab` uses `&#160;`/`&#8195;` instead of `&nbsp;`/`&emsp;`
  (same rendering, but well-formed XML).
- `New-ConfluenceContentTable`, `New-ConfluenceContentLink -TextBlock`: only addresses with a
  scheme (`http://`, `https://`, `mailto:`) become links; file names such as `report.pdf`, bare
  domains and bare e-mail addresses stay text.
- `ConvertTo-ConfluenceHTML` escapes text (HTML in the Markdown is shown as text) and turns a
  Markdown table into one table with one cell per column; the `|---|` row is skipped and marks the
  header row.
- `New-HtmlTable` and `Get-HtmlTableRowData` are removed. They duplicated
  `New-ConfluenceContentTable`, were not used by the module and `New-HtmlTable` clashed with the
  PSWriteHTML command of the same name.
- `Invoke-ConfluenceRequest -Resource` uses the API version set with
  `Set-ConfluenceContext -ApiVersion` when `-ApiVersion` is not passed (it always used v2). The page,
  space, label and attachment commands still use the API version they are written for.
- A space key that cannot be resolved to a space ID is reported as an error and no request is sent
  (the request was previously sent without a space filter).
- `New-ConfluencePage -Force` only updates a page in the same space, and reports an error instead
  of choosing when several pages with the title exist outside the parent.

### Added

- `Search-ConfluenceContent -Cql` (CQL search through `/wiki/rest/api/search`).
- `Get-ConfluencePageChild` (child pages; `-Recurse` for every descendant).
- `Get-ConfluencePageLabel`, `Add-ConfluencePageLabel`, `Remove-ConfluencePageLabel`.
- `Get-ConfluenceAttachment`, `Add-ConfluenceAttachment` (multipart upload that also works on
  Windows PowerShell 5.1; uploading an existing file name adds a version).
- `Clear-ConfluenceContext`.
- `New-ConfluenceContentStatus` (status lozenge), `New-ConfluenceContentExpand` (expand macro) and
  `New-ConfluenceContentJiraIssue` (Jira issue or JQL macro).
- `Get-ConfluencePage -All`, `Get-ConfluencePageContent -All` and `Invoke-ConfluenceRequest -All`
  read every result page; a warning is written when `-MaxQueryPages` stops while more results
  exist.
- `Update-ConfluencePage` without `-Version` reads the page and sends its version number plus one.
- `Remove-ConfluencePage -Purge` moves the page to the trash and purges it.
- `New-ConfluencePage`/`Update-ConfluencePage -SpaceId` (the old name `-SpaceKey` is an alias);
  a space key is resolved to the space ID.
- `New-ConfluenceContentTable` accepts hashtables and ordered dictionaries as rows.
- `New-ConfluencePageLayout -Section` builds a layout with several sections.
- Requests that get 429 Too Many Requests are retried up to four times, honouring `Retry-After`.
- Space key to space ID lookups are cached for the session (cleared by `Set-ConfluenceContext` and
  `Clear-ConfluenceContext`).
- Tests that the documented Confluence parameters are sent (`space-id`, `content/search?cql=`,
  `expand=body.storage`) and that every builder's output parses as XML.

### Fixed

- The v2 pages space filter was sent as `spaceId`; the documented parameter is `space-id`, so
  space filters were ignored and pages from every space were returned.
- Wildcard and CQL title searches were sent to `/wiki/rest/api/content`, which ignores `cql`; they
  now use `/wiki/rest/api/content/search`. Page searches add `type = page`, and a space filter is
  added to the CQL query. `Get-ConfluencePageContent` requests the body of search results with
  `expand=body.<format>` (the v2 `body-format` parameter does not apply to the v1 search).
- CQL values escaped `"` but not `\`, so a title ending in a backslash could change the query;
  backslashes are escaped first.
- `Get-ConfluencePageContent` removed `& # [ ] { }` and other characters from exact titles although
  the value is URL-encoded, so such pages were never found.
- `Remove-DuplicateConfluencePage` kept the page with the highest version number (an old page that
  was edited often) instead of the newest page, and only looked at the first 75 pages of the space.
  It now asks for the pages with that title across all result pages and sorts by `createdAt`.
- Builders produced invalid storage format (and allowed markup injection) for text containing
  `&` or `<`, and `New-ConfluenceContentDivider` returned an unclosed `<hr>`.
- `New-ConfluenceContentTable` threw for hashtable rows.
- `Invoke-ConfluenceRequest`: about 90 lines of URL repair were replaced by the URL normalised by
  `Set-ConfluenceContext`.

## [0.1.1] - 2026-09-24

### Fixed
- Restored the hand-written `about_tcs.confluence` help topic. The shared docs workflow had replaced it with the PlatyPS placeholder.

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
