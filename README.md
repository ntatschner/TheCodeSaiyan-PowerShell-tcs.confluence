# TheCodeSaiyan PowerShell tcs.confluence Module

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/tcs.confluence.svg?style=flat-square&label=PowerShell%20Gallery)](https://www.powershellgallery.com/packages/tcs.confluence)
[![Build Status](https://img.shields.io/github/actions/workflow/status/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/ci-validate.yml?branch=main&style=flat-square&label=Build)](https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/actions/workflows/ci-validate.yml)

Work with Confluence Cloud from PowerShell: read, create, update and delete pages, search with CQL,
manage labels and attachments through the REST API, and build page bodies in Confluence storage
format (headings, tables, code blocks, links, information panels, status lozenges, expand sections,
Jira issues, layouts and tables of contents). Part of the TheCodeSaiyan tcs suite.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7 on Windows, Linux or macOS
- [tcs.core](https://www.powershellgallery.com/packages/tcs.core) 0.3.0 or later (installed
  automatically from the PowerShell Gallery as a dependency)
- For the REST commands: a Confluence Cloud site, an account e-mail address and an
  [API token](https://id.atlassian.com/manage-profile/security/api-tokens)

## Installation

```powershell
Install-Module -Name tcs.confluence -Scope CurrentUser
```

From source:

```powershell
git clone https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence.git
Install-Module -Name tcs.core -MinimumVersion 0.3.0 -Scope CurrentUser
Import-Module ./TheCodeSaiyan-PowerShell-tcs.confluence/modules/tcs.confluence/tcs.confluence.psd1
```

## Quick start

```powershell
# Connect (the token is kept in memory only, for this session)
Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Credential (Get-Credential)

# Build a page body
$body = Join-ConfluenceContent -Separator HorizontalRule -ContentBlocks @(
    New-ConfluenceContentTOC
    New-ConfluenceContentHeader -Header 'Services' -Level 2
    New-ConfluenceContentTable -TableData (Get-Service | Select-Object -First 5 Name, Status) -HeaderStringFormatting Bold
    New-ConfluenceContentCodeBlock -Content 'Get-Service' -Language powershell
)

# Create the page, or update it if a page with this title exists (-Force)
New-ConfluencePage -SpaceId ENG -ParentId 123456 -Title 'Service report' -Status current -Content $body -Force

# Read pages (page objects are written to the pipeline; -All reads every result page)
Get-ConfluencePage -SpaceKey ENG -Search 'Service*' -All | Select-Object id, title

# Search with CQL, label a page and attach a file
Search-ConfluenceContent -Cql 'type = page AND space = ENG AND lastmodified > now("-7d")'
Add-ConfluencePageLabel -PageId 123456 -Label report
Add-ConfluenceAttachment -PageId 123456 -Path ./report.pdf
```

## Functions

### REST API

| Function | Purpose |
| --- | --- |
| `Set-ConfluenceContext` | Sets the site URL and credential for the session (`-Credential` or `-Username`/`-PersonalAccessToken`) |
| `Get-ConfluenceContext` | Shows the current connection (never the token) |
| `Clear-ConfluenceContext` | Forgets the connection and credential |
| `Invoke-ConfluenceRequest` | Low-level request with endpoint building, CQL search, cached space key resolution, 429 retry and pagination |
| `Get-ConfluencePage` | Gets pages by ID, space or title (wildcards use CQL); `-All` reads every result page |
| `Get-ConfluencePageContent` | Gets pages with their body in storage, view or another format |
| `Get-ConfluencePageChild` | Gets the child pages of a page, or every descendant with `-Recurse` |
| `Search-ConfluenceContent` | Runs a CQL search (`-Cql`) |
| `Get-ConfluenceSpace` | Lists spaces or gets one by ID (alias `Get-ConfluenceSpaces`) |
| `New-ConfluencePage` | Creates a page; with `-Force` updates the page with the same title in the same space (`-WhatIf` supported) |
| `Update-ConfluencePage` | Writes a new version of a page; without `-Version` the next version is used (`-WhatIf` supported) |
| `Remove-ConfluencePage` | Deletes a page, or with `-Purge` deletes it permanently (asks for confirmation) |
| `Remove-DuplicateConfluencePage` | Deletes duplicate pages with the same title and parent, keeping the newest or oldest by creation date (asks for confirmation; alias `Remove-DuplicateConfluencePages`) |
| `Get-ConfluencePageLabel` | Gets the labels of a page |
| `Add-ConfluencePageLabel` | Adds labels to a page |
| `Remove-ConfluencePageLabel` | Removes labels from a page |
| `Get-ConfluenceAttachment` | Gets the attachments of a page |
| `Add-ConfluenceAttachment` | Uploads files to a page (a new version when the file name exists) |

### Content builders (no connection needed)

The builders escape text (`&`, `<`, `>` and quotes), so their output is always well-formed storage
format. Where a builder takes content that is itself storage format (table cells, panel and expand
bodies, headers), pass `-Raw` to insert it unchanged.

| Function | Purpose |
| --- | --- |
| `New-ConfluenceContentHeader` | Heading `h1`-`h6` with optional formatting |
| `New-ConfluenceContentTable` | Table from objects, hashtables or ordered dictionaries, with nested tables and link detection (http, https, mailto) |
| `New-ConfluenceContentCodeBlock` | Code block macro |
| `New-ConfluenceContentTOC` | Table of contents macro |
| `New-ConfluenceContentInfo` | Info, tip, note or warning panel macro (`error` uses the warning panel) |
| `New-ConfluenceContentStatus` | Status lozenge macro |
| `New-ConfluenceContentExpand` | Expand (collapsible section) macro |
| `New-ConfluenceContentJiraIssue` | Jira issue or JQL table macro |
| `New-ConfluenceContentLink` | Link, or links for every http/https/mailto address in a text |
| `New-ConfluenceContentInternalLink` | Storage-format link (`<ac:link>`) to a Confluence page or a heading on it; `-AsUrl`/`-InternalLinkURL` for a plain URL link |
| `New-ConfluenceContentDivider` | Horizontal rule or spacer |
| `New-ConfluencePageLayout` | Page layout with one or more one-, two- or three-column sections |
| `Join-ConfluenceContent` | Joins content blocks with a separator |
| `ConvertTo-ConfluenceHTML` | Converts simple Markdown (headings, emphasis, lists, code blocks, tables) to storage format |

Every function has full help: `Get-Help New-ConfluencePage -Full`.

## Credentials

`Set-ConfluenceContext` stores the credential in memory as a `PSCredential` for the current
session only. Nothing is written to disk, the token is never returned by `Get-ConfluenceContext`
and never written to the verbose, warning or error streams. Only HTTPS URLs are accepted.

Versions before 0.1.0 stored the Authorization header in the global variable
`$ConfluenceContext`; use `Get-ConfluenceContext` instead.

## Configuration

Settings are managed by tcs.core and stored per user at
`<ApplicationData>/PowerShell/Config/tcs.confluence/Module.Config.json` (`%APPDATA%` on Windows,
`~/.config` on Linux/macOS). Change them with `Set-ModuleConfig`:

```powershell
Set-ModuleConfig -ModuleName tcs.confluence -UpdateWarning $false   # no update warnings on import
Set-ModuleConfig -ModuleName tcs.confluence -Telemetry $false       # no telemetry
Set-ModuleConfig -ModuleName tcs.confluence -Reset                  # back to defaults
```

The update check against the PowerShell Gallery runs at most once a day and never blocks import.
Set `TCS_SKIP_UPDATE_CHECK=1` to turn it off (for example in CI). See the
[tcs.core README](https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.core/blob/main/README.md)
for all settings and environment variables.

## Privacy and telemetry

tcs modules send anonymous usage telemetry to help find failing commands. tcs.confluence records
one event when the module is loaded and one each time an exported command runs. Telemetry is on by default and a notice is shown the first
time a module is loaded. Nothing is sent until a telemetry endpoint is configured.

Each event contains: time (UTC), module and command name, module version, duration, success,
the exception **type** on failure, PowerShell version and edition, OS family, PowerShell host
name, and a random installation ID created on first use.

It **never** contains: user names, machine names, file paths, hardware serial numbers, IP-based
identifiers, command arguments, Confluence URLs, page content, credentials or error messages.

Turn it off with `Set-ModuleConfig -ModuleName tcs.confluence -Telemetry $false`, or for all tcs
modules with the environment variable `TCS_TELEMETRY_OPTOUT=1`.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for the coding standards and how to run the tests, and
[SECURITY.md](SECURITY.md) for reporting security issues. Changes are listed in
[CHANGELOG.md](CHANGELOG.md).

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE)
file for details.

## Author

**Nigel Tatschner** - TheCodeSaiyan
- GitHub: [@ntatschner](https://github.com/ntatschner)
