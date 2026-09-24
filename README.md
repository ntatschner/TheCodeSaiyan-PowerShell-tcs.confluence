# TheCodeSaiyan PowerShell tcs.confluence Module

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/tcs.confluence.svg?style=flat-square&label=PowerShell%20Gallery)](https://www.powershellgallery.com/packages/tcs.confluence)
[![Build Status](https://img.shields.io/github/actions/workflow/status/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/ci-validate.yml?branch=main&style=flat-square&label=Build)](https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/actions/workflows/ci-validate.yml)

Work with Confluence Cloud from PowerShell: read, create, update and delete pages through the REST
API, and build page bodies in Confluence storage format (headings, tables, code blocks, links,
information panels, layouts and tables of contents). Part of the TheCodeSaiyan tcs suite.

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
$space = Get-ConfluenceSpace -Search 'Engineering'
New-ConfluencePage -SpaceKey $space.id -ParentId 123456 -Title 'Service report' -Status current -Content $body -Force

# Read pages
(Get-ConfluencePage -SpaceKey ENG -Search 'Service*').Results | Select-Object id, title
```

## Functions

### REST API

| Function | Purpose |
| --- | --- |
| `Set-ConfluenceContext` | Sets the site URL and credential for the session (`-Credential` or `-Username`/`-PersonalAccessToken`) |
| `Get-ConfluenceContext` | Shows the current connection (never the token) |
| `Invoke-ConfluenceRequest` | Low-level request with endpoint building, CQL search, spaceKey resolution and pagination |
| `Get-ConfluencePage` | Gets pages by ID, space or title (wildcards use CQL) |
| `Get-ConfluencePageContent` | Gets pages with their body in storage, view or another format |
| `Get-ConfluenceSpace` | Lists spaces or gets one by ID (alias `Get-ConfluenceSpaces`) |
| `New-ConfluencePage` | Creates a page; with `-Force` updates the page with the same title (`-WhatIf` supported) |
| `Update-ConfluencePage` | Writes a new version of a page (`-WhatIf` supported) |
| `Remove-ConfluencePage` | Deletes a page (asks for confirmation) |
| `Remove-DuplicateConfluencePage` | Deletes duplicate pages with the same title and parent, keeping one (asks for confirmation; alias `Remove-DuplicateConfluencePages`) |

### Content builders (no connection needed)

| Function | Purpose |
| --- | --- |
| `New-ConfluenceContentHeader` | Heading `h1`-`h6` with optional formatting |
| `New-ConfluenceContentTable` | Table from objects, with nested tables and link detection |
| `New-HtmlTable` | Table from objects with rowspan merging, or rows appended to an existing table |
| `Get-HtmlTableRowData` | Reads the rows of a table (including Confluence macros) back into objects |
| `New-ConfluenceContentCodeBlock` | Code block macro |
| `New-ConfluenceContentTOC` | Table of contents macro |
| `New-ConfluenceContentInfo` | Info, tip, note, warning or error panel |
| `New-ConfluenceContentLink` | Link, or links for every URL in a text |
| `New-ConfluenceContentInternalLink` | Link to a Confluence page or a heading on it |
| `New-ConfluenceContentDivider` | Horizontal rule or spacer |
| `New-ConfluencePageLayout` | One-, two- or three-column page layout |
| `Join-ConfluenceContent` | Joins content blocks with a separator |
| `ConvertTo-ConfluenceHTML` | Converts simple Markdown to HTML |

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
one event when the module is loaded. Telemetry is on by default and a notice is shown the first
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
