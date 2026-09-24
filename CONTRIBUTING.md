# Contributing to tcs.confluence

tcs.confluence is part of the tcs PowerShell suite and depends on
[tcs.core](https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.core) (0.3.0 or later) for
configuration, update checks and telemetry.

## Getting started

Requirements: PowerShell 7.2+ for development, Pester 5.7.1, PSScriptAnalyzer 1.23.0 and tcs.core
0.3.0 or later.

```powershell
Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser -SkipPublisherCheck
Install-Module PSScriptAnalyzer -RequiredVersion 1.23.0 -Scope CurrentUser
Install-Module tcs.core -MinimumVersion 0.3.0 -Scope CurrentUser

# Keep test runs offline and away from your profile
$env:TCS_SKIP_UPDATE_CHECK = '1'
$env:TCS_TELEMETRY_OPTOUT = '1'

Import-Module Pester -RequiredVersion 5.7.1
Invoke-Pester -Path ./modules/tcs.confluence, ./tests
Invoke-ScriptAnalyzer -Path ./modules/tcs.confluence -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
```

## Layout

| Path | Contents |
| --- | --- |
| `modules/tcs.confluence/Public/` | Exported functions, one per file, named after the function |
| `modules/tcs.confluence/Private/` | Internal helpers (not exported) |
| `modules/tcs.confluence/Public/Tests/` | Pester tests, `<Function>.Tests.ps1` |
| `modules/tcs.confluence/Private/Tests/` | Pester tests for the private helpers |
| `tests/` | Module-wide tests (manifest, exports, help, PSScriptAnalyzer) |

Every file in `Public/` must also be listed in `FunctionsToExport` in `tcs.confluence.psd1`;
`tests/Module.Tests.ps1` checks this.

## Standards

- **Compatibility:** code must run on Windows PowerShell 5.1 and PowerShell 7 on Windows,
  Linux and macOS. Avoid PS7-only syntax (`??`, `?:`, `&&`, `ForEach-Object -Parallel`),
  PS7-only parameters (`-SkipHttpErrorCheck` outside the version check in
  `Invoke-ConfluenceHttpRequest`) and .NET Core-only APIs. CI runs the tests on all four.
- **REST calls** go through `Invoke-ConfluenceRequest`, which uses the private
  `Invoke-ConfluenceHttpRequest`. Never call `Invoke-WebRequest`/`Invoke-RestMethod` directly and
  never write the credential or the Authorization header to any stream.
- **Style:** 4-space indentation, `CmdletBinding()` on every function, approved verbs, singular
  nouns, full command names (no aliases). PSScriptAnalyzer runs with `PSScriptAnalyzerSettings.psd1`
  and **warnings fail the build**. Suppress a rule only with a written justification.
- **State-changing functions** (`New-`, `Set-`, `Update-`, `Remove-` against Confluence) support
  `-WhatIf`/`-Confirm`; `Remove-` functions use `ConfirmImpact = 'High'`.
- **Help:** every exported function has comment-based help with a synopsis, description,
  every parameter and at least one example.
- **Tests:** new behaviour and bug fixes come with Pester tests. Tests must not touch the real
  user profile or network: set `TCS_CONFIG_ROOT` to `$TestDrive` and mock `Invoke-WebRequest`
  with `Mock -ModuleName tcs.confluence`.
- **Versioning:** [Semantic Versioning](https://semver.org). Record changes in `CHANGELOG.md`.
