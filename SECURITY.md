# Security Policy

## Supported versions

Only the latest released version of tcs.confluence receives security fixes.

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Report them privately through
[GitHub security advisories](https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/security/advisories/new).
Include the affected version, the steps to reproduce and the impact you expect.

You should get a first response within 7 days.

## Scope notes

- `Set-ConfluenceContext` keeps the API token in memory only, as a `PSCredential`, for the current
  PowerShell session. It is never written to disk, never returned by `Get-ConfluenceContext` and
  never written to the verbose, warning or error streams. Anything that suggests otherwise is a
  security issue.
- Requests are only sent over HTTPS, and pagination links that point to a host other than the
  Confluence site in the context are not followed.
- Code already running in the same PowerShell session can read module state; the token is not
  protected against that.
- Telemetry (through tcs.core) never sends user names, machine names, paths, page content or error
  messages.
