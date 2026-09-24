## Summary

<!-- What does this change and why? -->

## Type of change

- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change (callers need updating)
- [ ] Documentation / CI only

## Checklist

- [ ] Pester tests pass locally (`Invoke-Pester -Path ./modules/tcs.confluence, ./tests`)
- [ ] PSScriptAnalyzer reports no findings with `PSScriptAnalyzerSettings.psd1`
- [ ] New or changed behaviour has Pester tests (REST calls mocked, no real Confluence site)
- [ ] Comment-based help is updated for changed functions
- [ ] `CHANGELOG.md` is updated
- [ ] Works on Windows PowerShell 5.1 and PowerShell 7 (no PS7-only syntax)
