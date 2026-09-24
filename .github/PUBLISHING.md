# Publishing tcs.confluence to the PowerShell Gallery

This repository uses GitHub Actions to validate the `tcs.confluence` module, tag new versions,
generate help and publish to the [PowerShell Gallery](https://www.powershellgallery.com/packages/tcs.confluence).
The workflows call the reusable workflows in
[ntatschner/tcs-shared-workflows](https://github.com/ntatschner/tcs-shared-workflows).

## Workflows

| File | Name in the Actions tab | Runs on | What it does |
| --- | --- | --- | --- |
| `.github/workflows/ci-validate.yml` | CI Validate | Pull requests and pushes to `main` that change `modules/`, `tests/`, `.github/` or the analyzer settings | Validates the manifest and import, runs the smoke test, PSScriptAnalyzer and Pester (PowerShell 7 on Ubuntu, Windows and macOS, and Windows PowerShell 5.1) |
| `.github/workflows/create-version-tag.yml` | Create Version Tag | After CI Validate succeeds on `main`, a push to `main` that changes `modules/tcs.confluence/tcs.confluence.psd1`, or manually | Creates and pushes the tag `v<ModuleVersion>` when the manifest version is greater than the latest `v*` tag |
| `.github/workflows/generate-docs.yml` | Generate PowerShell Documentation | After CI Validate succeeds on `main`, pushes and pull requests that change module code, or manually | Generates Markdown help in `docs/` and external help in `modules/tcs.confluence/en-GB/` with PlatyPS and commits them (pull requests only check that help generates) |
| `.github/workflows/publish-to-psgallery.yml` | Publish to PSGallery | A pushed `v*` tag, or manually | Validates the module, checks that the version is not already in the Gallery, publishes it and creates a GitHub release |

`create-version-tag.yml` and `generate-docs.yml` start when a workflow named `CI Validate`
completes, so keep the `name:` in `ci-validate.yml` exactly `CI Validate`.

## Prerequisites

### PowerShell Gallery API key

1. Sign in to the [PowerShell Gallery](https://www.powershellgallery.com/).
2. Open your account name, then **API Keys**, and create a key.
3. Give it the **Push new packages and package versions** scope and a glob pattern that covers
   `tcs.confluence` (for example `tcs.*` or `tcs.confluence`).
4. Copy the key.

### Repository secret

1. In the GitHub repository open **Settings > Secrets and variables > Actions**.
2. Select **New repository secret**.
3. Name: `PSGALLERY_API_KEY`; value: the API key.

No other secret is needed: the tag and docs workflows use the workflow's `GITHUB_TOKEN`.

## Releasing a new version

1. Update `ModuleVersion` in `modules/tcs.confluence/tcs.confluence.psd1` (see
   [Semantic Versioning](https://semver.org)) and add a section for the version to `CHANGELOG.md`.
2. Merge the change to `main`.
3. **Create Version Tag** runs and pushes the tag `v<ModuleVersion>` (for example `v0.1.0`). It
   does nothing when that tag already exists or the version is not greater than the latest tag.
4. **Start the publish manually.** A tag pushed by a workflow with the `GITHUB_TOKEN` does not
   trigger other workflows, so the tag from step 3 does not start **Publish to PSGallery** on
   its own. Open **Actions > Publish to PSGallery > Run workflow**, choose the tag
   `v<ModuleVersion>` under **Use workflow from**, and run it. Running it on the tag also creates
   the GitHub release for that tag.
5. Check the run, then the [Gallery page](https://www.powershellgallery.com/packages/tcs.confluence).

To make step 4 automatic, give `create-version-tag.yml` a personal access token (or GitHub App
token) with `contents: write` on this repository instead of `GITHUB_TOKEN`: store it as a
repository secret and pass it as `repo-token`. A tag pushed with such a token triggers the
publish workflow.

A tag you push yourself (`git tag v0.1.0` then `git push origin v0.1.0`) also triggers
**Publish to PSGallery**, because it is not pushed by the `GITHUB_TOKEN`.

## What the publish workflow checks

- `Test-ModuleManifest` on the manifest.
- PSScriptAnalyzer on `modules/tcs.confluence` (errors fail the run; warnings are listed).
- The module and its `RequiredModules` (tcs.core) install and import.
- The version is not already in the Gallery. If it is, publishing is skipped unless
  **Force publish** is selected when running the workflow manually.

It then publishes with `Publish-PSResource` (falling back to `Publish-Module`) and, for runs on
a `v*` tag, creates the GitHub release `v<ModuleVersion>`.

## Troubleshooting

- **API key rejected:** check that the secret is named exactly `PSGALLERY_API_KEY`, that the key
  has not expired, has push rights and that its glob pattern covers `tcs.confluence`.
- **Version already exists:** increase `ModuleVersion`. The Gallery does not accept the same
  version twice; use **Force publish** only when you know the version is not there yet.
- **Tag not created:** CI Validate must succeed on `main` first, and the manifest version must be
  greater than the latest `v*` tag. Run **Create Version Tag** manually to see its log.
- **Publish did not start after the tag was created:** expected with `GITHUB_TOKEN`; start it
  manually on the tag (step 4 above).

## Security

- Never commit API keys; keep them in repository secrets.
- Rotate the Gallery API key regularly and scope it to the tcs packages only.
