# Release Checklist (Definition of Done)

Most of this list is **enforced by CI**, not by memory. Items marked
`(CI)` fail `build.yml` on their own; the rest are human judgement calls.

## Pre-Release — local

- [ ] `ruff check .` is clean `(CI)`
- [ ] `python -m pytest --cov` passes and meets the coverage gate in
      `pyproject.toml` `(CI)`
- [ ] `python build.py --validate` exits 0 — golden suite **and** pytest `(CI)`
- [ ] `python build.py --output output --no-protect` exits 0 `(CI)`
- [ ] `python scripts/audit_excel.py --output output` reports 0 failed checks `(CI)`
- [ ] `python scripts/make_release_manifest.py --output output --dest landing/release.json --check` reports no drift `(CI)`
- [ ] `pre-commit run --all-files` is clean `(CI)`
- [ ] The golden suite has **no failure and no unexplained warning**: every
      warning must name a test that genuinely has no engine implementation
      yet (the QA sheet prints the reason)
- [ ] Every test the engine implements has a golden case that is actually
      executed against it — enforced by `tests/test_golden_corpus.py` `(CI)`

## Pre-Release — version identity

Bump the version **once**, in `src/concrete_lab/__init__.py`. Everything else
must follow, and `tests/test_consistency.py` fails until it does `(CI)`:

- [ ] `pyproject.toml` → `[project].version`
- [ ] `config.yaml` → `project.version` (and `project.release_date`)
- [ ] `CITATION.cff` → `version` and `date-released`
- [ ] `CHANGELOG.md` → a new `## [X.Y.Z] - YYYY-MM-DD` section
- [ ] `landing/release.json` → regenerate with
      `scripts/make_release_manifest.py` (never edit by hand)
- [ ] `README.md` → only if a _claim_ changed (implemented-test count,
      features). The version number itself is not written in the README.

Also update, if the change affects them:

- [ ] `config.yaml` → `sheets.implemented` / `sheets.pending` / `sheets.key_sheets`
      (must equal the live registry and the rendered workbook) `(CI)`
- [ ] `config.yaml` → `errata.total` / `confirmed` / `pending` (must equal
      `validation/errata.yaml`) `(CI)`

## Release

- [ ] Tag created: `vX.Y.Z`, matching `__version__`
- [ ] A GitHub **Release** is created from that tag — this is what triggers
      asset upload and the Pages redeploy
- [ ] CI uploaded exactly three assets from the same build: the workbook
      (`.xlsx`), its manifest (`.json`) and its checksum (`.sha256`)
- [ ] The landing page shows the version, artifact name and SHA-256 **of that
      release** (it reads `release.json`; nothing is hard-coded)
- [ ] Release notes link to the `CHANGELOG.md` section for this version
- [ ] Artifacts are **not** committed to the repository (`output/`, `*.xlsx`
      and `*.sha256` are gitignored)

## Post-Release

- [ ] Dependabot PRs reviewed (GitHub Actions and pip ecosystems)
- [ ] Golden cases extended for anything the release changed
- [ ] `docs/qa/EXCEL_AUDIT_REPORT.json` is a generated file: it is produced by
      `scripts/audit_excel.py` and published as a CI artifact, not committed

## Notes on reproducibility

The workbook embeds its build timestamp, so rebuilding the same commit does
**not** reproduce the same SHA-256. That is expected. The published checksum
always describes the artifact of a specific release, which is why CI derives
`landing/release.json` from the same job that uploads the assets.

## Command Reference

```bash
ruff check .
python -m pytest --cov
python build.py --validate
python build.py --output output --no-protect
python scripts/audit_excel.py --output output
python scripts/make_release_manifest.py --output output --dest landing/release.json
python scripts/make_release_manifest.py --output output --dest landing/release.json --check
pre-commit run --all-files
```
