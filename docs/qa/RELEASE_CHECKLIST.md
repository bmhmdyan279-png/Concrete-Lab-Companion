# Release Checklist (Definition of Done)

Before creating a new release, verify:

## Pre-Release

- [ ] All 20 golden test cases pass (`pytest tests/`)
- [ ] `validation/formulas_audit.txt` is up to date
- [ ] `docs/qa/EXCEL_AUDIT_REPORT.json` shows 0 critical issues
- [ ] GitHub Actions `Text & Markdown Perfection` is GREEN
- [ ] `build.py` removed from root (only tests/ remains)

## Release

- [ ] Tag created: `vX.Y.Z`
- [ ] Release notes written in CHANGELOG.md
- [ ] Excel file attached as Release Asset (NOT committed to repo)
- [ ] SHA-256 hash published in Release notes
- [ ] Landing page (`landing/index.html`) points to latest release

## Post-Release

- [ ] Dependabot PRs reviewed
- [ ] Golden tests extended for new features
- [ ] Documentation updated

## Command Reference

```bash
# Run all tests
pytest tests/ -v

# Audit Excel (downloads from latest release)
.\scripts\run-phase2.ps1

# Auto-fix formatting
pre-commit run --all-files
```
