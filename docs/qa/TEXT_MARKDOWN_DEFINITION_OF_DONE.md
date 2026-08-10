# Definition of Done: Text & Markdown

This repository enforces a strict quality gate for text and Markdown files.

## Rules

- UTF-8 without BOM
- LF line endings for text files, except Windows script files
- One final newline
- No trailing whitespace
- No merge conflict markers
- Valid YAML and JSON
- Markdown lint passes
- Prettier formatting passes

## Checks

- pre-commit runs on every pull request.
- GitHub Actions job `Text & Markdown Perfection` must pass.
- Auto-fix workflow can open a pull request with mechanical fixes.

## Manual review checklist

- [ ] The intent of the document is clear.
- [ ] Persian text is readable and right-to-left friendly.
- [ ] Code blocks have language tags.
- [ ] Links are intentional and accessible.
- [ ] File names are lowercase and hyphenated where possible.
