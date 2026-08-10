<#
.SYNOPSIS
  بستهٔ نهایی کنترل کیفیت متن و مارک‌داون - نسخهٔ پاورشل
.DESCRIPTION
  این اسکریپت تمام فایل‌های پیکربندی (editorconfig, prettier, markdownlint, pre-commit, GitHub Actions, ...) رو ایجاد می‌کنه،
  فایل‌های متنی رو نرمال‌سازی می‌کنه، commit و push انجام میده.
  اجرا در ترمینال پاورشل پایچارم که به گیتهاب متصله.
#>

# سخت‌گیری کامل
$ErrorActionPreference = "Stop"
# Set-StrictMode -Version Latest   # اگر خواستی خطاهای متغیرهای تعریف‌نشده رو بگیری فعال کن

# بررسی وجود git
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "❌ git نصب نیست." -ForegroundColor Red
    exit 1
}

# بررسی اینکه داخل مخزن گیت هستیم
$insideWorkTree = git rev-parse --is-inside-work-tree 2>$null
if ($insideWorkTree -ne 'true') {
    Write-Host "❌ اینجا ریشه مخزن گیت نیست. اول وارد ریپازیتوری شو." -ForegroundColor Red
    exit 1
}

# ساخت پوشه‌های ضروری
New-Item -ItemType Directory -Force -Path .github/workflows, .vscode, docs/qa | Out-Null

# ========== تولید فایل‌های پیکربندی ==========

# .editorconfig
@'
root = true

[*]
charset = utf-8
end_of_line = lf
insert_final_newline = true
trim_trailing_whitespace = true
indent_style = space
indent_size = 2

[*.md]
indent_size = 2

[*.{yml,yaml}]
indent_size = 2

[*.json]
indent_size = 2

[*.sh]
end_of_line = lf

[*.{bat,cmd,ps1}]
end_of_line = crlf
indent_size = 2

[Makefile]
indent_style = tab
'@ | Set-Content -NoNewline -Encoding utf8 .editorconfig

# .gitattributes
@'
* text=auto eol=lf

*.md text eol=lf
*.markdown text eol=lf
*.txt text eol=lf
*.yml text eol=lf
*.yaml text eol=lf
*.json text eol=lf
*.jsonc text eol=lf
*.sh text eol=lf
*.py text eol=lf
*.js text eol=lf
*.ts text eol=lf

*.bat text eol=crlf
*.cmd text eol=crlf
*.ps1 text eol=crlf

*.xlsx binary
*.xlsm binary
*.xls binary
*.png binary
*.jpg binary
*.jpeg binary
*.gif binary
*.pdf binary
*.zip binary
'@ | Set-Content -NoNewline -Encoding utf8 .gitattributes

# .markdownlint.json
@'
{
  "default": true,
  "MD003": { "style": "atx" },
  "MD004": { "style": "dash" },
  "MD007": { "indent": 2 },
  "MD009": { "br_spaces": 2, "strict": true },
  "MD012": { "maximum": 1 },
  "MD013": false,
  "MD022": { "lines_above": 1, "lines_below": 1 },
  "MD024": { "allow_different_nesting": true },
  "MD029": { "style": "one_or_ordered" },
  "MD030": { "ul_single": 1, "ol_single": 1, "ul_multi": 1, "ol_multi": 1 },
  "MD031": true,
  "MD032": true,
  "MD033": false,
  "MD036": false,
  "MD040": false,
  "MD041": false,
  "MD046": { "style": "fenced" },
  "MD047": true,
  "MD048": { "style": "backtick" },
  "MD049": { "style": "underscore" },
  "MD050": { "style": "asterisk" }
}
'@ | Set-Content -NoNewline -Encoding utf8 .markdownlint.json

# .markdownlintignore
@'
.git/
node_modules/
package-lock.json
*.min.md
'@ | Set-Content -NoNewline -Encoding utf8 .markdownlintignore

# .prettierrc.json
@'
{
  "proseWrap": "preserve",
  "tabWidth": 2,
  "useTabs": false,
  "semi": true,
  "singleQuote": true,
  "endOfLine": "lf",
  "printWidth": 120,
  "overrides": [
    {
      "files": ["*.md", "*.markdown"],
      "options": {
        "proseWrap": "preserve"
      }
    },
    {
      "files": ["*.yml", "*.yaml"],
      "options": {
        "tabWidth": 2
      }
    }
  ]
}
'@ | Set-Content -NoNewline -Encoding utf8 .prettierrc.json

# .prettierignore
@'
.git/
node_modules/
package-lock.json
*.min.js
*.min.css
*.snap
'@ | Set-Content -NoNewline -Encoding utf8 .prettierignore

# .pre-commit-config.yaml
@'
repos:
  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.6.0
    hooks:
      - id: trailing-whitespace
      - id: end-of-file-fixer
      - id: mixed-line-ending
        args:
          - --fix=lf
        exclude: '\.(bat|cmd|ps1)$'
      - id: check-yaml
        args:
          - --allow-multiple-documents
      - id: check-json
      - id: check-added-large-files
        args:
          - --maxkb=10240
        exclude: '\.(xlsx|xlsm|xls|png|jpe?g|gif|pdf|zip|7z|mp4|mp3|woff2?|ttf|otf|bin|dll|exe|so|dylib)$'
      - id: check-case-conflict
      - id: check-merge-conflict
      - id: detect-private-key
      - id: check-executables-have-shebangs
      - id: check-shebang-scripts-are-executable
      - id: check-builtin-literals
      - id: debug-statements

  - repo: https://github.com/igorshubovych/markdownlint-cli
    rev: v0.40.0
    hooks:
      - id: markdownlint
        args:
          - --fix

  - repo: https://github.com/pre-commit/mirrors-prettier
    rev: v3.1.0
    hooks:
      - id: prettier
        types_or:
          - markdown
          - yaml
          - json
        exclude: 'package-lock\.json'
'@ | Set-Content -NoNewline -Encoding utf8 .pre-commit-config.yaml

# .github/workflows/text-markdown-qa.yml
@'
name: Text & Markdown Perfection

on:
  push:
    paths:
      - "**/*.md"
      - "**/*.markdown"
      - "**/*.txt"
      - "**/*.yml"
      - "**/*.yaml"
      - "**/*.json"
      - ".editorconfig"
      - ".gitattributes"
      - ".markdownlint.json"
      - ".markdownlintignore"
      - ".prettierrc.json"
      - ".prettierignore"
      - ".pre-commit-config.yaml"
      - ".github/workflows/text-markdown-qa.yml"
      - ".github/workflows/text-markdown-auto-fix.yml"
  pull_request:
    paths:
      - "**/*.md"
      - "**/*.markdown"
      - "**/*.txt"
      - "**/*.yml"
      - "**/*.yaml"
      - "**/*.json"
      - ".editorconfig"
      - ".gitattributes"
      - ".markdownlint.json"
      - ".markdownlintignore"
      - ".prettierrc.json"
      - ".prettierignore"
      - ".pre-commit-config.yaml"
      - ".github/workflows/text-markdown-qa.yml"
      - ".github/workflows/text-markdown-auto-fix.yml"
  workflow_dispatch:

permissions:
  contents: read

concurrency:
  group: "text-markdown-${{ github.ref }}"
  cancel-in-progress: true

jobs:
  text-markdown:
    name: pre-commit
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Cache pre-commit
        uses: actions/cache@v4
        with:
          path: ~/.cache/pre-commit
          key: "pre-commit-${{ hashFiles('.pre-commit-config.yaml') }}"
          restore-keys: |
            pre-commit-

      - name: Install pre-commit
        run: |
          python -m pip install --upgrade pip
          python -m pip install pre-commit

      - name: Run pre-commit
        run: pre-commit run --all-files --show-diff-on-failure
'@ | Set-Content -NoNewline -Encoding utf8 .github/workflows/text-markdown-qa.yml

# .github/workflows/text-markdown-auto-fix.yml
@'
name: Auto Fix Text & Markdown

on:
  workflow_dispatch:

permissions:
  contents: write
  pull-requests: write

concurrency:
  group: "text-markdown-auto-fix-${{ github.ref }}"
  cancel-in-progress: false

jobs:
  auto-fix:
    name: Auto fix and open PR
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install pre-commit
        run: |
          python -m pip install --upgrade pip
          python -m pip install pre-commit

      - name: Run pre-commit fixes
        run: pre-commit run --all-files || true

      - name: Detect changes
        id: changes
        run: |
          if [ -n "$(git status --porcelain)" ]; then
            echo "changed=true" >> "$GITHUB_OUTPUT"
          else
            echo "changed=false" >> "$GITHUB_OUTPUT"
          fi

      - name: Create Pull Request
        if: steps.changes.outputs.changed == 'true'
        uses: peter-evans/create-pull-request@v6
        with:
          token: ${{ secrets.GITHUB_TOKEN }}
          commit-message: "style: auto-fix text and markdown"
          branch: bot/text-markdown-perfection
          delete-branch: true
          title: "style: auto-fix text and markdown"
          body: |
            This PR applies automated text and markdown fixes.
'@ | Set-Content -NoNewline -Encoding utf8 .github/workflows/text-markdown-auto-fix.yml

# docs/qa/TEXT_MARKDOWN_DEFINITION_OF_DONE.md
@'
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
'@ | Set-Content -NoNewline -Encoding utf8 docs/qa/TEXT_MARKDOWN_DEFINITION_OF_DONE.md

# .github/pull_request_template.md
@'
## Checklist

- [ ] Text and Markdown QA job is green.
- [ ] No trailing spaces, no BOM, and final newline exists.
- [ ] Persian content is readable and technical terms are consistent.
- [ ] Screenshots or sample output are added if UI/output changes.

## Notes

Describe what changed and why.
'@ | Set-Content -NoNewline -Encoding utf8 .github/pull_request_template.md

# .github/dependabot.yml
@'
version: 2
updates:
  - package-ecosystem: "github-actions"
    directory: "/"
    schedule:
      interval: "weekly"
    open-pull-requests-limit: 5
'@ | Set-Content -NoNewline -Encoding utf8 .github/dependabot.yml

# .vscode/extensions.json
@'
{
  "recommendations": [
    "DavidAnson.vscode-markdownlint",
    "esbenp.prettier-vscode",
    "EditorConfig.EditorConfig"
  ]
}
'@ | Set-Content -NoNewline -Encoding utf8 .vscode/extensions.json

# .vscode/settings.json
@'
{
  "editor.formatOnSave": true,
  "editor.defaultFormatter": "esbenp.prettier-vscode",
  "files.eol": "\n",
  "files.insertFinalNewline": true,
  "files.trimTrailingWhitespace": true,
  "[markdown]": {
    "editor.defaultFormatter": "esbenp.prettier-vscode",
    "editor.rulers": [120]
  },
  "[yaml]": {
    "editor.defaultFormatter": "esbenp.prettier-vscode"
  },
  "[json]": {
    "editor.defaultFormatter": "esbenp.prettier-vscode"
  }
}
'@ | Set-Content -NoNewline -Encoding utf8 .vscode/settings.json

# ========== نرمال‌سازی فایل‌های متنی ==========
Write-Host "🔧 نرمال‌سازی فایل‌های متنی..." -ForegroundColor Cyan

# الگوهای فایل‌های متنی که باید پردازش بشن
$patterns = @('*.md', '*.markdown', '*.txt', '*.yml', '*.yaml', '*.json',
             '.editorconfig', '.gitattributes', '.markdownlintignore', '.prettierignore')

# گرفتن لیست فایل‌ها از گیت (با جداکننده null برای امنیت)
$nullSeparatedList = git ls-files -z -- $patterns 2>$null
if ($nullSeparatedList) {
    $fileList = $nullSeparatedList -split "`0" | Where-Object { $_ -ne '' }

    foreach ($f in $fileList) {
        # رد کردن package-lock.json
        if ($f -match '(^|/)package-lock\.json$') { continue }

        if (Test-Path -LiteralPath $f -PathType Leaf) {
            Write-Host "  📄 $f" -ForegroundColor DarkGray
            $rawContent = Get-Content -LiteralPath $f -Raw -Encoding UTF8

            # 1. تبدیل CRLF → LF (حذف carriage return)
            $cleanContent = $rawContent -replace "`r`n", "`n"
            # 2. حذف فضاهای خالی آخر هر خط (و tab)
            $cleanContent = $cleanContent -replace '(?m)[\t ]+$', ''
            # 3. اطمینان از وجود newline پایانی (اگر فایل خالی نباشد)
            if ($cleanContent.Length -gt 0 -and -not $cleanContent.EndsWith("`n")) {
                $cleanContent += "`n"
            }

            # نوشتن فایل با UTF-8 بدون BOM، با خط‌های LF
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllText($f, $cleanContent, $utf8NoBom)
        }
    }
}

# ========== آماده‌سازی commit ==========
Write-Host "📦 اضافه کردن فایل‌ها به stage..." -ForegroundColor Cyan

# فایل‌های ساخته‌شده
$filesToAdd = @(
    '.editorconfig',
    '.gitattributes',
    '.markdownlint.json',
    '.markdownlintignore',
    '.prettierrc.json',
    '.prettierignore',
    '.pre-commit-config.yaml',
    '.github/workflows/text-markdown-qa.yml',
    '.github/workflows/text-markdown-auto-fix.yml',
    '.github/pull_request_template.md',
    '.github/dependabot.yml',
    '.vscode/extensions.json',
    '.vscode/settings.json',
    'docs/qa/TEXT_MARKDOWN_DEFINITION_OF_DONE.md'
)

git add -- $filesToAdd 2>$null
# اضافه کردن فایل‌های تغییرکرده‌ی متنی
git add -u -- $patterns 2>$null

# بررسی تغییرات stage
$stagedDiff = git diff --cached --quiet 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "ℹ️ فایلی برای کامیت تغییر نکرد." -ForegroundColor Yellow
} else {
    # تنظیم نام و ایمیل در صورت عدم وجود
    if (-not (git config user.name 2>$null)) {
        git config user.name "Text QA Bot"
    }
    if (-not (git config user.email 2>$null)) {
        git config user.email "text-qa-bot@localhost"
    }

    git commit --no-verify -m "chore(qa): add text/markdown perfection pack"
    Write-Host "✅ کامیت انجام شد." -ForegroundColor Green
}

# ========== push ==========
$remote = git remote get-url origin 2>$null
if ($remote) {
    Write-Host "🚀 Push به origin..." -ForegroundColor Cyan
    $env:GIT_TERMINAL_PROMPT = 0
    git push -u origin HEAD
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ به ریموت push شد." -ForegroundColor Green
    } else {
        Write-Host "⚠️ push نشد. اگر نیاز به لاگین دارد، فقط این را بزن: git push -u origin HEAD" -ForegroundColor Magenta
    }
} else {
    Write-Host "ℹ️ ریموت origin ندارد." -ForegroundColor Yellow
}

Write-Host "✅ تمام. حالا Text & Markdown Perfection فعال است." -ForegroundColor Green