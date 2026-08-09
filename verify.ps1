# ═══════════════════════════════════════════
#  verify.ps1 — بررسی سلامت مخزن
#  اجرا: powershell -ExecutionPolicy Bypass -File verify.ps1
# ═══════════════════════════════════════════

$ErrorActionPreference = "Stop"
$pass = 0; $fail = 0; $warn = 0

function Write-Check($name, $ok, $msg) {
    if ($ok) { Write-Host "  ✅ $name" -ForegroundColor Green; $script:pass++ }
    else     { Write-Host "  ❌ $name — $msg" -ForegroundColor Red; $script:fail++ }
}
function Write-Warn($name, $msg) {
    Write-Host "  ⚠️  $name — $msg" -ForegroundColor Yellow; $script:warn++
}

Write-Host ""
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Concrete Lab Companion — Verify"     -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ─── ۱. فایل‌های ضروری ───
Write-Host "📁 فایل‌های ضروری:" -ForegroundColor White
$requiredFiles = @(
    "build.py", "config.yaml", "requirements.txt",
    "README.md", "LICENSE", "CITATION.cff",
    "CHANGELOG.md", "CONTRIBUTING.md", ".gitignore",
    "validation/errata.yaml",
    "landing/index.html",
    ".github/workflows/build.yml",
    ".github/workflows/pages.yml",
    ".github/ISSUE_TEMPLATE/bug_report.md",
    ".github/ISSUE_TEMPLATE/feature_request.md",
    ".github/PULL_REQUEST_TEMPLATE.md"
)
foreach ($f in $requiredFiles) {
    Write-Check $f (Test-Path $f) "فایل یافت نشد"
}

# ─── ۲. فایل‌های ممنوعه ───
Write-Host ""
Write-Host "🚫 فایل‌هایی که باید حذف شده باشند:" -ForegroundColor White
$forbidden = @(
    "build_workbook.py", "build_project.py", "update_hash.py",
    "excel/", "releases/"
)
foreach ($f in $forbidden) {
    $exists = Test-Path $f
    if ($exists) { Write-Check $f $false "هنوز وجود دارد! حذف کن" }
    else         { Write-Check "$f (حذف شده)" $true "" }
}

# ─── ۳. Golden Cases ───
Write-Host ""
Write-Host "🧪 Golden Test Cases:" -ForegroundColor White
$goldenDir = "validation/golden_cases"
if (Test-Path $goldenDir) {
    $cases = Get-ChildItem $goldenDir -Filter "*.json"
    Write-Check "پوشه golden_cases" $true ""
    Write-Host "     تعداد: $($cases.Count) فایل" -ForegroundColor Gray
    if ($cases.Count -lt 4) {
        Write-Warn "تعداد Golden Cases" "فقط $($cases.Count) — هدف: حداقل ۴"
    }
} else {
    Write-Check "پوشه golden_cases" $false "یافت نشد"
}

# ─── ۴. اجرای build.py ───
Write-Host ""
Write-Host "🔨 اجرای build.py:" -ForegroundColor White
try {
    $buildOutput = & python build.py 2>&1
    $buildExit = $LASTEXITCODE
    if ($buildExit -eq 0) {
        Write-Check "build.py اجرا شد" $true ""
    } else {
        Write-Check "build.py" $false "Exit code: $buildExit"
        Write-Host $buildOutput -ForegroundColor Red
    }
} catch {
    Write-Check "build.py" $false $_.Exception.Message
}

# ─── ۵. بررسی خروجی ───
Write-Host ""
Write-Host "📦 خروجی ساخت:" -ForegroundColor White
$xlsx = Get-ChildItem "output" -Filter "*.xlsx" -ErrorAction SilentlyContinue
if ($xlsx) {
    Write-Check "فایل xlsx" $true ""
    Write-Host "     $($xlsx.Name) ($([math]::Round($xlsx.Length/1KB, 1)) KB)" -ForegroundColor Gray

    # محاسبه هش
    $hash = (Get-FileHash $xlsx.FullName -Algorithm SHA256).Hash.ToLower()
    Write-Host "     SHA-256: $hash" -ForegroundColor Gray

    # بررسی هش در index.html
    $indexContent = Get-Content "landing/index.html" -Raw -Encoding UTF8
    if ($indexContent -match $hash) {
        Write-Check "تطابق هش با index.html" $true ""
    } else {
        Write-Warn "هش در index.html" "هش build ($hash) با index.html همخوان نیست"
        Write-Host "     دستور جایگزینی:" -ForegroundColor Gray
        Write-Host "     (Get-Content landing/index.html -Raw) -replace '[a-f0-9]{64}', '$hash' | Set-Content landing/index.html -Encoding UTF8" -ForegroundColor Cyan
    }
} else {
    Write-Check "فایل xlsx" $false "در output/ یافت نشد"
}

# ─── ۶. بررسی config.yaml ───
Write-Host ""
Write-Host "⚙️ config.yaml:" -ForegroundColor White
$configContent = Get-Content "config.yaml" -Raw -Encoding UTF8
if ($configContent -match "calc_fill") {
    Write-Check "نام فیلد calc_fill" $true ""
} elseif ($configContent -match "calculation_fill") {
    Write-Warn "نام فیلد" "calculation_fill → باید calc_fill باشد (مطابق build.py)"
} else {
    Write-Warn "نام فیلد رنگ" "هیچ کلید calc_fill یا calculation_fill یافت نشد"
}

# ─── ۷. بررسی .gitignore ───
Write-Host ""
Write-Host "🛡️ .gitignore:" -ForegroundColor White
$giContent = Get-Content ".gitignore" -Raw -Encoding UTF8
$giChecks = @("*.xlsx", "__pycache__", ".idea", ".vscode", ".env")
foreach ($item in $giChecks) {
    Write-Check "  $item" ($giContent -match [regex]::Escape($item)) "در .gitignore نیست"
}

# ─── جمع‌بندی ───
Write-Host ""
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  نتیجه: $pass ✅ | $fail ❌ | $warn ⚠️" -ForegroundColor White
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan

if ($fail -gt 0) {
    Write-Host "  ❌ مخزن آماده انتشار نیست." -ForegroundColor Red
    exit 1
} else {
    Write-Host "  ✅ مخزن آماده انتشار است." -ForegroundColor Green
    if ($warn -gt 0) {
        Write-Host "  ⚠️  $warn هشدار را بررسی کن." -ForegroundColor Yellow
    }
    exit 0
}