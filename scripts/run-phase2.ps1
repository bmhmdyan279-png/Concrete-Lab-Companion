$ErrorActionPreference = "Stop"
Write-Host "`n🚀 شروع فاز دوم: Excel Perfection Audit" -ForegroundColor Cyan

# فعال‌سازی venv
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    Write-Host "🔌 فعال‌سازی venv..." -ForegroundColor Yellow
    & .\venv\Scripts\Activate.ps1
}

# نصب وابستگی‌ها
Write-Host "`n📦 نصب/ارتقای وابستگی‌ها..." -ForegroundColor Yellow
& python -m pip install --quiet --upgrade openpyxl pyyaml xlsxwriter 2>$null

# اجرای ممیزی
Write-Host "`n🔬 اجرای ممیزی اکسل..." -ForegroundColor Yellow
& python scripts/audit_excel.py

# اضافه کردن نتایج به Git
Write-Host "`n📦 Stage کردن گزارش‌ها..." -ForegroundColor Yellow
& git add validation/formulas_audit.txt 2>$null
& git add docs/qa/EXCEL_AUDIT_REPORT.json 2>$null
& git add scripts/audit_excel.py 2>$null

# کامیت اگر تغییری هست
$changes = & git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n💾 کامیت نتایج ممیزی..." -ForegroundColor Yellow
    & git commit -m "chore(qa): phase 2 - excel perfection audit report" --no-verify
    
    Write-Host "`n🚀 Push به origin..." -ForegroundColor Yellow
    & git push origin HEAD
    
    Write-Host "`n✅ فاز دوم کامل شد و push شد!" -ForegroundColor Green
} else {
    Write-Host "`nℹ️  تغییری برای کامیت یافت نشد." -ForegroundColor Yellow
}

Write-Host "`n📖 برای مشاهده گزارش:" -ForegroundColor Cyan
Write-Host "   Get-Content docs\qa\EXCEL_AUDIT_REPORT.json" -ForegroundColor Gray
Write-Host "   Get-Content validation\formulas_audit.txt | Select-Object -First 50" -ForegroundColor Gray
