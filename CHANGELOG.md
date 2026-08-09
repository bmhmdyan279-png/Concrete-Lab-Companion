# Changelog

فرمت بر اساس [Keep a Changelog](https://keepachangelog.com/fa/1.1.0/) و نسخه‌گذاری [SemVer](https://semver.org/lang/fa/).

## [Unreleased]

### Added
- ساختار `validation/golden_cases/` برای تست‌های مرجع
- `config.yaml` برای پیکربندی مرکزی
- GitHub Actions CI/CD pipeline
- Issue Templates و PR Template
- `CITATION.cff` برای ارجاع علمی
- `CONTRIBUTING.md`

### Changed
- ادغام `build_workbook.py` و `build_project.py` → `build.py` واحد
- اصلاح ادعاهای README (حذف «۱۰۰٪»، «زنده»، «هوشمند»)
- انتقال فایل اکسل به GitHub Releases (حذف از پوشه‌های excel/ و releases/)
- اصلاح هش SHA-256 در صفحه فرود و README

### Fixed
- هش SHA-256 خراب در index.html و README
- حذف پسوند `update_hash.py]` از متن هش

## [1.0.0] - 2026-08-09

### Added
- ساختار اولیه مخزن
- شیت‌های راهنما، اطلاعات آزمون، ۱-۲، ۲-۴
- شیت Validation (۴ مورد)
- صفحه فرود GitHub Pages
- README اولیه