<div dir="rtl" align="right">

# 🧪 Concrete Lab Companion

**پیوست دیجیتال کتاب «فناوری بتن» — ماشین‌حساب، اعتبارسنجی و QA آزمایشگاهی**

[![Release](https://img.shields.io/github/v/release/bmhmdyan279-png/Concrete-Lab-Companion?label=%D9%86%D8%B3%D8%AE%D9%87&color=blue)](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/releases/latest)
[![Build](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/actions/workflows/build.yml/badge.svg)](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/actions/workflows/build.yml)
[![License: CC BY-NC 4.0](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc/4.0/)
[](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/releases/latest)

---

## ⬇️ دانلود سریع

| فایل                   | لینک                                                                                                          |
| ---------------------- | ------------------------------------------------------------------------------------------------------------- |
| 📥 آخرین نسخه Workbook | [GitHub Releases → دانلود xlsx](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/releases/latest)    |
| 🔑 رمز شیت‌ها          | `ConcreteLab2026!` _(عمدی و عمومی — فقط برای جلوگیری از ویرایش تصادفی)_                                       |
| 📖 صفحه فرود           | [bmhmdyan279-png.github.io/Concrete-Lab-Companion](https://bmhmdyan279-png.github.io/Concrete-Lab-Companion/) |

---

## این ابزار چه کاری انجام می‌دهد؟

- **۲۰ آزمایش بتن** (اسلامپ، مقاومت فشاری، دانه‌بندی و…) با فرمول‌های Excel
- **بدون ماکرو (VBA-Free)** → روی Desktop، Excel Online و موبایل اجرا می‌شود
- **اعتبارسنجی ورودی** با پیام‌های خطای فارسی
- **رنگ‌بندی استاندارد**: 🟡 ورودی | ⬜ محاسبه | 🟢 قبولی | 🔴 خطا
- **شیت‌های مرجع پنهان** برای ردیابی استانداردها

---

## ⚠️ وضعیت فعلی — شفاف و صادقانه

| معیار                     | وضعیت           | توضیح                                    |
| ------------------------- | --------------- | ---------------------------------------- |
| شیت‌های پیاده‌سازی‌شده    | `12/20`         | بقیه در دست توسعه                        |
| فرمول‌های فعال            | `~70%`          | نه ۱۰۰٪ — هنوز تکمیل نشده                |
| اعتبارسنجی (Validation)   | `4/8 خطا`       | فقط ۴ مورد در `_Validation_Data` ثبت شده |
| تست خودکار (Golden Tests) | `0`             | ساختار آماده، داده‌ها در حال تکمیل       |
| داشبورد                   | **ایستا**       | مقادیر hard-coded، نه زنده               |
| سازگاری iOS/Android       | **آزمایش‌نشده** | Designed for، نه Verified                |

> **توجه:** این پروژه هنوز به نسخه پایدار نرسیده. ادعاهای قبلی (۱۰۰٪ فرمول فعال، داشبورد زنده، سیستم هوشمند) اغراق‌آمیز بودند و اصلاح شدند.

---

## 🔐 امنیت — دقیق و شفاف

| موضوع              | واقعیت                                                                             |
| ------------------ | ---------------------------------------------------------------------------------- |
| SHA-256            | **File Integrity Verification** — تأیید فایل دانلودی دست‌نخورده است. «امنیت» نیست. |
| رمز شیت            | **Formula accidental-edit protection** — مکانیزم امنیتی واقعی نیست.                |
| رمز عمومی          | عمداً عمومی است. `Password ≠ Secret`                                               |
| زنجیره اعتماد کامل | نیازمند Signed Release + Provenance (در نقشه راه)                                  |

### تأیید هش فایل

```powershell
# PowerShell
Get-FileHash .\Concrete_Lab_Companion_v1.0.0.xlsx -Algorithm SHA256
```

مقدار مورد انتظار در [صفحه فرود](https://bmhmdyan279-png.github.io/Concrete-Lab-Companion/) و هر Release درج شده است.

---

## 📐 ساختار مخزن

```
Concrete-Lab-Companion/
├── build.py                  # تنها نقطه ساخت (Canonical)
├── config.yaml               # پیکربندی مرکزی
├── requirements.txt
├── README.md
├── CHANGELOG.md
├── LICENSE
├── CITATION.cff
├── CONTRIBUTING.md
├── .gitignore
├── .github/
│   ├── workflows/build.yml
│   ├── ISSUE_TEMPLATE/
│   │   ├── bug_report.md
│   │   └── feature_request.md
│   └── PULL_REQUEST_TEMPLATE.md
├── src/
│   ├── workbook/
│   ├── standards/
│   ├── validation/
│   └── reporting/
├── validation/
│   ├── errata.yaml           # Single Source of Truth خطاها
│   └── golden_cases/
│       ├── 1-1.json
│       ├── 1-2.json
│       └── ...
├── landing/
│   └── index.html
└── docs/
    └── screenshots/
```

---

## 🛠️ ساخت و اجرا

```bash
# ۱. Clone
git clone https://github.com/bmhmdyan279-png/Concrete-Lab-Companion.git
cd Concrete-Lab-Companion

# ۲. نصب وابستگی‌ها
pip install -r requirements.txt

# ۳. ساخت Workbook
python build.py

# ۴. اجرای تست‌ها
python -m pytest validation/ -v
```

---

## 📋 استانداردهای مرجع

هر آزمایش باید مشخصاً بگوید:

```yaml
test_id: '2-4'
standard: 'ASTM C39'
edition: '2023'
clause: 'Section 8.2'
formula_source: "f'c = P / A"
```

---

## 📊 تعریف «پیاده‌سازی کامل» یک آزمایش

یک آزمایش **implemented** محسوب می‌شود اگر:

- [x] شیت اختصاصی داشته باشد
- [x] سلول‌های ورودی مشخص و unlocked
- [x] فرمول محاسبه با قفل
- [x] اعتبارسنجی ورودی (DV)
- [x] واحد اندازه‌گیری
- [x] خروجی نهایی
- [x] ارجاع به استاندارد (Standard + Clause)
- [x] حداقل یک Golden Test Case
- [x] در `_Validation_Data` ثبت شده باشد

---

## 📜 مجوز

[CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/) — استفاده آزاد آموزشی با ارجاع به کتاب. استفاده تجاری ممنوع.

## 📖 ارجاع علمی

فایل [`CITATION.cff`](CITATION.cff) را ببینید.

---

## 🔗 لینک‌های مفید

- [CHANGELOG](CHANGELOG.md) — تاریخچه تغییرات
- [CONTRIBUTING](CONTRIBUTING.md) — راهنمای مشارکت
- [Issue Templates](.github/ISSUE_TEMPLATE/) — گزارش باگ / درخواست قابلیت
- [GitHub Actions](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/actions) — وضعیت build

</div>
