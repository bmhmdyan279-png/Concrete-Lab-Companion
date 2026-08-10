# 🧪 Concrete Lab Companion

> پیوست دیجیتال کتاب «فناوری بتن در آزمایشگاه» — ابزار محاسباتی ۲۰ آزمایش استاندارد ASTM/ISIRI

[![GitHub release](https://img.shields.io/github/v/release/bmhmdyan279-png/Concrete-Lab-Companion)](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 📥 دانلود سریع

👉 **[دانلود آخرین نسخه از GitHub Releases](https://github.com/bmhmdyan279-png/Concrete-Lab-Companion/releases/latest)**

یا از صفحه رسمی: <https://bmhmdyan279-png.github.io/Concrete-Lab-Companion/>

## ✨ ویژگی‌ها

- 🧪 **۲۰ آزمایش بتن** (دانه‌بندی، مقاومت، اسلامپ، اشمیت و...)
- 🛡️ **بدون ماکرو** — سازگار با Excel Desktop / Web / Mobile
- 📊 **نمودار دانه‌بندی لگاریتمی** + پاکت استاندارد ISIRI 302
- ⚠️ **شناسایی ۸ خطای چاپی کتاب** با توضیح اصلاحی
- 🔒 **محافظت شیت‌ها** (رمز: `ConcreteLab2026!` — فقط برای جلوگیری از ویرایش تصادفی)
- 🔐 **SHA-256** برای تأیید اصالت فایل

## 📋 وضعیت فعلی (v1.1.0)

| معیار                     | وضعیت                    |
| ------------------------- | ------------------------ |
| آزمایش‌های پیاده‌سازی‌شده | **۲۰ از ۲۰** ✅          |
| فرمول‌های فعال            | **۱۰۰٪** ✅              |
| خطاهای کتاب شناسایی‌شده   | **۸ از ۸** ✅            |
| داشبورد                   | **پویا و فرمول‌محور** ✅ |
| Golden Tests              | **۵ تست خودکار** ✅      |

## 🔐 تأیید اصالت فایل

پس از دانلود، هش SHA-256 را بررسی کنید:

**PowerShell:**

```powershell
Get-FileHash .\Concrete_Lab_Companion_v1.1.0.xlsx -Algorithm SHA256
```

**مقدار مورد انتظار:** `b2668023a3dade735eebfc07dd0bbb554b6ddd43da280c209888b71b66c6804e`

## 🛠️ ساخت از سورس (برای توسعه‌دهندگان)

```bash
git clone https://github.com/bmhmdyan279-png/Concrete-Lab-Companion.git
cd Concrete-Lab-Companion
pip install -r requirements.txt
python build.py
```

فایل خروجی در `output/Concrete_Lab_Companion_v1.1.0.xlsx` ساخته می‌شود.

## 📝 Changelog

### v1.1.0 (2026-08-10)

- 🔧 رفع ۴ باگ بحرانی (داشبورد، ۲-۴، ۴-۱، نمودار)
- ✅ استفاده از MROUND برای گرد کردن مهندسی
- ✅ ستون میانی NA() برای نمودار
- ✅ کنترل جرم شامل پان

### v1.0.0 (2026-08-09)

- 🎉 انتشار اولیه
- 🧪 پیاده‌سازی ۲۰ آزمایش ASTM/ISIRI
- 📊 نمودار دانه‌بندی لگاریتمی
- ⚠️ شناسایی ۸ خطای کتاب

## 📖 استفاده

1. فایل را دانلود کنید
2. در Excel (Desktop/Online/Mobile) باز کنید
3. در سلول‌های **زرد** مقادیر آزمایش را وارد کنید
4. نتایج در سلول‌های **خاکستری** محاسبه می‌شوند
5. رمز شیت‌ها: `ConcreteLab2026!`

## ⚠️ سلب مسئولیت

این ابزار برای **راستی‌آزمایی** محاسبات دستی است، نه جایگزین آن. همیشه طبق استاندارد مرجع محاسبه کنید و نتایج را با این ابزار مقایسه نمایید.

## 📚 منبع

کتاب «فناوری بتن در آزمایشگاه» — چاپ ۱۴۰۵

## 📄 لایسنس

MIT License
