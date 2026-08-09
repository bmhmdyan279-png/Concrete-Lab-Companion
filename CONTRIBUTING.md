<div dir="rtl">

# راهنمای مشارکت

از علاقه‌مندی شما به بهبود Concrete Lab Companion سپاسگزاریم! 🙏

## شروع سریع

1. **Fork** → **Clone** → **Branch**
2. تغییرات را اعمال کنید
3. `python build.py` را اجرا کنید و مطمئن شوید خطا ندارد
4. Pull Request بسازید

## قوانین

- هر PR باید یک Issue مرتبط داشته باشد
- از Semantic Versioning پیروی کنید
- فرمول‌ها را hard-code نکنید؛ از `config.yaml` یا `validation/errata.yaml` بخوانید
- هر آزمایش جدید باید Golden Test Case داشته باشد
- commit message‌ها به فارسی یا انگلیسی، ولی واضح و descriptive

## ساختار Branch
feature/add-test-3-2
fix/hash-display
docs/readme-update

text

## تست قبل از PR

```bash
python build.py
python -m pytest validation/ -v
گزارش باگ
از Issue Template باگ استفاده کنید.

درخواست قابلیت
از Issue Template قابلیت استفاده کنید.

</div> ```