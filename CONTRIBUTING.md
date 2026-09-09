# راهنمای مشارکت

از علاقه‌مندی شما به بهبود **Concrete Lab Companion** سپاسگزاریم! 🙏

این مخزن یک موتور محاسباتی/کنترل‌کیفیت برای آزمایشگاه بتن است. قاعدهٔ حاکم بر همه‌چیز
این است: **علم در پایتون محاسبه می‌شود، اکسل فقط نمایش می‌دهد.** هر تغییری که این
مرز را جابه‌جا کند، با تست شکست می‌خورد.

---

## شروع سریع

```bash
git clone https://github.com/bmhmdyan279-png/Concrete-Lab-Companion.git
cd Concrete-Lab-Companion

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\Activate.ps1

pip install -r requirements-dev.txt
python -m pytest                 # همهٔ تست‌ها
python build.py --validate       # سوئیت طلایی + pytest (خروجی غیرصفر در صورت شکست)
python build.py --output output  # ساخت محصول + QA ساختاری + manifest
```

یا نصب قابل‌ویرایش بسته:

```bash
pip install -e ".[dev]"
concrete-lab-companion --help
```

---

## گردش کار

1. یک **Issue** باز کنید (یا یکی را پیدا کنید) و در PR به آن ارجاع دهید.
2. از `main` یک برنچ بسازید.
3. تغییر را همراه **تست** و — اگر منطق آزمایش را تغییر می‌دهد — همراه **مورد طلایی** انجام دهید.
4. قبل از PR، همهٔ دروازه‌های محلی را اجرا کنید (بخش بعد).
5. Pull Request بفرستید؛ CI همان دروازه‌ها را دوباره اجرا می‌کند.

### نام‌گذاری برنچ

```text
feature/add-test-3-2
fix/golden-coverage-4-1
docs/readme-usage
refactor/qa-tier-split
```

### پیام کامیت

از [Conventional Commits](https://www.conventionalcommits.org/) پیروی کنید؛ زبان فارسی یا
انگلیسی هر دو پذیرفته است، ولی پیام باید **واضح و قابل‌ردیابی** باشد:

```text
feat(engine): add fineness modulus to test 1-1
fix(qa): report a corrupt golden case as a failure, not a crash
docs(readme): document the --input data path
```

---

## دروازه‌های قبل از PR

همهٔ اینها باید سبز باشند. CI دقیقاً همین‌ها را اجرا می‌کند:

```bash
ruff check .                     # lint (تنظیمات در pyproject.toml)
python -m pytest --cov           # تست + پوشش (آستانه در pyproject.toml)
python build.py --validate       # سوئیت طلایی + pytest
python build.py --output output --no-protect
python scripts/audit_excel.py --output output
pre-commit run --all-files       # قالب‌بندی markdown/yaml/json
```

---

## قوانین سخت (این‌ها با تست اجبار می‌شوند)

| قانون                                                                                                 | کجا اجبار می‌شود                                 |
| ----------------------------------------------------------------------------------------------------- | ------------------------------------------------ |
| رندر اکسل **هیچ** فرمولی نمی‌نویسد                                                                    | `qa/structural.py` + `tests/test_consistency.py` |
| هر آزمایشِ `implemented=True` باید محاسبه‌گر **و** مورد طلایی داشته باشد                              | `tests/test_golden_corpus.py`                    |
| عدد `implemented` در `config.yaml` باید با رجیستری زنده برابر باشد                                    | `tests/test_consistency.py`                      |
| نسخه در ۵ جا (`__init__.py`, `pyproject.toml`, `config.yaml`, `CITATION.cff`, `CHANGELOG.md`) یکی است | `tests/test_consistency.py`                      |
| صفحهٔ فرود هیچ هش یا نسخه‌ای را hard-code نمی‌کند                                                     | `tests/test_consistency.py`                      |
| هر ماژول docstring دارد                                                                               | `tests/test_consistency.py`                      |
| نتیجهٔ تخمینی (چکش اشمیت) همیشه بج و سلب مسئولیت دارد                                                 | `qa/golden.py` + `tests/test_golden_dispatch.py` |

---

## افزودن یک آزمایش جدید

۱. **مشخصات** را در فصل مربوطهٔ `src/concrete_lab/specs/` ثبت کنید:

```python
register_test(TestSpec(
    id="3-2", title="آب‌اندازی", standard_code="C232",
    sheet_name="15_آزمایش_3-2", tab_color="1565C0",
    notes="Bleeding capacity per ASTM C232.",
))
```

۲. اگر استانداردش رولسیت دارد، در `src/concrete_lab/standards/` پیاده‌سازی و در
`registry.py` ثبت کنید (با `RULESET_VERSION` و `SCOPE`).

۳. **محاسبه‌گر** را در `src/concrete_lab/domain/engine.py` بنویسید و به `CALCULATORS`
اضافه کنید؛ ورودی‌ها را اعتبارسنجی کنید و قوانین فیزیکی را صریح رد کنید.

۴. `implemented=True` را روی spec بگذارید.

۵. **مورد طلایی** در `validation/golden_cases/3-2.json` بسازید و یک checker در
`ENGINE_CASE_CHECKERS` (`src/concrete_lab/qa/golden.py`) به آن وصل کنید.
بدون این، تست‌ها شکست می‌خورند — این عمدی است.

۶. `config.yaml` را به‌روز کنید: `implemented` و `pending` و `key_sheets`.

۷. `README.md` (جدول فهرست آزمایش‌ها) و `CHANGELOG.md` را به‌روز کنید.

> ⚠️ **هرگز نتیجهٔ آزمایشی را که پیاده‌سازی نشده جعل نکنید.** کاتالوگ آن را
> «در انتظار» نشان می‌دهد و این صداقت، ارزش اصلی محصول است.

---

## تولید مجدد `landing/release.json`

هش SHA-256 صفحهٔ فرود دستی وارد نمی‌شود؛ از manifest همان ساخت تولید می‌شود:

```bash
python build.py --output output --no-protect
python scripts/make_release_manifest.py --output output --dest landing/release.json
python scripts/make_release_manifest.py --output output --dest landing/release.json --check
```

CI در هر ریلیز این کار را انجام می‌دهد، پس معمولاً نیازی به اجرای دستی نیست.

---

## گزارش باگ

از [Issue Template باگ](.github/ISSUE_TEMPLATE/bug_report.md) استفاده کنید و این‌ها را
ضمیمه کنید:

- خروجی `python build.py --version`
- فایل `manifest JSON` همان ساخت (همراه ریلیز منتشر می‌شود)
- مسیر کامل reproduce

## درخواست قابلیت

از [Issue Template قابلیت](.github/ISSUE_TEMPLATE/feature_request.md) استفاده کنید.
اگر درخواست شما محاسبهٔ یک آزمایش جدید است، **منبع استاندارد و بند آن** را ذکر کنید —
بدون ارجاع به استاندارد، منطق علمی پذیرفته نمی‌شود.

---

## سبک کد

- `ruff` تنها داور سبک است؛ تنظیمات در `pyproject.toml` (طول خط ۱۲۰).
- type hints کامل، dataclassهای `frozen`، ثابت‌های نام‌دار در `constants.py`.
- docstring با قالب Google (Args / Returns / Raises).
- هیچ عدد جادویی در منطق؛ هر ثابت با منبع استانداردش مستند شود.
- markdown/yaml/json با `prettier` و `markdownlint` قالب‌بندی می‌شوند.

---

## مجوز

با ارسال PR، موافقت می‌کنید مشارکت شما تحت [CC BY-NC 4.0](LICENSE) منتشر شود.
توجه کنید که این مجوز **غیرتجاری** است.
