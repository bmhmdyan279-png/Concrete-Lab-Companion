import openpyxl
import hashlib
import os
import glob

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Concrete Lab Companion | نسخه بتا</title>
    <style>
        body { font-family: 'Vazirmatn', 'Arial', sans-serif; background: #f4f7f6; color: #333; margin: 0; padding: 20px; line-height: 1.6; }
        .container { max-width: 800px; margin: auto; background: #fff; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .beta-banner { background: #ff9800; color: #fff; padding: 10px; text-align: center; border-radius: 6px; margin-bottom: 20px; font-weight: bold; }
        .download-btn { display: block; width: 100%; padding: 15px; background: #28a745; color: #fff; text-align: center; text-decoration: none; border-radius: 8px; font-size: 1.2em; margin: 20px 0; box-shadow: 0 4px 6px rgba(40,167,69,0.3); }
        .hash-box { background: #2d2d2d; color: #00ff00; padding: 15px; border-radius: 6px; font-family: monospace; word-break: break-all; position: relative; direction: ltr; text-align: left; }
        .copy-btn { position: absolute; top: 10px; right: 10px; background: #555; color: #fff; border: none; padding: 5px 10px; cursor: pointer; border-radius: 4px; font-size: 12px; }
        .status-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin: 20px 0; }
        .status-card { background: #e9ecef; padding: 15px; border-radius: 8px; text-align: center; }
        .errata { background: #fff3cd; border-right: 5px solid #ffc107; padding: 15px; margin: 15px 0; border-radius: 4px; }
        h1 { color: #2c3e50; }
    </style>
</head>
<body>
    <div class="container">
        <div class="beta-banner">⚠️ توجه: این فایل یک نسخه بتا (Beta) است. ۱۲ از ۲۰ آزمایش پیاده‌سازی شده‌اند.</div>
        <h1>🧪 Concrete Lab Companion</h1>
        <p>ابزار همراه آزمایشگاه بتن | نسخه <span id="version">v1.0.0</span></p>

        <a id="downloadLink" href="#" class="download-btn" target="_blank">📥 دانلود مستقیم فایل اکسل (xlsx)</a>

        <div class="status-grid">
            <div class="status-card"><h3>12/20</h3><p>آزمایش‌های فعال</p></div>
            <div class="status-card"><h3>~70%</h3><p>فرمول‌های تکمیل‌شده</p></div>
        </div>

        <h3>🔐 اصالت‌سنجی فایل (SHA-256)</h3>
        <div class="hash-box">
            <button class="copy-btn" onclick="copyHash()">Copy</button>
            <span id="sha256">در حال محاسبه...</span>
        </div>

        <div class="errata">
            <strong>⚠️ اِراتای کتاب:</strong> در آزمایش ۱-۴ج، اگر OD=1.51 وارد کردید، اکسل هشدار می‌دهد. مقدار صحیح فیزیکی SSD همواره بزرگتر از OD است.
        </div>
    </div>

    <script>
        const repo = 'bmhmdyan279-png/Concrete-Lab-Companion';
        fetch(`https://api.github.com/repos/${repo}/releases/latest`)
            .then(r => r.json())
            .then(data => {
                document.getElementById('version').innerText = data.tag_name;
                const asset = data.assets.find(a => a.name.endsWith('.xlsx'));
                if(asset) document.getElementById('downloadLink').href = asset.browser_download_url;
                document.getElementById('sha256').innerText = '__SHA256_PLACEHOLDER__';
            }).catch(() => {
                document.getElementById('sha256').innerText = '__SHA256_PLACEHOLDER__';
            });

        function copyHash() {
            navigator.clipboard.writeText(document.getElementById('sha256').innerText);
            alert('هش کپی شد!');
        }
    </script>
</body>
</html>"""


def get_sha256(filepath):
  sha256_hash = hashlib.sha256()
  with open(filepath, "rb") as f:
    for byte_block in iter(lambda: f.read(4096), b""):
      sha256_hash.update(byte_block)
  return sha256_hash.hexdigest()


def audit_excel(excel_path, audit_out):
  wb = openpyxl.load_workbook(excel_path, data_only=False)
  volatile_funcs = ['INDIRECT', 'OFFSET', 'TODAY', 'NOW', 'RAND', 'RANDBETWEEN']
  audit_log = ["# Formula Audit Report", f"# File: {excel_path}\n"]

  for ws in wb.worksheets:
    audit_log.append(f"\n### Sheet: {ws.title} ###")
    for row in ws.iter_rows():
      for cell in row:
        if cell.value and str(cell.value).startswith('='):
          formula = str(cell.value)
          audit_log.append(f"[{cell.coordinate}] {formula}")
          for vf in volatile_funcs:
            if vf in formula.upper():
              audit_log.append(f"  ⚠️ VOLATILE: {vf} found in {cell.coordinate}")

  with open(audit_out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(audit_log))
  print(f"✅ Audit saved to {audit_out}")


if __name__ == "__main__":
  excel_files = glob.glob("**/*.xlsx", recursive=True)
  # پیدا کردن آخرین نسخه بر اساس نام فایل (مثلاً v1.1.0 > v1.0.0)
  xlsx_files = sorted(
    [f for f in excel_files if 'audit' not in f and 'template' not in f],
    reverse=True
  )
  target_excel = xlsx_files[0] if xlsx_files else None

  if not target_excel:
    print("❌ No .xlsx file found!")
  else:
    print(f"🔍 Auditing {target_excel}...")
    audit_excel(target_excel, "formulas_audit.txt")
    sha = get_sha256(target_excel)
    html_content = HTML_TEMPLATE.replace('__SHA256_PLACEHOLDER__', sha)
    with open("index.html", 'w', encoding='utf-8') as f:
      f.write(html_content)
    print(f"✅ Landing Page updated with SHA-256: {sha}")
