import os, hashlib, re

def get_hash(path):
    h = hashlib.sha256()
    with open(path,"rb") as f:
        for b in iter(lambda: f.read(4096),b""): h.update(b)
    return h.hexdigest()

def update(file, hash_val):
    with open(file, 'r', encoding='utf-8') as f: c = f.read()
    # استفاده از \g<1> برای جلوگیری از تداخل با ارقام هش (The Golden Standard)
    if file.endswith('.md'): 
        c = re.sub(r'(\*\*SHA-256:\*\*\s*`?)([a-fA-F0-9]{64}|[^\s`\]]+)(`?)', r'\g<1>' + hash_val + r'\g<3>', c)
    elif file.endswith('.html'): 
        c = re.sub(r'(SHA-256:\s*)([a-fA-F0-9]{64}|[A-Z_\-\[\]a-z\s]+)', r'\g<1>' + hash_val, c)
    
    with open(file, 'w', encoding='utf-8') as f: f.write(c)
    print(f"✅ {file} updated successfully.")

files = [f for f in os.listdir('releases') if f.endswith('.xlsx')]
if not files: print("❌ فایل اکسل در پوشه releases یافت نشد.")
else:
    h = get_hash(os.path.join('releases', files[0]))
    print(f"🔐 Hash: {h}")
    update('README.md', h)
    update('index.html', h)
