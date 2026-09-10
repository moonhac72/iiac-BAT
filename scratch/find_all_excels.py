import os

base = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더"
for root, dirs, files in os.walk(base):
    for f in files:
        if f.endswith(('.xlsx', '.xls', '.csv')):
            fp = os.path.join(root, f)
            print(f"[{f}] in {root} ({os.path.getsize(fp):,} bytes)")
