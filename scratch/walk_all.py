import os

base = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더"
for root, dirs, files in os.walk(base):
    print(f"\nDIR: {root}")
    for f in files:
        fp = os.path.join(root, f)
        print(f"  FILE: {f} ({os.path.getsize(fp):,} bytes)")
