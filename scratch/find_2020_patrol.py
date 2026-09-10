import os

desktop = r"C:\Users\moonh\Desktop"
for root, dirs, files in os.walk(desktop):
    for f in files:
        if any(k in f for k in ['일일점검', '순찰', '2020', '야통대']):
            if f.endswith(('.xlsx', '.xls', '.csv', '.hwp', '.hwpx')):
                print(os.path.join(root, f))
