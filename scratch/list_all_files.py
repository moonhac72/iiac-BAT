import os

folder = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\조류 통계"
for f in os.listdir(folder):
    print(repr(f))

# Also check subdirectories
print("\nSubdirectories:")
for sub in ["17~20년 조류 통계", "21~23년 조류 통계", "완료된 자료"]:
    sub_p = os.path.join(folder, sub)
    if os.path.exists(sub_p):
        print(f"Inside {sub}:")
        for f in os.listdir(sub_p):
            print("  ", repr(f))
