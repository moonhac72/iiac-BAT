import pandas as pd

df_s = pd.read_csv(r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2017_2020_birdstrike_resampled.csv")
df_c = pd.read_csv(r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2017_2020_carcass_resampled.csv")
df_l = pd.read_csv(r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2017_2020_live_mammals_resampled.csv")

s_2020 = df_s[df_s['Year'] == 2020]
c_2020 = df_c[df_c['Year'] == 2020]
l_2020 = df_l[df_l['Year'] == 2020]

print("2020 Strikes by Month:")
print(s_2020['Month'].value_counts().sort_index())
print("\n2020 Carcass by Month:")
print(c_2020['Month'].value_counts().sort_index())
print("\n2020 Live Animals by Month:")
print(l_2020['Month'].value_counts().sort_index())
