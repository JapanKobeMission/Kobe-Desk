import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt 
from matplotlib import rcParams
from datetime import datetime
import sys
import os

# pandas options
pd.set_option("display.max_columns", None)  # Show up to 50 columns
pd.set_option("display.max_rows", None)
pd.set_option('display.width', 1000)

# seaborn options
sns.set_theme(style="darkgrid")
rcParams['font.family'] = 'Yu Gothic'

# color storage
red = '#ff0000'
light_red = '#ff698c'
blue = '#0011ff'
light_blue = '#0576ff'

# sys.argv is the array of arguments passed to the script from the js
# key_indicator_path = sys.argv[1]
# finding_detail_path = sys.argv[2]
# output_path = sys.argv[3]

# temp paths for local testing
key_indicator_path = "C:\\Users\\2016702-REF\\Downloads\\Missionary KI Table (1).xlsx"
finding_detail_path = "C:\\Users\\2016702-REF\\Downloads\\Detail (7).xlsx"
output_path = "C:\\Users\\2016702-REF\\VSCode Python Projects\\Kobe Desk swaHekuL\\Kobe-Desk\\src\\output"

def read_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(file_path)
    elif ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

df_fd = read_data(finding_detail_path)
print(df_fd.columns)

# === Look at statistics concering first week sacrament attendance YTD ===
df_fd['First Week Sac Att?'] = np.where(
    (pd.to_datetime(df_fd['First Finding Event Date (truncated)']) >= pd.to_datetime(df_fd['First Sacrament Date'])) |
    ((pd.to_datetime(df_fd['First Sacrament Date']) - pd.to_datetime(df_fd['First Finding Event Date (truncated)'])).dt.days < 8),
    True,
    False
)

df_fd_true = df_fd[df_fd['First Week Sac Att?'] == True]
df_fd_true_missionary = df_fd_true[df_fd_true['Finding Category (copy)'] == 'Missionary']
df_fd_true_member = df_fd_true[df_fd_true['Finding Category (copy)'] == 'Member']
df_fd_true_media = df_fd_true[df_fd_true['Finding Category (copy)'] == 'Media']
df_fd_false = df_fd[df_fd['First Week Sac Att?'] == False]
df_fd_false_missionary = df_fd_false[df_fd_false['Finding Category (copy)'] == 'Missionary']
df_fd_false_member = df_fd_false[df_fd_false['Finding Category (copy)'] == 'Member']
df_fd_false_media = df_fd_false[df_fd_false['Finding Category (copy)'] == 'Media']

print(f'% baptized who attended sacrament meeting in the first week: {len(df_fd_true[df_fd_true['Confirmation Date'].notnull()]) / len(df_fd_true) * 100:.2f}%')
print(f'% baptized who attended sacrament meeting in the first week (missionary): {len(df_fd_true_missionary[df_fd_true_missionary['Confirmation Date'].notnull()]) / len(df_fd_true_missionary) * 100:.2f}%')
print(f'% baptized who attended sacrament meeting in the first week (member): {len(df_fd_true_member[df_fd_true_member['Confirmation Date'].notnull()]) / len(df_fd_true_member) * 100:.2f}%')
print(f'% baptized who attended sacrament meeting in the first week (media): {len(df_fd_true_media[df_fd_true_media['Confirmation Date'].notnull()]) / len(df_fd_true_media) * 100:.2f}%')
print(f'% baptized who did not attend sacrament meeting in the first week: {len(df_fd_false[df_fd_false['Confirmation Date'].notnull()]) / len(df_fd_false) * 100:.2f}%')
print(f'% baptized who did not attend sacrament meeting in the first week (missionary): {len(df_fd_false_missionary[df_fd_false_missionary['Confirmation Date'].notnull()]) / len(df_fd_false_missionary) * 100:.2f}%')
print(f'% baptized who did not attend sacrament meeting in the first week (member): {len(df_fd_false_member[df_fd_false_member['Confirmation Date'].notnull()]) / len(df_fd_false_member) * 100:.2f}%')
print(f'% baptized who did not attend sacrament meeting in the first week (media): {len(df_fd_false_media[df_fd_false_media['Confirmation Date'].notnull()]) / len(df_fd_false_media) * 100:.2f}%')