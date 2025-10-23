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
key_indicator_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Missionary KI Table.xlsx"
finding_detail_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Detail.xlsx"
output_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Output Graphs"

def read_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(file_path)
    elif ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

df_fd = read_data(finding_detail_path)

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

def print_confirmation_percentage(df, label):
    total = len(df)
    if total == 0:
        percent = 0
    else:
        percent = df['Confirmation Date'].notnull().sum() / total * 100
    print(f"{label}: {percent:.2f}% have a confirmation date ({df['Confirmation Date'].notnull().sum()}/{total})")

print_confirmation_percentage(df_fd_true, "First Week Sac Att True")
print_confirmation_percentage(df_fd_true_missionary, "First Week Sac Att True - Missionary")
print_confirmation_percentage(df_fd_true_member, "First Week Sac Att True - Member")
print_confirmation_percentage(df_fd_true_media, "First Week Sac Att True - Media")
print_confirmation_percentage(df_fd_false, "First Week Sac Att False")
print_confirmation_percentage(df_fd_false_missionary, "First Week Sac Att False - Missionary")
print_confirmation_percentage(df_fd_false_member, "First Week Sac Att False - Member")
print_confirmation_percentage(df_fd_false_media, "First Week Sac Att False - Media")