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
key_indicator_path = sys.argv[1]
finding_detail_path = sys.argv[2]
output_path = sys.argv[3]

# temp paths for local testing
# key_indicator_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Missionary KI Table (2).xlsx"
# finding_detail_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Detail (11).xlsx"
# output_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Output Graphs"

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

# Count number of first-time sacrament attendances within a week of first finding event per week (filter to last year)
df_fd['First Sacrament Date'] = pd.to_datetime(df_fd['First Sacrament Date'], errors='coerce')
df_fd['First Finding Event Date (truncated)'] = pd.to_datetime(df_fd['First Finding Event Date (truncated)'], errors='coerce')
df_fd['Confirmation Date'] = pd.to_datetime(df_fd['Confirmation Date'], errors='coerce')

# Filter for those who attended sacrament within a week of finding event
within_week_mask = (
    (df_fd['First Sacrament Date'].notnull()) &
    (df_fd['First Finding Event Date (truncated)'].notnull()) &
    ((df_fd['First Sacrament Date'] - df_fd['First Finding Event Date (truncated)']).dt.days >= 0) &
    ((df_fd['First Sacrament Date'] - df_fd['First Finding Event Date (truncated)']).dt.days < 8)
)

one_year_ago = pd.Timestamp.today() - pd.Timedelta(days=365)
today = pd.Timestamp.today()
date_mask = (df_fd['First Sacrament Date'] >= one_year_ago) & (df_fd['First Sacrament Date'] <= today)

attendance_within_week = (
    df_fd.loc[within_week_mask & date_mask, 'First Sacrament Date']
    .dropna()
    .dt.to_period('W')
    .value_counts()
    .sort_index()
)

# Count number of confirmations per week (within last year)
confirmation_date_mask = (df_fd['Confirmation Date'] >= one_year_ago) & (df_fd['Confirmation Date'] <= today)
confirmations_per_week = (
    df_fd.loc[confirmation_date_mask, 'Confirmation Date']
    .dropna()
    .dt.to_period('W')
    .value_counts()
    .sort_index()
)

# Align both series to have the same index (weeks)
all_weeks = attendance_within_week.index.union(confirmations_per_week.index)
attendance_within_week = attendance_within_week.reindex(all_weeks, fill_value=0)
confirmations_per_week = confirmations_per_week.reindex(all_weeks, fill_value=0)

# Dual y-axis plot
fig, ax1 = plt.subplots(figsize=(16, 9))

color1 = blue
color2 = red

ax1.bar(attendance_within_week.index.astype(str), attendance_within_week.values, color=color1, label='First-Time Sacrament Attendance (Within 1 Week)')
ax1.set_xlabel('Week')
ax1.set_ylabel('First Week Attendees', color=color1)
ax1.tick_params(axis='y', labelcolor=color1)
plt.xticks(rotation=90)

ax2 = ax1.twinx()
ax2.plot(confirmations_per_week.index.astype(str), confirmations_per_week.values, color=color2, marker='o', label='Confirmations per Week')
ax2.set_ylabel('Baptisms', color=color2)
ax2.tick_params(axis='y', labelcolor=color2)

plt.title('First Sacrament Attendance Within a Week of Being Found vs. Baptisms per Week', fontsize=16, fontweight='bold')
fig.tight_layout()
# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'first_week_sac_att_12mon_en.png'), dpi=100)
plt.close()

# Japanese version of the plot
rcParams['font.family'] = 'Yu Gothic'
fig, ax1 = plt.subplots(figsize=(16, 9))

color1 = blue
color2 = red

ax1.bar(attendance_within_week.index.astype(str), attendance_within_week.values, color=color1, label='First-Time Sacrament Attendance (Within 1 Week)')
ax1.set_xlabel('週')
ax1.set_ylabel('初めての聖餐会出席人数', color=color1)
ax1.tick_params(axis='y', labelcolor=color1)
plt.xticks(rotation=90)

ax2 = ax1.twinx()
ax2.plot(confirmations_per_week.index.astype(str), confirmations_per_week.values, color=color2, marker='o', label='Confirmations per Week')
ax2.set_ylabel('バプテスマ人数', color=color2)
ax2.tick_params(axis='y', labelcolor=color2)

plt.title('見つけられてから1週間以内の初めての聖餐会出席人数とバプテスマ人数の比較', fontsize=16, fontweight='bold')
fig.tight_layout()
# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'first_week_sac_att_12mon_jp.png'), dpi=100)
plt.close()