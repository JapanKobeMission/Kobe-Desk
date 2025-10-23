import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt 
from matplotlib import rcParams
from datetime import datetime, timedelta
import os
import sys
import numpy as np

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
# key_indicator_path = "C:\\Users\\2016702-REF\\Downloads\\Missionary KI Table (1).xlsx"
# finding_detail_path = "C:\\Users\\2016702-REF\\Downloads\\Detail (7).xlsx"
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
df_fd = df_fd.query("`Finding Category (copy)` == 'Member'") # Filter for 'Member' category in 'Finding Category (copy)' column

# columns ['Sort', 'Event Date Selected, Finding Type Group, Finding Category and 5 more (Combined)', 'Event Date Selected', 'First Finding Event Date (truncated)', 'Latest Zone Name', 'Latest District Name', 'Latest Teaching Area Name', 'Finding Category (copy)', 'Finding Source', 'Full Name', 'Person Id', 'First Referral Event Date', 'First Contact Attempt Event Date', 'First Successful Contact Attempt Event Date', 'First New Person Being Taught Date', 'First Lesson Date', 'Second Lesson Date', 'First Sacrament Date', 'Latest Sacrament Date', 'First Baptism Goal Date Set', 'Latest Baptism Goal Date Set', 'Confirmation Date', 'Event Type Date', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '12', '16', '17', '20', '21', '23', '24']
df_fd.drop(columns=['Sort', 'Event Date Selected, Finding Type Group, Finding Category and 5 more (Combined)', 'Event Date Selected', 'Finding Category (copy)', 'Finding Source',  'First Referral Event Date', 'First Contact Attempt Event Date', 'First Successful Contact Attempt Event Date', 'First New Person Being Taught Date', 'First Lesson Date', 'Second Lesson Date', 'First Sacrament Date', 'Latest Sacrament Date', 'First Baptism Goal Date Set', 'Latest Baptism Goal Date Set', 'Confirmation Date', 'Event Type Date', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '12', '16', '17', '20', '21', '23', '24'], inplace=True)  # Drop unnecessary columns
df_fd.rename(columns={
    'Latest Zone Name': 'Zone',
    'Latest District Name': 'District',
    'Latest Teaching Area Name': 'Area',
    'First Finding Event Date (truncated)': 'First Finding Event Date',
}, inplace=True)  # Rename columns for clarity and consistency
df_fd['First Finding Event Date'] = pd.to_datetime(df_fd['First Finding Event Date'], errors='coerce')  # Convert 'First Finding Event Date' column to datetime
df_fd['Previous Sunday Date'] = df_fd['First Finding Event Date'].apply(lambda x: x - timedelta(days=x.weekday() + 1))  # Calculate the previous Sunday date
df_fd['Previous Sunday Date'] = pd.to_datetime(df_fd['Previous Sunday Date'], errors='coerce')  # Ensure the new column is in datetime format
df_fd = df_fd[['First Finding Event Date', 'Previous Sunday Date', 'Full Name', 'Person Id', 'Zone', 'District', 'Area']]

# Load the Key Indicators Excel file
df_ki = read_data(key_indicator_path)

# Rename columns for clarity and consistency
# columns ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'New People Goal', 'New People Actual', 'Lessons with Member Participating Goal', 'Lessons with Member Participating Actual', 'Potential Member Sacrament Goal', 'Potential Member Sacrament Actual', 'Has Baptismal Date Goal', 'Has Baptismal Date Actual', 'Baptized and Confirmed Goal', 'Baptized and Confirmed Actual', 'New Members at Sacrament Goal', 'New Members at Sacrament Actual']
df_ki = df_ki.rename(columns={
    'Unnamed: 0': 'Ignore 0',
    'Unnamed: 1': 'Ignore 1',
    'Unnamed: 2': 'Ignore 2',
    'Unnamed: 3': 'Sunday Date',
    'Unnamed: 4': 'Zone',
    'Unnamed: 5': 'District',
    'Unnamed: 6': 'Area',
    'Unnamed: 7': 'Missionaries',
    'New People Goal': 'NP Goal',
    'New People Actual': 'NP Actual',
    'Lessons with Member Participating Goal': 'ML Goal',
    'Lessons with Member Participating Actual': 'ML Actual',
    'Potential Member Sacrament Goal': 'SA Goal',
    'Potential Member Sacrament Actual': 'SA Actual',
    'Has Baptismal Date Goal': 'BD Goal',
    'Has Baptismal Date Actual': 'BD Actual',
    'Baptized and Confirmed Goal': 'BC Goal',
    'Baptized and Confirmed Actual': 'BC Actual',
    'New Members at Sacrament Goal': 'NMSA Goal',
    'New Members at Sacrament Actual': 'NMSA Actual',
    # Add more mappings as needed based on your actual column names
})
df_ki.drop(columns=['Ignore 0', 'Ignore 1', 'Ignore 2', 'Missionaries'], inplace=True)  # Drop unnecessary columns
df_ki['Sunday Date'] = pd.to_datetime(df_ki['Sunday Date'], errors='coerce')  # Convert 'Date' column to datetime
df_ki = df_ki.fillna(0)  # Fill NaN values with 0 for numerical columns

# Group df_fd by 'Previous Sunday Date' and count unique 'Person Id' (member referrals per week)
mr_counts = df_fd.groupby('Previous Sunday Date')['Person Id'].nunique().reset_index()
mr_counts.rename(columns={'Person Id': 'Member Referrals'}, inplace=True)

# Group df_ki by 'Sunday Date' and sum 'SA Actual' (sacrament attendance per week)
sa_counts = df_ki.groupby('Sunday Date')['SA Actual'].sum().reset_index()

# Merge the two DataFrames on the date columns (aligning weeks)
comparison_df = pd.merge(
    mr_counts, sa_counts,
    left_on='Previous Sunday Date', right_on='Sunday Date',
    how='outer'
).sort_values('Previous Sunday Date').reset_index(drop=True)

# Filter dates: from one year ago (from today) to today
one_year_ago = pd.Timestamp('today').normalize() - pd.DateOffset(years=1)
today = pd.Timestamp('today').normalize()
comparison_df = comparison_df[
    (comparison_df['Previous Sunday Date'] >= one_year_ago) &
    (comparison_df['Previous Sunday Date'] <= today)
]

# Prepare for plotting
comparison_df['Date'] = comparison_df['Previous Sunday Date'].combine_first(comparison_df['Sunday Date'])

# For Member Referrals
x = np.arange(len(comparison_df['Date']))
y1 = comparison_df['Member Referrals'].fillna(0)
z1 = np.polyfit(x, y1, 2)
p1 = np.poly1d(z1)

# For Sacrament Attendance
y2 = comparison_df['SA Actual'].fillna(0)
z2 = np.polyfit(x, y2, 2)
p2 = np.poly1d(z2)

# Double y-axis plot
fig, ax1 = plt.subplots(figsize=(16, 9))

color1 = 'tab:green'
ax1.set_xlabel('Week')
ax1.set_ylabel('Finds Through Members', color=color1)
ax1.plot(comparison_df['Date'], comparison_df['Member Referrals'], marker='o', color=color1, label='Member Referrals')
ax1.tick_params(axis='y', labelcolor=color1)
ax1.plot(comparison_df['Date'], p1(x), color=color1, linestyle='--', linewidth=3, label='Member Referrals Trend')

ax2 = ax1.twinx()
color2 = 'tab:orange'
ax2.set_ylabel('Sacrament Attendance', color=color2)
ax2.plot(comparison_df['Date'], comparison_df['SA Actual'], marker='s', color=color2, label='SA Actual')
ax2.tick_params(axis='y', labelcolor=color2)
ax2.plot(comparison_df['Date'], p2(x), color=color2, linestyle='--', linewidth=3, label='SA Actual Trend')

plt.title('Weekly Comparison: Finds Through Members vs. Sacrament Attendance', fontsize=16, fontweight='bold')

fig.tight_layout()

# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'sac_att_vs_member_ref_ytd_en.png'), dpi=100)
plt.close()

# Japanese version of the plot
fig, ax1 = plt.subplots(figsize=(16, 9))

color1 = 'tab:green'
ax1.set_xlabel('週')
ax1.set_ylabel('会員からのリフェロー', color=color1)
ax1.plot(comparison_df['Date'], comparison_df['Member Referrals'], marker='o', color=color1, label='Member Referrals')
ax1.tick_params(axis='y', labelcolor=color1)
ax1.plot(comparison_df['Date'], p1(x), color=color1, linestyle='--', linewidth=3, label='Member Referrals Trend')

ax2 = ax1.twinx()
color2 = 'tab:orange'
ax2.set_ylabel('聖餐会の出席人数', color=color2)
ax2.plot(comparison_df['Date'], comparison_df['SA Actual'], marker='s', color=color2, label='SA Actual')
ax2.tick_params(axis='y', labelcolor=color2)
ax2.plot(comparison_df['Date'], p2(x), color=color2, linestyle='--', linewidth=3, label='SA Actual Trend')

plt.title('週ごとの比較：会員からのリフェローと聖餐会の出席人数', fontsize=16, fontweight='bold')
fig.tight_layout()

# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'sac_att_vs_member_ref_ytd_jp.png'), dpi=100)
plt.close()
