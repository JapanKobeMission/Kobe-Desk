import pandas as pd
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
# key_indicator_path = "C:\\Users\\2016702-REF\\Downloads\\Missionary KI Table (1).xlsx"
# finding_detail_path = "C:\\Users\\2016702-REF\\Downloads\\Detail (7).xlsx"
# output_path = "C:\\Users\\2016702-REF\\VSCode Python Projects\\Kobe Desk swaHekuL\\Kobe-Desk\\src\\output"

def read_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(file_path)
    elif ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

df_ki = read_data(key_indicator_path)

# === Compare Sacrament Attendance between Current Year and Previous Year ===
# rename columns for clarity
df_ki = df_ki.rename(columns={
    'Unnamed: 0': 'Week Index',
    'Unnamed: 1': 'Weekly Sunday Date',
    'Unnamed: 2': 'Area ID',
    'Unnamed: 3': 'Sunday Date',
    'Unnamed: 4': 'Zone',
    'Unnamed: 5': 'District',
    'Unnamed: 6': 'Area',
    'Unnamed: 7': 'Missionaries',
    'New People Goal': 'NPG',
    'New People Actual': 'NP',
    'Lessons with Member Participating Goal': 'MLG', 
    'Lessons with Member Participating Actual': 'ML', 
    'Potential Member Sacrament Goal': 'SAG', 
    'Potential Member Sacrament Actual': 'SA', 
    'Has Baptismal Date Goal': 'BDG', 
    'Has Baptismal Date Actual': 'BD', 
    'Baptized and Confirmed Goal': 'BCG', 
    'Baptized and Confirmed Actual': 'BC', 
    'New Members at Sacrament Goal': 'NMSAG', 
    'New Members at Sacrament Actual': 'NMSA',
})
# drop unnecessary columns
df_ki = df_ki.drop(columns=['Week Index',
                            'Weekly Sunday Date',
                            'Area ID',
                            'Missionaries'
                            ], errors='ignore')

# Ensure 'Sunday Date' is datetime
df_ki['Sunday Date'] = pd.to_datetime(df_ki['Sunday Date'], errors='coerce')

# Add 'Week of Year' column (ISO week number)
df_ki['Week'] = df_ki['Sunday Date'].dt.isocalendar().week
df_ki['Year'] = df_ki['Sunday Date'].dt.year

# Filter for the current year and previous year
current_year = datetime.now().year
previous_year = current_year - 1
df_ki = df_ki[df_ki['Year'].isin([current_year, previous_year])]

# Group by 'Zone', 'District', 'Area', 'Week', and 'Year' to get weekly sums
df_ki_grouped = df_ki.drop(columns=['Sunday Date'], errors='ignore') \
    .groupby(['Zone', 'District', 'Area', 'Week', 'Year']).sum().reset_index()

# Pivot the DataFrame to have separate columns for each year
df_ki_pivot = df_ki_grouped.pivot_table(index=['Zone', 'District', 'Area', 'Week'], 
                                         columns='Year', 
                                         values=['NPG', 'NP', 'MLG', 'ML', 'SAG', 'SA', 
                                                 'BDG', 'BD', 'BCG', 'BC', 'NMSAG', 'NMSA'], 
                                         fill_value=0,
                                         aggfunc='sum'
                                         ).reset_index()

# Flatten the MultiIndex columns
df_ki_pivot.columns = [' '.join(map(str, col)).strip() for col in df_ki_pivot.columns.values]

# Group by 'Week' and sum across all areas/zones/districts
df_weekly_sum = df_ki_pivot.groupby('Week')[[col for col in df_ki_pivot.columns if col.startswith('SA')]].sum().reset_index()

# Plot comparison of SA between previous year and current year, only up to the current week, English
years = [previous_year, current_year]
sa_cols = [f'SA {year}' for year in years if f'SA {year}' in df_ki_pivot.columns]

if len(sa_cols) == 2:
    current_week = datetime.now().isocalendar()[1]
    df_plot = df_weekly_sum[df_weekly_sum['Week'] <= current_week]
    plt.figure(figsize=(12, 6))
    sns.lineplot(
        data=df_plot,
        x='Week',
        y=sa_cols[0],
        label=f'Sacrament Attendance {years[0]}',
        color=red,
        errorbar=None
    )
    sns.lineplot(
        data=df_plot,
        x='Week',
        y=sa_cols[1],
        label=f'Sacrament Attendance {years[1]}',
        color=blue,
        errorbar=None
    )
    plt.title(f'Sacrament Attendance: {previous_year} vs {current_year}')
    plt.xlabel('Week Number')
    plt.ylabel('Sacrament Attendance')
    plt.legend()
    plt.tight_layout()
    output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(os.path.join(output_dir, 'sac_att_year_comp_en.png'))
    plt.close()

# Plot comparison of SA between previous year and current year, only up to the current week, Japanese
if len(sa_cols) == 2:
    current_week = datetime.now().isocalendar()[1]
    df_plot = df_weekly_sum[df_weekly_sum['Week'] <= current_week]
    plt.figure(figsize=(12, 6))
    sns.lineplot(
        data=df_plot,
        x='Week',
        y=sa_cols[0],
        label=f'聖餐会の出席 {years[0]}年',
        color=red,
        errorbar=None
    )
    sns.lineplot(
        data=df_plot,
        x='Week',
        y=sa_cols[1],
        label=f'聖餐会の出席 {years[1]}年',
        color=blue,
        errorbar=None
    )
    plt.title(f'聖餐会の出席: {previous_year}年と{current_year}年')
    plt.xlabel('週番号')
    plt.ylabel('聖餐会の出席')
    plt.legend()
    plt.tight_layout()
    output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(os.path.join(output_dir, 'sac_att_year_comp_jp.png'))
    plt.close()