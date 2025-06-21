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

df_fd = read_data(finding_detail_path)

# === Compare Baptisms between Current Year and Previous Year ===

# drop unnecessary columns
df_fd = df_fd.drop(columns=['Sort', 
                            'Event Date Selected, Finding Type Group, Finding Category and 5 more (Combined)',
                            'Event Date Selected',
                            'First Finding Event Date (truncated)',
                            '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '16', '17', '18', '19', '20', '21', '22', '24', '25', '28', '32', '35'
                            ], errors='ignore')

# Convert all date columns to datetime
date_columns = ['First Referral Event Date', 
                'First Contact Attempt Event Date', 
                'First Successful Contact Attempt Event Date', 
                'First New Person Being Taught Date', 
                'First Lesson Date', 
                'Second Lesson Date', 
                'First Sacrament Date', 
                'Latest Sacrament Date', 
                'First Baptism Goal Date Set', 
                'Latest Baptism Goal Date Set', 
                'Confirmation Date', 
                'Event Type Date',
                ]
for col in date_columns:
    if col in df_fd.columns:
        df_fd[col] = pd.to_datetime(df_fd[col], errors='coerce')

# Filter for confirmed baptisms
df_fd_confirmed = df_fd[df_fd['Confirmation Date'].notna()]

# Extract year from Confirmation Date
df_fd_confirmed['Confirmation Year'] = df_fd_confirmed['Confirmation Date'].dt.year

# Filter for the current year and previous year
current_year = datetime.now().year
previous_year = current_year - 1
df_fd_confirmed = df_fd_confirmed[df_fd_confirmed['Confirmation Year'].isin([current_year, previous_year])]

# Prepare data for weekly cumulative sum comparison
df_fd_confirmed['Week'] = df_fd_confirmed['Confirmation Date'].dt.isocalendar().week
df_fd_confirmed['Year'] = df_fd_confirmed['Confirmation Date'].dt.year

# Group by year and week, count confirmations
weekly_counts = df_fd_confirmed.groupby(['Year', 'Week']).size().reset_index(name='Confirmations')

# Pivot to have weeks as index and years as columns
pivot_df = weekly_counts.pivot(index='Week', columns='Year', values='Confirmations').fillna(0)

# Only keep previous year and current year columns if they exist
years_to_plot = [previous_year, current_year]
pivot_df = pivot_df[[y for y in years_to_plot if y in pivot_df.columns]]

# Compute cumulative sum for each year
cumulative_df = pivot_df.cumsum()

# Only plot up to the current week
current_week = datetime.now().isocalendar()[1]
cumulative_df = cumulative_df.loc[cumulative_df.index <= current_week]

# Plot with predictions for end of year English
plt.figure(figsize=(16, 9))
for year in cumulative_df.columns:
    color1 = red if year == previous_year else blue if year == current_year else None
    plt.plot(cumulative_df.index, cumulative_df[year], marker='o', label=str(year), color=color1)
    # Linear regression
    x = cumulative_df.index.values
    y = cumulative_df[year].values
    if len(x) > 1:
        coef = np.polyfit(x, y, 1)
        poly1d_fn = np.poly1d(coef)
        color2 = light_red if year == previous_year else light_blue if year == current_year else None
        x_reg = np.arange(x.min(), 53)
        plt.plot(x_reg, poly1d_fn(x_reg), linestyle='--', label=f'{year} Trend', color=color2)
        # Predict value at week 52
        pred_week = 52
        pred_value = poly1d_fn(pred_week)
        plt.scatter([pred_week], [pred_value], color=color2, marker='x')
        plt.text(pred_week, pred_value, f'{int(pred_value):,}', color=color2, fontsize=10, va='bottom', ha='left')
        
plt.title(f'Baptism Count: {previous_year} vs {current_year}')
plt.xlabel('Week Number')
plt.ylabel('Baptisms')
plt.legend()
plt.tight_layout()
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'bap_year_comp_pred_en.png'), dpi=100)
plt.close()

# Plot with predictions for end of year Japanese
plt.figure(figsize=(16, 9))
for year in cumulative_df.columns:
    color1 = red if year == previous_year else blue if year == current_year else None
    plt.plot(cumulative_df.index, cumulative_df[year], marker='o', label=str(year), color=color1)
    # Linear regression
    x = cumulative_df.index.values
    y = cumulative_df[year].values
    if len(x) > 1:
        coef = np.polyfit(x, y, 1)
        poly1d_fn = np.poly1d(coef)
        color2 = light_red if year == previous_year else light_blue if year == current_year else None
        x_reg = np.arange(x.min(), 53)
        plt.plot(x_reg, poly1d_fn(x_reg), linestyle='--', label=f'{year} 動向', color=color2)
        # Predict value at week 52
        pred_week = 52
        pred_value = poly1d_fn(pred_week)
        plt.scatter([pred_week], [pred_value], color=color2, marker='x')
        plt.text(pred_week, pred_value, f'{int(pred_value):,}', color=color2, fontsize=10, va='bottom', ha='left')
        
plt.title(f'バプテスマを受けた人数の比較: {previous_year}年と{current_year}年')
plt.xlabel('週番号')
plt.ylabel('バプテスマを受けた人数')
plt.legend()
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'bap_year_comp_pred_jp.png'), dpi=100)
plt.close()