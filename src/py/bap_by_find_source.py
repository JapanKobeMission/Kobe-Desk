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

# === Compare Baptisms by Finding Source ===

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

# Year stuff
current_year = datetime.now().year
previous_year = current_year - 1
df_fd_confirmed['Confirmation Year'] = df_fd_confirmed['Confirmation Date'].dt.year
df_fd_confirmed = df_fd_confirmed[df_fd_confirmed['Confirmation Year'].isin([current_year, previous_year])] 

# Calculate percentage of baptisms by finding source
df_fd_confirmed['Finding Source'] = df_fd_confirmed['Finding Source'].fillna('Unknown')
finding_source_counts = df_fd_confirmed['Finding Source'].value_counts()
finding_source_percentages = df_fd_confirmed['Finding Source'].value_counts(normalize=True) * 100
finding_source_percentages = finding_source_percentages.reset_index()
finding_source_percentages.columns = ['Finding Source', 'Percentage'] 

# Plotting the percentage of baptisms by finding source
plt.figure(figsize=(16, 9))
ax = sns.barplot(
        data=finding_source_percentages,
        y='Finding Source',
        x='Percentage',
        palette='viridis',
        order=finding_source_percentages.sort_values('Percentage', ascending=False)['Finding Source']
    )

plt.xlabel('Percentage of Total (%)', fontsize=12)  # Percentage of Total (%)
plt.ylabel('Finding Source', fontsize=12)  # Finding Source
plt.title(f'Baptisms by Finding Source: {previous_year} - {current_year}', fontsize=14, fontweight='bold')    
plt.tight_layout()

# Add count labels to each bar
# Get the counts in the same order as the bars
ordered_sources = finding_source_percentages.sort_values('Percentage', ascending=False)['Finding Source']
counts_in_order = finding_source_counts[ordered_sources].values
for i, (bar, count) in enumerate(zip(ax.patches, counts_in_order)):
    width = bar.get_width()
    y = bar.get_y() + bar.get_height() / 2
    ax.text(width + 0.5, y, f'{count}', va='center', fontsize=11, color='black')

# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'bap_by_find_source_en.png'), dpi=100)
plt.close()

# Plotting the percentage of baptisms by finding source in Japanese
# Japanese translation for Finding Source
translation_dict = {
'Member': '会員からのリフェロー',
'Contacting in Public': '公の場でのコンタクト',
'Home to Home Contacting': '家でコンタクト',
'Through Person Being Taught': 'レッスンを受けている人を通じて',
'Headquarters Non-Paid Ad': '他のメディアファインディング',
'Facebook - Mission Ad': 'Facebook - 伝道部広告',
'English Class': '英会話クラス',
'Sought out Church or Missionaries': '教会または宣教師を自ら見つけた',
'Other': 'その他',
'New Member': '新会員を通じて',
'Ward Council': 'ワード調整集会',
# Add other translations as needed
}
df_fd_confirmed['Finding Source'] = df_fd_confirmed['Finding Source'].replace(translation_dict)
finding_source_counts = df_fd_confirmed['Finding Source'].value_counts()
finding_source_percentages = df_fd_confirmed['Finding Source'].value_counts(normalize=True) * 100
finding_source_percentages = finding_source_percentages.reset_index()
finding_source_percentages.columns = ['Finding Source', 'Percentage'] 

plt.figure(figsize=(16, 9))
ax = sns.barplot(
        data=finding_source_percentages,
        y='Finding Source',
        x='Percentage',
        palette='viridis',
        order=finding_source_percentages.sort_values('Percentage', ascending=False)['Finding Source']
    )

plt.xlabel('各種の割合 (%)', fontsize=12)  # Percentage of Total (%)
plt.ylabel('ファインディング方法', fontsize=12)  # Finding Source
plt.title(f'ファインディング方法別のバプテスマ割合: {previous_year}年と{current_year}年', fontsize=14, fontweight='bold')    
plt.tight_layout()

# Add count labels to each bar
# Get the counts in the same order as the bars
ordered_sources = finding_source_percentages.sort_values('Percentage', ascending=False)['Finding Source']
counts_in_order = finding_source_counts[ordered_sources].values
for i, (bar, count) in enumerate(zip(ax.patches, counts_in_order)):
    width = bar.get_width()
    y = bar.get_y() + bar.get_height() / 2
    ax.text(width + 0.5, y, f'{count}', va='center', fontsize=11, color='black')

# Save the plot
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'bap_by_find_source_jp.png'), dpi=100)
plt.close()
