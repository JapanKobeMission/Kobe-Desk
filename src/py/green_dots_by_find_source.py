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

# === Compare Green Dots by Finding Source ===

# drop unnecessary columns
df_fd = df_fd.drop(columns=['Sort', 
                            'Event Date Selected, Finding Type Group, Finding Category and 5 more (Combined)',
                            'Event Date Selected',
                            'First Finding Event Date (truncated)',
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

# Filter for green dots
# Green dots are defined as those with a 'First New Person Being Taught Date'
df_fd_green_dots = df_fd[df_fd['First New Person Being Taught Date'].notna()]
df_fd_green_dots['Green Dot Year'] = df_fd_green_dots['First New Person Being Taught Date'].dt.year
current_year = datetime.now().year
previous_year = current_year - 1

plt.figure(figsize=(12, 6))
sns.countplot(data=df_fd_green_dots, y='Finding Source', order=df_fd_green_dots['Finding Source'].value_counts().index, palette='viridis')
plt.title(f'Green Dots per Finding Source: {previous_year} - {current_year}', fontsize=14, fontweight='bold')
plt.xlabel('Green Dots', fontsize=12)
plt.ylabel('Finding Source', fontsize=12)
plt.tight_layout()
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'green_dots_by_find_source_en.png'))
plt.close()

plt.figure(figsize=(12, 6))
sns.countplot(data=df_fd_green_dots, y='Finding Source', order=df_fd_green_dots['Finding Source'].value_counts().index, palette='viridis')
plt.title(f'ファインディング方法によるグリーンドット: {previous_year}年と{current_year}年', fontsize=14, fontweight='bold')
plt.xlabel('グリーンドット', fontsize=12)
plt.ylabel('ファインディング方法', fontsize=12)
plt.tight_layout()
output_dir = os.path.join(output_path, datetime.now().strftime('%Y-%m-%d'))
os.makedirs(output_dir, exist_ok=True)
plt.savefig(os.path.join(output_dir, 'green_dots_by_find_source_jp.png'))
plt.close()