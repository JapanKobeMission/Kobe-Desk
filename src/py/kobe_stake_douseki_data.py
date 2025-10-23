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
# key_indicator_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Missionary KI Table (3).xlsx"
# finding_detail_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Detail (12).xlsx"
# output_path = r"C:\Users\2016702-REF\OneDrive - Church of Jesus Christ\Desktop\Kobe Desk\Output Graphs"

def read_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(file_path)
    elif ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

df_ki = read_data(key_indicator_path)

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

print(f'Columns in Key Indicator DataFrame: {df_ki.columns.tolist()}')

# Filter for Kobe zone and Toyooka area
kobe_df = df_ki[(df_ki['Zone'] == 'Kobe') | (df_ki['Area'] == 'Toyooka')]

# Convert 'Sunday Date' to datetime if not already
kobe_df['Sunday Date'] = pd.to_datetime(kobe_df['Sunday Date'], errors='coerce')

# Filter for dates between start of the year and today
start_date = pd.to_datetime(f"{pd.Timestamp.now().year}-01-01")
end_date = pd.to_datetime('today').normalize()
kobe_df = kobe_df[(kobe_df['Sunday Date'] >= start_date) & (kobe_df['Sunday Date'] <= end_date)]

ml_totals = kobe_df.groupby('Area')['ML'].sum()
print(f"Total Dousekis in Kobe Stake from {start_date} to {end_date}: ")
print(ml_totals)

txt_path = os.path.join(output_path, 'kobe_stake_douseki_totals.txt')
with open(txt_path, 'w', encoding='utf-8') as f:
    f.write(f"Total Dousekis in Kobe Stake from {start_date.date()} to {end_date.date()}:\n")
    for area, total in ml_totals.items():
        f.write(f"{area}: {total}\n")