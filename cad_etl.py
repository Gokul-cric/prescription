import pandas as pd
import numpy as np
from datetime import datetime, time,timedelta
import os
import json
import warnings
from datetime import time
warnings.filterwarnings("ignore")

cd=os.getcwd()
data_dir = os.path.join(cd,"Data")
config_dir=os.path.join(cd,'Config')
config = os.path.join(config_dir,"config.json")

with open(config, 'r') as file:
    config = json.load(file)



def get_shift_data_with_proportional_split(df, shift, date,config):

    shift_start_str, shift_end_str = config["shift_time"][shift]

    if shift_end_str < time(4,0,0):
        previous_day = date + timedelta(days=1)
        shift_start = pd.to_datetime(f"{date} {shift_start_str}") 
        shift_end = pd.to_datetime(f"{previous_day} {shift_end_str}") 

    elif shift_start_str > shift_end_str:
        previous_day = date - timedelta(days=1)
        shift_start = pd.to_datetime(f"{previous_day} {shift_start_str}") 
        shift_end = pd.to_datetime(f"{date} {shift_end_str}") 
    else:
        shift_start = pd.to_datetime(f"{date} {shift_start_str}")
        shift_end = pd.to_datetime(f"{date} {shift_end_str}")


    df['StartDateTime'] = pd.to_datetime(df['Date'].dt.strftime('%Y-%m-%d') + ' ' + df['StartTime'].astype(str))
    df['EndDateTime'] = pd.to_datetime(df['Date'].dt.strftime('%Y-%m-%d') + ' ' + df['EndTime'].astype(str))
    
    df.loc[df['EndDateTime'] < df['StartDateTime'], 'EndDateTime'] += timedelta(days=1)
    df.loc[(df["EndDateTime"].dt.time<=time(4,0,0))&(df["StartDateTime"].dt.time<=time(4,0,0)),"EndDateTime"]+= timedelta(days=1)
    df.loc[(df["EndDateTime"].dt.time<=time(4,0,0))&(df["StartDateTime"].dt.time<=time(4,0,0)),"StartDateTime"]+= timedelta(days=1)

    mask = (df['StartDateTime'] < shift_end) & (df['EndDateTime'] > shift_start)
    # previous_day_data = df[(df['Date'] == (date - timedelta(days=1)))]

    # if shift_start_str == '15:00:00' and previous_day_data.empty:
    #     exclude_start_time = pd.to_datetime(f"{date} 00:00:00")
    #     exclude_end_time = pd.to_datetime(f"{date} 04:00:00")

    #     same_day_mask = (df['StartDateTime'] >= exclude_start_time) & (df['StartDateTime'] < exclude_end_time)
    #     mask &= ~same_day_mask
    # if shift_start_str == '22:30:00' and previous_day_data.empty:
    #     exclude_start_time = pd.to_datetime(f"{date} 00:00:00")
    #     exclude_end_time = pd.to_datetime(f"{date} 04:00:00")
    #     same_day_mask = (df['StartDateTime'] >= exclude_start_time) & (df['StartDateTime'] < exclude_end_time)
    #     mask &= ~same_day_mask

    # if not previous_day_data.empty:
    #        include_start_time = pd.to_datetime(f"{date - timedelta(days=1)} 00:00:00")  # Adjusted to previous day
    #        include_end_time = pd.to_datetime(f"{date - timedelta(days=1)} 04:00:00") 

    #        same_day_mask = (df['StartDateTime'] >= include_start_time) & (df['StartDateTime'] < include_end_time)
    #        mask |= same_day_mask 

    shift_data = df[mask].copy()
    for index, row in shift_data.iterrows():
        total_duration = (row['EndDateTime'] - row['StartDateTime']).total_seconds()

        in_shift_start = max(row['StartDateTime'], shift_start)
        in_shift_end = min(row['EndDateTime'], shift_end)
        in_shift_duration = (in_shift_end - in_shift_start).total_seconds()

        if in_shift_duration > 0:
            ratio = in_shift_duration / total_duration
            numerical_columns = df.select_dtypes(include=['number']).columns.tolist()
            numerical_columns.remove('Recycle Sand / Batch (Kg)')
            shift_data.loc[index, numerical_columns] *= ratio
  
    shift_data['No of Boxes'] = shift_data['No of Boxes'].round(0)
    shift_data['No of Batches'] = shift_data['No of Batches'].round(0)
    shift_data['Unpoured Moulds (Nos)'] = shift_data['Unpoured Moulds (Nos)'].round(0)

    return shift_data.reset_index(drop=True)

def calculate_data_for_all_shifts(df,config):
    unique_dates = df['Date'].dt.date.unique()  
    all_shift_data = []

    for date in unique_dates:
        for shift in config["shift_time"].keys():
            shift_data = get_shift_data_with_proportional_split(df, shift, pd.to_datetime(date),config)

            if not shift_data.empty:
                shift_data['Date'] = date
                shift_data['Shift'] = shift

                all_shift_data.append(shift_data)

    return pd.concat(all_shift_data, ignore_index=True) if all_shift_data else pd.DataFrame()

#shift_summaries_df = calculate_data_for_all_shifts(df)



