import pandas as pd
import os
import json
import numpy as np
from datetime import datetime, timedelta,time
from preprocess import load_all_configs


def date_adjust(data: pd.DataFrame, col_name: str, to_foundry: bool, shift_time):
    multipler = -1 if not to_foundry else 1
    data["DateTime"] = np.where(data[col_name] < datetime.strptime("04:00:00", "%H:%M:%S").time(),
                                 data["DateTime"] + pd.Timedelta(days=-1 * multipler),
                                 data["DateTime"])
    return data

def is_between(time, start, end):
    if start < end:  # same day shift
        return start <= time < end
    else:  # overnight shift
        return time >= start or time < end

def process_shift_data(data, config) -> pd.DataFrame:
    """
    Process the shift data by assigning shifts, adjusting dates, and imputing missing values.

    Parameters:
    - data_path: Path to the input Excel data file.
    - config_path: Path to the configuration JSON file containing shift times.
    - results_dir: Directory to save the aggregated results.

    Returns:
    - df: Final DataFrame with imputed values.
    """

    data['Time'] = pd.to_datetime(data['Time'], format='%H:%M').dt.time
    data["DateTime"] = pd.to_datetime(data["Date"].astype(str) + " " + data["Time"].astype(str))


    data = date_adjust(data, "Time", False, config["shift_time"])

    shifts = config["shift_time"]
    for shift in shifts:
        shifts[shift] = [datetime.strptime(t, "%H:%M:%S").time() for t in shifts[shift]]

    new_rows = []
    for _, row in data.iterrows():
        row_time = row['DateTime'].time()
        row_date = row['DateTime'].date()

        applicable_shifts = []
        for shift, (start, end) in shifts.items():
            if is_between(row_time, start, end):
                if end >start:
                    shift_date = row_date if row_time >= start else row_date - timedelta(days=1)
                    applicable_shifts.append((shift, shift_date))
                else:
                    shift_date = row_date + timedelta(days=1) if row_time >= start else row_date 
                    applicable_shifts.append((shift, shift_date))

                # if end >start:
                #     shift_date = row_date 
                #     applicable_shifts.append((shift, shift_date))
                # else:
                #     shift_date = row_date 
                #     applicable_shifts.append((shift, shift_date))

        for shift, shift_date in applicable_shifts:
            new_row = row.copy()
            new_row['ShiftedShift'] = shift
            new_row['ShiftDate'] = shift_date
            new_rows.append(new_row)

    

    shifted_df = pd.DataFrame(new_rows)

    for shift, (start, end) in shifts.items():
        if end < time(4,0,0):
            shifted_df.loc[shifted_df["ShiftedShift"]==shift,"ShiftDate"]=shifted_df.loc[shifted_df["ShiftedShift"]==shift,"ShiftDate"]-timedelta(days =1)

    aggregated_data = shifted_df.groupby(['ShiftDate', 'ShiftedShift']).mean().reset_index()
    # aggregated_data.loc[(aggregated_data["ShiftedShift"] == "B") | (aggregated_data["ShiftedShift"] == "A"), "ShiftDate"] += timedelta(days=1)
    aggregated_data[["GFN (no)", "LOI (%)"]] = aggregated_data[["GFN (no)", "LOI (%)"]].ffill()
    aggregated_data[["GFN (no)", "LOI (%)"]] = aggregated_data[["GFN (no)", "LOI (%)"]].bfill()
    aggregated_data.sort_values(by=['ShiftDate', 'ShiftedShift'], ascending=True, inplace=True)

    #aggregated_data.to_excel(os.path.join(results_dir, "Prepared Sand aggg.xlsx"), index=False)

    shifted_df.drop(["Date", "Shift", "DateTime"], axis=1, inplace=True)
    shifted_df.rename(columns=dict(zip(['ShiftDate', 'ShiftedShift'], ["Date", "Shift"])), inplace=True)
    # shifted_df.loc[(shifted_df["Shift"] == "B") | (shifted_df["Shift"] == "A"), "Date"] += timedelta(days=1)

    data = shifted_df.copy()
    data.reset_index(drop=True, inplace=True)

    lc = data.groupby(['Date', 'Shift'], group_keys=False, as_index=False)[['GFN (no)', 'LOI (%)']].mean().sort_values(by=['Date', "Shift"])
    lc.reset_index(drop=True, inplace=True)

    df = data.groupby(['Date', 'Shift'], group_keys=False).apply(lambda group: impute_last_record(group, lc))





    return df


def impute_last_record(group, lc):
    if group['GFN (no)'].isnull().all():
        date = group['Date']
        shift = group['Shift']
        index_new = lc[(lc['Date'] == date.iloc[0]) & (lc['Shift'] == shift.iloc[0])].index[0]
        temp_data = lc.loc[:index_new, :]
        temp_data = temp_data['GFN (no)'].dropna()
        gfn_value = temp_data.iloc[-1] if not temp_data.empty else np.nan
        last_index = group.index[-1]
        if pd.isna(group.at[last_index, 'GFN (no)']):
            group.at[last_index, 'GFN (no)'] = gfn_value

    if group['LOI (%)'].isnull().all():
        date = group['Date']
        shift = group['Shift']
        index_new = lc[(lc['Date'] == date.iloc[0]) & (lc['Shift'] == shift.iloc[0])].index[0]
        temp_data = lc.loc[:index_new, :]
        temp_data = temp_data['LOI (%)'].dropna()
        loi = temp_data.iloc[-1] if not temp_data.empty else np.nan
        last_index = group.index[-1]
        if pd.isna(group.at[last_index, 'LOI (%)']):
            group.at[last_index, 'LOI (%)'] = loi

    return group


