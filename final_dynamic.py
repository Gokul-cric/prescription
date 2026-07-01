import pandas as pd
from collections import Counter
from preprocess import load_all_configs
import os

configs=load_all_configs()    
config1 = configs.get("config.json")

cwd=os.getcwd()
data_dir = os.path.join(cwd,"Data",config1["foundry_name"])
ressults_dir = os.path.join(cwd,"Model Results",config1["foundry_name"])
analysis_dir = os.path.join(cwd,"Analysis",config1["foundry_name"])


df = pd.read_excel(os.path.join(data_dir,"Test Data","Consumptionbooking_01-Dec-2024_TO_20-Jan-2025_Spomatic.xlsx"),skiprows=5)
df['No of Boxes'] = df['Total Mould Made'] - df['Unpoured Mould']
df = df[['Date', 'Shift', 'ComponentId', 'StartTime', 'EndTime', 'No of Boxes']]
# df = df[df['StartTime'] != df['EndTime']]


df['DateShift'] = df['Date'].astype(str) + "_" + df['Shift']


df2 = pd.read_excel(os.path.join(data_dir,"Test Data","Consumption_01-Dec-2024_TO_20-Jan-2025_Spomatic.xlsx"),skiprows=5)
df2['StartTime'] = ''
df2['EndTime'] = ''


df2['DateShift'] = df2['Date'].astype(str) + "_" + df2['Shift']

processed_dfs = []


unique_dateshifts = df['DateShift'].unique()
for dateshift in unique_dateshifts:
   
    df_shift = df[df['DateShift'] == dateshift]
    df2_shift = df2[df2['DateShift'] == dateshift]

    component_counts = Counter(df_shift['ComponentId'])
    
 
    unique_components_df = df_shift[df_shift['ComponentId'].map(component_counts) == 1]
    repeated_components_df = df_shift[df_shift['ComponentId'].map(component_counts) > 1]
 
    unique_components_dict = unique_components_df.set_index('ComponentId').T.to_dict('dict')
    repeated_components_dict = repeated_components_df.groupby('ComponentId').apply(
        lambda x: x.drop('ComponentId', axis=1).to_dict('records')
    ).to_dict()
    
 
    for index, row in df2_shift.iterrows():
        component_id = row['Component ID']
        if component_id in unique_components_dict:
            df2_shift.at[index, 'StartTime'] = unique_components_dict[component_id]['StartTime']
            df2_shift.at[index, 'EndTime'] = unique_components_dict[component_id]['EndTime']
    
    if len(repeated_components_df)>0:
        for component_id, time_blocks in repeated_components_dict.items():
            component_rows = df2_shift[df2_shift['Component ID'] == component_id]
            cumulative_sum = 0
            start_index = None
            
            for i, row in component_rows.iterrows():
                if start_index is None:
                    start_index = i
                
                cumulative_sum += row['No of Boxes']
                
                for block in time_blocks:
                    if abs(cumulative_sum - block['No of Boxes']) <= 1:
                        df2_shift.loc[start_index:i, 'StartTime'] = block['StartTime']
                        df2_shift.loc[start_index:i, 'EndTime'] = block['EndTime']
                        cumulative_sum = 0
                        start_index = None
                        break
    

    processed_dfs.append(df2_shift)


final_df = pd.concat(processed_dfs, ignore_index=True)

final_df.sort_values(["Date","Shift","StartTime"],ascending=True,inplace=True)

final_df["Date"] = final_df["Date"].dt.date


final_df.to_excel(os.path.join(data_dir,"Test Data","dummy.xlsx"),index=False)
