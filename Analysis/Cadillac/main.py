import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pandas.api.types import is_numeric_dtype

data = pd.read_excel("dummy - Copy.xlsx")

data["Date Shift"] = data["Date"].astype(str)+" "+data["Shift"]

for j in [1]: 
    temp_df = data
    for i in temp_df.columns:
        if is_numeric_dtype(temp_df[i]):
            min_c = round(temp_df[i].min(),2)    
            max_c = round(temp_df[i].max(),2)
            quantile_1 = round(temp_df[i].quantile(.10),2)    
            quantile_2 = round(temp_df[i].quantile(.90),2)
            mean = round(temp_df[i].mean(),2)
            

            # Plotting the graph
            plt.figure(figsize=(14, 8))
            plt.plot(temp_df['Date Shift'], temp_df[i], marker='o', linestyle='-', color='b', label=i)

            # Adding horizontal lines for mean, min, and max
            plt.axhline(y=mean, color='g', linestyle='--', label=f'Mean: {mean}')
            plt.axhline(y=max_c, color='r', linestyle='--', label=f'Max: {max_c}')
            plt.axhline(y=min_c, color='r', linestyle='--', label=f'Min: {min_c}')
            plt.axhline(y=quantile_1, color='y', linestyle='--', label=f'10th percentile: {quantile_1}')
            plt.axhline(y=quantile_2, color='y', linestyle='--', label=f'90th percentile: {quantile_2}')
            # Set labels and title
            # Set labels and title
            plt.xlabel('Date and Shift',fontsize = 14)
            plt.ylabel(i,fontsize = 14)
            plt.title(f'{i} by Date and Shift', fontsize=18)

            # Rotate the x-axis labels for better readability
            plt.xticks(rotation=45, ha='right')

            # Adjust layout to avoid label cutoff
            plt.tight_layout()

            # Add legend
            plt.legend(fontsize=14, fancybox=True, framealpha=0.8, loc="upper left", bbox_to_anchor=(1, 1))

            # Show the plot
            plt.show()