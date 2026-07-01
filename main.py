import os
from preprocess import load_all_configs,read_data,perform_modeling,calculate_correlations,without_bound,dict_conv,save_dict_to_excel,convert_dict,save_dict_to_excel_single_sheet,save_model_info_to_excel
import pandas as pd
import json

configs=load_all_configs()    
config1 = configs.get("config.json")
config2 = configs.get(config1["foundry_name"]+".json")

cwd=os.getcwd()
ressults_dir = os.path.join(cwd,"Model Results",config1["foundry_name"])

data_all=read_data(config1["foundry_name"],config2["group_name"])

data=data_all.copy() 
data_without_bound=without_bound(data,config2)
data_without_bound_=data_without_bound.to_dict(orient='index')


data_all_corr=data_all.select_dtypes(include=['number']).corr()
data_correlation=data_all_corr.to_dict()

df_test,predicted_dict, adjusted_pred_dict, pred_prop_predicted_dict, adj_prop_predicted_dict, model_coeff_dict, input_data_dict,model_info,all_model_info_new=perform_modeling(data_all,config1,config2)

correlation, pct_change_dict,pred_correlation, pred_prop_correlation,adj_prop_correlation=calculate_correlations(adjusted_pred_dict, predicted_dict, pred_prop_predicted_dict, adj_prop_predicted_dict)

predicted_properties,Adjusted_prediction,Adjusted_properties,Percentage_change,input_dict,Adjusted_correlation,Prediction_properties_correlation=dict_conv(predicted_dict, adjusted_pred_dict,pct_change_dict,adj_prop_predicted_dict,input_data_dict,correlation,pred_prop_correlation)
model_output_statistics=df_test.reset_index(drop=True).to_dict(orient='index')


data_correlation 
model_info
model_output_statistics
Adjusted_properties
input_dict
Adjusted_correlation
Prediction_properties_correlation
Percentage_change
data_without_bound_

#(pd.DataFrame(data_all_corr_).to_excel(os.path.join("Results",config1['foundry_name'],"data_all_corr.xlsx"))
#(pd.DataFrame(data_without_bound_).T).to_excel(os.path.join("Results",config1['foundry_name'],"without_bound_.xlsx"))
#save_dict_to_excel_single_sheet(data_without_bound_,"data_without_bound.xlsx",'without_bound',config1['foundry_name'])
#save_dict_to_excel(df_test_,"df_test.xlsx",'df_test',config1['foundry_name'])
#save_dict_to_excel(adj_pred, "Adj_predicted_dict_opt.xlsx",'adjusted_pred_dict',config1['foundry_name'])
#save_dict_to_excel(adj_prop, "Adj_prop_predicted_dict_opt.xlsx",'adj_prop_predicted_dict',config1['foundry_name'])
#save_dict_to_excel(input_dict,"Input_data_dict_opt.xlsx",'input_data_dict',config1['foundry_name'])
#save_dict_to_excel(corr,"correlation_dict_opt.xlsx",'correlation',config1['foundry_name'])
#save_dict_to_excel(pred_corr,"pred_correlation_dict_opt.xlsx",'pred_correlation',config1['foundry_name'])
#save_model_info_to_excel(model_info, config1['foundry_name'],base_dir="Results")


if config1["save_model"]:
    with open(os.path.join(ressults_dir,'output1_props.json'), 'w') as json_file:
        json.dump(model_info, json_file, indent=4) 

    with open(os.path.join(ressults_dir,'output2_props.json'), 'w') as json_file:
        json.dump(all_model_info_new, json_file, indent=4) 


