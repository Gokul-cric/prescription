import json
import pandas as pd
import numpy as np
import os
from API.helper import findPCs,estimateModelParameters,additives_simulation,predict,additive_refinement
import json


def load_all_configs():
    """
    Function to load all config files from a directory based on a given prefix.
    
    :param prefix: (Optional) String to match the start of the filenames. 
                   If None, all files in the directory will be loaded.
    :return: Dictionary of configs with file names as keys.
    """
    configs = {}
    config_path = os.path.join(os.getcwd(), "Config")
    
    for file_name in os.listdir(config_path):
        if file_name.endswith('.json'):
            file_path = os.path.join(config_path, file_name)
            with open(file_path, "r") as config_file:
                config = json.load(config_file)
            configs[file_name] = config
    
    return configs

def read_data(foundry_name,group_name):
    """function to read data

    :param config: config file
    :type config:dict 
    :return: data
    :rtype: pd.dataframe
    """
    cd=os.getcwd()
    file_path = os.path.join(cd, 'Data')
    
    data_path = os.path.join(file_path,foundry_name)
    for file_name in os.listdir(data_path):
        if file_name.startswith(group_name) and file_name.endswith('.xlsx'):
            full_file_path = os.path.join(data_path, file_name)
            data = pd.read_excel(full_file_path) 

    return data

def without_bound(data, config):
    """
    Processes input data by filtering, renaming columns, and calculating mean values for specified columns.

    Parameters:
    -----------
    data : pandas.DataFrame
        The input dataframe to process.
    
    config : dict
        Configuration dictionary containing column names and settings.
    
    Returns:
    --------
    Dataframe containing properties and mean values
    """
    data = data.dropna(how="any")
    data_frac = np.arange(65, 101)
    data = rename_columns(data, config)
    data_len = len(data)

    result_dfs = [] 
    
    for idx, frac in enumerate(data_frac):
        data_lentrain = int(0.01 * data_len * frac)
        
        get_water_location = data.columns.get_loc('water_frac')
        input_data_block_total = data.iloc[:, :get_water_location + 1]
        
        Date_shift = input_data_block_total[config["date_shift"]]

        start_date, start_shift = Date_shift.iloc[0][config["date_shift"]]
        end_date, end_shift = Date_shift.iloc[data_lentrain - 1][config["date_shift"]]
    
        input_data_block_total = input_data_block_total[
            (input_data_block_total['date'] >= start_date) & 
            (input_data_block_total['date'] <= end_date)
        ]
        existing_columns = [col for col in config['columns_to_mean'] if col in input_data_block_total.columns]
        df_mean = input_data_block_total[existing_columns].mean()
        df_mean.T
        df_mean_df = pd.DataFrame(df_mean).reset_index()
        df_mean_df.columns = ['Column', 'Mean']
        df_transposed = df_mean_df.T
        df_transposed.columns = df_transposed.iloc[0]
        df_transposed = df_transposed[1:]
        df_transposed['Start Date'] = start_date
        df_transposed['End Date'] = end_date
        df_transposed['start_shift']=start_shift
        df_transposed['end_shift']=end_shift
    
        result_dfs.append(df_transposed)

    final_df = pd.concat(result_dfs, ignore_index=True)

    return final_df


def rename_columns(data, config):
    """
    Rename columns of a DataFrame based on a given configuration.

    Parameters:
    - data: pd.DataFrame, the DataFrame whose columns need to be renamed.
    - config: dict, configuration containing the 'columns_to_rename' mapping.

    Returns:
    - pd.DataFrame, the DataFrame with renamed columns.
    """
    column_to_rename = config.get("columns_to_rename", {})
    column_to_rename_ = {v: k for k, v in column_to_rename.items()}
    data=data.rename(columns=column_to_rename_)
    return data

def preprocessing(data,config):
    """
    This function performs the following operations:
    - Filter the data in the range of dates
    - Drops rows with any missing values.
    - Resets the index of the DataFrame.
    - Renames columns according to a mapping specified in the configuration.
    - Filters additives based on additional configuration settings.
    - Sets the index range of the data.
    - Identifies the location of the 'Water_frac' column.
    - Splits the data into input and output data blocks.
    - Renames columns of the output data block.
    - Extracts 'Date' and 'Shift' columns from the input data block.
    - Determines the list of shifted properties and additive columns.
    - Sorts and organizes input and output field lists.
    - Rearranges the input data block based on the computed column list.


    Args:
        data (pd.DataFrame): The input DataFrame containing raw data that needs preprocessing.
        config1 (dict): Configuration dictionary containing column renaming mappings and other settings.
        config2 (dict): Additional configuration dictionary for filtering additives.

    Returns:
        tuple: A tuple containing the following elements:
        - input_data_block_total (pd.DataFrame): The processed input data block.
        - output_data_block_total (pd.DataFrame): The processed output data block.
        - outputFieldsList (list): List of output field names.
        - inputFieldsList2 (list): List of input field names not in the shifted properties and additives list.
        - Date_shift (pd.DataFrame): DataFrame containing 'Date' and 'Shift' columns.
        - input_columnlist (list): List of all input columns.
        - shifted_prop_additivelist (list): List of shifted properties and additive columns.
        - data_len (int): Length of the processed input data block.

    """
    dates = config["date_range"]
    data['Date'] = pd.to_datetime(data['Date'], format='%d-%m-%Y')
    print(data.head())
    data= data[(data["Date"]>= dates[0]) & (data["Date"]<= dates[1])]
    print(data.head())
    print(data.shape)
    data = data.dropna(how="any")
    data=data.reset_index(drop=True)
    data=rename_columns(data, config)
    data = filter_additives(data,config)
    print(data.head())
    data_all = data.reset_index(drop=True)

    get_water_location = data_all.columns.get_loc('water_frac')

    input_data_block_total = data_all.iloc[:, :get_water_location+1]
    output_data_block_total = data_all.iloc[:, get_water_location+1:]
    output_data_block_total.columns = [col.split(".")[0] for col in output_data_block_total.columns]
    Date_shift = input_data_block_total[config["date_shift"]]

    shifted_prop_additivelist = list(output_data_block_total.columns) + config["ratio_list"]

    inputFieldsList = list(input_data_block_total.columns)
    inputFieldsList1 = list(shifted_prop_additivelist)
    inputFieldsList2 = list(set(inputFieldsList) - set(inputFieldsList1) - set(config["date_shift"]))
    outputFieldsList = list(output_data_block_total.columns)
    inputFieldsList1.sort()
    inputFieldsList2.sort()
    outputFieldsList.sort()

    input_columnlist=inputFieldsList1 +inputFieldsList2
    input_data_block_total = input_data_block_total[input_columnlist]

    return input_data_block_total, output_data_block_total, outputFieldsList, inputFieldsList2, Date_shift, input_columnlist,inputFieldsList1, len(input_data_block_total)

def filter_additives(data,config):
    """
    Filters and processes additive data based on provided configuration.

    Args:
    -----
    data : pd.DataFrame
        The input data containing additive fractions and other relevant columns.
    config : dict
        Configuration dictionary containing batch size and filter criteria.

    Returns:
        Filtered data based on the bound values in the config
    """
    additives_frac_list = config["additives_list"]
    batch_size = config["batch_size"]
    return_sand_frac = data['return sand_frac']
    additives_kg_columns = {items.split("_")[0] + "_kg": (data[items] * batch_size) / return_sand_frac for items in additives_frac_list}

    for col, values in additives_kg_columns.items():
        data[col] = values
    additives_filter = config["filter"][config["group_name"]]

    for items in additives_kg_columns:
        data=data[(data[items]>additives_filter[items][0]) & (data[items]<additives_filter[items][1])]
    data=data.drop(additives_kg_columns,axis=1)

    return data

def update_model_info_json(model_info, idx, all_model_info):
    """
    Updates model information in the provided dictionary with only the required components.

    Args:
        model_info (dict): The model information to be added.
        idx (int): An index or identifier for the model information.
        all_model_info (dict): The dictionary that stores all model information.

    Returns:
        dict: The updated model information as a dictionary.
    """


    model_coef_df = model_info.get('model_coef_df')
    input_data = model_info.get('input')
    output_data = model_info.get('output')

    all_model_info[str(idx)] ={"model_info":{
        'model_coef_df': model_coef_df,
            'input': input_data,
            'output': output_data,
            "model_info1" : model_info
        }
    }

    return all_model_info

def perform_modeling(data_all,config1,config2):
    """
    Performs modeling and prediction on the provided data using specified configurations.

    This function processes the input data to perform modeling and prediction, including:
    - Splitting the data into training and testing subsets.
    - Building and estimating models using principal component analysis (PCA).
    - Simulating additive properties and predicting properties.
    - Refining predictions and generating statistics.

    Args:
        data_all (pd.DataFrame): The complete dataset containing sand properties and additives information.
        config1 (dict): Configuration dictionary for prediction lists and other settings.
        config2 (dict): Configuration dictionary for group optimization, batch size, additives adjustments, and other details.

    Returns:
        tuple: A tuple containing the following elements:
            - df_test (pd.DataFrame): DataFrame containing statistical summaries (mean and standard deviation) of predicted additives.
            - predicted_dict (dict): Dictionary where keys are indices and values are DataFrames with predicted additives and properties.
            - adjusted_pred_dict (dict): Dictionary where keys are indices and values are DataFrames with adjusted predictions.
            - pred_prop_predicted_dict (dict): Dictionary where keys are indices and values are DataFrames with predicted properties.
            - adj_prop_predicted_dict (dict): Dictionary where keys are indices and values are DataFrames with adjusted predicted properties.
            - model_coeff_dict (dict): Dictionary where keys are indices and values are DataFrames with model coefficients.
            - input_data_dict (dict): Dictionary where keys are indices and values are DataFrames with input data blocks used for training.
    """
    input_data_block_total, output_data_block_total, outputFieldsList, inputFieldsList2, Date_shift, inputFieldsList, delayedoutputfieldlist, data_len = preprocessing(data_all, config2)
    
    predicted_dict = {}
    adjusted_pred_dict = {}
    pred_prop_predicted_dict = {}
    adj_prop_predicted_dict = {}
    input_data_dict = {}
    model_coeff_dict = {}
    all_model_info={}
    df_test = pd.DataFrame([])
    data_frac = np.arange(65,101) 
    all_model_info_new = {}
    for idx, frac in enumerate(data_frac):
        data_lentrain = int(0.01 * data_len * frac)
        input_data_block = input_data_block_total.iloc[:data_lentrain]
        output_data_block = output_data_block_total.iloc[:data_lentrain]
        print(Date_shift.head())
        start_date, start_shift = Date_shift.iloc[0][config2["date_shift"]]
        print(Date_shift.head())
        end_date, end_shift = Date_shift.iloc[data_lentrain - 1][config2["date_shift"]]
        input_data_dict[idx]=input_data_block
        input_data_blocktest = input_data_block_total.loc[input_data_block_total.index > 5].reset_index(drop=True)
        Date_shift_test = Date_shift.loc[Date_shift.index > 5].reset_index(drop=True)
        output_data_blocktest = output_data_block_total.loc[input_data_block_total.index > 5].reset_index(drop=True)

        retainedInputEigenVectors, retainedInputEigenValues = findPCs(input_data_block, inputFieldsList)
        model_info = estimateModelParameters(input_data_block, output_data_block, retainedInputEigenVectors, inputFieldsList, outputFieldsList, delayedoutputfieldlist,idx,config1)

        all_model_info_new[idx] = model_info
        #model_coeff_dict[idx] = model_info["model_coef_df"]
        model_info_=update_model_info_json(model_info,idx,all_model_info)
            
        optimum = pd.Series(config2['Group_opt'][config2["group_name"]])
        Mixture_capacity = config2["batch_size"]
        # input_data_blocktest.to_excel("train_data.xlsx")
        predicted_additive, predicted_additive_column = additives_simulation(input_data_blocktest, inputFieldsList2, model_info, optimum, Mixture_capacity)
        actual_predicted = pd.concat([Date_shift_test, predicted_additive], axis=1)
        predicted_dict[idx] = actual_predicted
        pred_list_cols = config2["pred_list"]
        actual_predicted = actual_predicted.assign(**{i.split("_")[0] + "_frac": actual_predicted[i] / actual_predicted[pred_list_cols].sum(axis=1) for i in pred_list_cols})
        property_list = list(config2["Group_opt"][config2["group_name"]].keys())
        temp = predict(actual_predicted, model_info)
        temp.rename(columns=dict(zip(property_list, [i + "_pred" for i in property_list])), inplace=True)
        pred_prop_predicted_dict[idx] = pd.concat([predicted_dict[idx], temp], axis=1)

        additve_to_adjust = config2["additive_to_adjust"][config2["group_name"]]
        adjustment_details = config2["Adjustment"][config2["group_name"]]
        refined_prediction = additive_refinement(actual_predicted, optimum, adjustment_details, additve_to_adjust)
        adjusted_pred_dict[idx] = refined_prediction

        refined_prediction = refined_prediction.assign(**{i.split("_")[0] + "_frac": refined_prediction[i] / refined_prediction[pred_list_cols].sum(axis=1) for i in pred_list_cols})

        temp = predict(refined_prediction, model_info)
        temp.rename(columns=dict(zip(property_list, [i + "_pred" for i in property_list])), inplace=True)
        adj_prop_predicted_dict[idx] = pd.concat([adjusted_pred_dict[idx], temp], axis=1)

        date_shift_ = pd.DataFrame([[start_date, start_shift, end_date, end_shift]], columns=['start_date', 'start_shift', 'end_date', 'end_shift'])
        
        required_data_mean = actual_predicted[predicted_additive_column].mean().to_frame().T
        col_mean = [item.split("_")[0] + "_mean" for item in inputFieldsList2]
        col_std = [item.split("_")[0] + "_std" for item in inputFieldsList2]
        required_data_std = actual_predicted[predicted_additive_column].std().to_frame().T
        required_data_mean.columns = col_mean
        required_data_std.columns = col_std
        required_data = pd.concat([required_data_mean, required_data_std], axis=1)
        required_data = pd.concat([date_shift_, required_data], axis=1)
        
        df_test = pd.concat([df_test, required_data], axis=0)
    
    return df_test, predicted_dict, adjusted_pred_dict, pred_prop_predicted_dict, adj_prop_predicted_dict, model_coeff_dict, input_data_dict,model_info_,all_model_info_new


def save_model_info_to_excel(model_info_dict, foundry_name, base_dir="Results"):
    """
    Save each key's record in `model_info_dict` to separate Excel files.
    Each nested dictionary (e.g., 'coeff', 'input_mean') inside 'model_info'
    will be stored in a separate sheet.

    Args:
        model_info_dict (dict): Dictionary containing model information.
        foundry_name (str): Name of the foundry, used to create a subdirectory.
        base_dir (str): Base directory where Excel files will be saved. Default is "Results".
    """
    output_dir = os.path.join(base_dir, foundry_name)
    os.makedirs(output_dir, exist_ok=True)
    for key, record in model_info_dict.items():

        excel_file_path = os.path.join(output_dir, f"model_info_{key}.xlsx")
        model_info = record.get("model_info", {})

        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            for sheet_name, data in model_info.items():

                if isinstance(data, dict):
                    df = pd.DataFrame([data])
                elif isinstance(data, list):
                    df = pd.DataFrame(data)
                else:
                    df = pd.DataFrame([data])

                df.to_excel(writer, sheet_name=sheet_name, index=False)

def convert_dict(df):
    """
    Converts a DataFrame into a dictionary with a nested structure.

    This function iterates over each row of the DataFrame and converts it into a dictionary. 
    Each record is stored in the resulting dictionary with the row index as the key and the 
    row data wrapped inside a 'without_bound' key.

    Parameters:
    -----------
    df : pd.DataFrame
        The DataFrame containing the data to be converted into a nested dictionary.
    
    Returns:
    --------
    dict
        A dictionary where the keys are string representations of the row indices, and the values 
        are dictionaries containing the row data wrapped inside a 'without_bound' key.
    """
    result = {}
    
    for idx, row in df.iterrows():
        record_key = str(idx)
        result[record_key] = {"without_bound": [row.to_dict()]}
    
    return result
                   

def save_dict_to_excel(data_dict, output_file, foundry_name):
    """
    Saves multiple dictionaries (from a structure like {0: {0: {key: value}}}) to an Excel file,
    with each dictionary stored in a separate sheet, and ensures the 'index' column is the first in the DataFrame.

    Parameters:
    -----------
    data_dict : dict
        A dictionary where keys are integers (e.g., 0, 1, 2) and values are nested dictionaries
        containing data that will be converted to DataFrames and stored in separate sheets.
    
    output_file : str
        The path of the Excel file where the DataFrames should be saved.
    
    foundry_name : str
        The name of the foundry, used to create a subdirectory for storing the results.

    Returns:
    --------
    None
        The function saves the DataFrames to the specified Excel file with multiple sheets.
    """
    file_path = os.path.join("Results", foundry_name, output_file)
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        for sheet_num, sheet_data in data_dict.items():
            df = pd.DataFrame.from_dict(sheet_data, orient='index')

            if 'index' in df.columns:
                cols = ['index'] + [col for col in df.columns if col != 'index']
                df = df[cols] 

            df.to_excel(writer, sheet_name=f"Sheet_{sheet_num}", index=False)

def dict_of_dfs_to_json(dict_of_dfs):
    """
    Converts a dictionary of DataFrames to JSON where each row is keyed by its index.
    Additionally, adds an 'index' column to each row in the DataFrame.

    Args:
        dict_of_dfs (dict): A dictionary where the keys are sheet names (str) and the values are DataFrames.

    Returns:
        dict_of_json (dict): The dictionary of JSON data.
    """

    dict_of_json = {}
    for key, df in dict_of_dfs.items():

        df = df.reset_index(drop=True)
        df['index'] = df.index 

        records = df.to_dict(orient='index')
        for index, record in records.items():
            for k, v in record.items():
                if isinstance(v, pd.Timestamp):
                    record[k] = v.strftime('%Y-%m-%d %H:%M:%S')

        dict_of_json[key] = records

    return dict_of_json


def calculate_correlations(adjusted_pred_dict, predicted_dict, pred_prop_predicted_dict, adj_prop_predicted_dict):
    """
    Calculates correlation matrices and percentage changes for numeric columns in multiple dictionaries of DataFrames.

    Args:
        adjusted_pred_dict (dict): A dictionary where the keys are identifiers (e.g., component names) and the values are DataFrames containing adjusted additive predictions.
        predicted_dict (dict): A dictionary where the keys are identifiers and the values are DataFrames containing additive predicted values.
        pred_prop_predicted_dict (dict): A dictionary where the keys are identifiers and the values are DataFrames containing predicted properties.
        adj_prop_predicted_dict (dict): A dictionary where the keys are identifiers and the values are DataFrames containing adjusted properties.

    Returns:
        tuple: A tuple containing five dictionaries:
            - correlation (dict): Correlation matrices for numeric columns in the adjusted prediction DataFrames.
            - pct_change_dict (dict): Percentage change for numeric columns in the predicted values DataFrames.
            - pred_correlation (dict): Correlation matrices for numeric columns in the predicted values DataFrames.
            - pred_prop_correlation (dict): Correlation matrices for numeric columns in the predicted properties DataFrames.
            - adj_prop_correlation (dict): Correlation matrices for numeric columns in the adjusted properties DataFrames.
    """
    correlation = {}
    pct_change_dict = {}
    pred_correlation = {}
    pred_prop_correlation = {}
    adj_prop_correlation = {}

    for i in adjusted_pred_dict:
        numeric_cols_adjusted = adjusted_pred_dict[i].select_dtypes(include=['number'])
        numeric_cols_predicted = predicted_dict[i].select_dtypes(include=['number'])
        numeric_cols_pred_prop = pred_prop_predicted_dict[i].select_dtypes(include=['number'])
        numeric_cols_adj_prop = adj_prop_predicted_dict[i].select_dtypes(include=['number'])

        correlation[i] = numeric_cols_adjusted.corr()
        pred_correlation[i] = numeric_cols_predicted.corr()
        pred_prop_correlation[i] = numeric_cols_pred_prop.corr()
        adj_prop_correlation[i] = numeric_cols_adj_prop.corr()

        pct_change_dict[i] = numeric_cols_predicted.pct_change() * 100

    return correlation, pct_change_dict, pred_correlation, pred_prop_correlation, adj_prop_correlation

def dict_conv(predicted_dict, adjusted_pred_dict,pct_change_dict,adj_prop_predicted_dict,input_data_dict,correlation,pred_corr):
  """Converts multiple dictionaries of dataframes into JSON format.

    Parameters:
    -----------
    df_test:dict
    predicted_dict : dict
        Dictionary containing predicted data.
    adjusted_pred_dict : dict
        Dictionary containing adjusted predicted data.
    adj_prop_predicted_dict : dict
        Dictionary containing adjusted predicted proportion data.
    input_data_dict : dict
        Dictionary containing input data.
    correlation : dict
        Dictionary containing correlation data.
    pred_corr : dict
        Dictionary containing predicted correlation data.

    Returns:
    --------
    Tuple
        A tuple containing JSON strings of the converted dictionaries.

    """
  pred_dict=dict_of_dfs_to_json(predicted_dict)
  pct_change= dict_of_dfs_to_json(pct_change_dict)
  adj_pred=dict_of_dfs_to_json(adjusted_pred_dict)
  adj_prop=dict_of_dfs_to_json(adj_prop_predicted_dict)
  input_dict=dict_of_dfs_to_json(input_data_dict)
  corr=dict_of_dfs_to_json(correlation)
  pred_corr=dict_of_dfs_to_json(pred_corr)

  return pred_dict,adj_pred,adj_prop,pct_change,input_dict,corr,pred_corr


def save_dict_to_excel_single_sheet(data_dict, output_file, data_key, foundry_name):
    """
    Saves data from multiple dictionaries (from a structure like {0: {key: value}}) to a single sheet in an Excel file.

    Parameters:
    -----------
    data_dict : dict
        A dictionary where keys are integers (e.g., 0, 1, 2) and values are dictionaries containing data 
        that will be converted to DataFrames and stored in a single sheet.
    
    output_file : str
        The path of the Excel file where the DataFrame should be saved.

    data_key : str
        The key within the nested dictionaries to extract the relevant data for saving.
    
    foundry_name : str
        The folder name where the results file should be saved.
    
    Returns:
    --------
    None
        The function saves the combined DataFrame to the specified Excel file.
    """
    file_path = os.path.join("Results", foundry_name, output_file)
    all_data = []
    
    for sheet_num, data in data_dict.items():
        df = pd.DataFrame(data[data_key])
        df['Record_ID'] = sheet_num  
        all_data.append(df)
    
    combined_df = pd.concat(all_data, ignore_index=True)
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, sheet_name="All_Data", index=False)

