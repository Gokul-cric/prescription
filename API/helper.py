# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 11:43:15 2022

@author: ABhishankar Kumar
"""

import pandas as pd
import numpy as np
import copy
from cvxopt import matrix, solvers
import os
import json

def computeBounds(data):
    """
    Computes the minimum and maximum values for each column in a DataFrame.

    Parameters:
    -----------
    data : pd.DataFrame
        A DataFrame for which the minimum and maximum values are to be computed.

    Returns:
    --------
    pd.DataFrame
        A DataFrame with two columns:
        - 'min': Contains the minimum values for each column in the input DataFrame.
        - 'max': Contains the maximum values for each column in the input DataFrame.
    """
    dataMinSeries = data.min().to_frame("min")
    dataMaxSeries = data.max().to_frame("max")
    dataMinMaxDataFrame = pd.concat([dataMinSeries, dataMaxSeries], axis=1, join='inner')
    return dataMinMaxDataFrame


def NormalizeColumns(columnNameList, dataDF, dataMinMaxDataFrame):
    """
    Normalizes specified columns in a DataFrame based on provided minimum and maximum values.

    This function scales the specified columns of a DataFrame to a range of [0, 1] using the 
    minimum and maximum values provided in `dataMinMaxDataFrame`. If the difference between 
    the max and min values is zero, the column is set to 0.0 to avoid division by zero.
    Parameters:
    -----------
    columnNameList : list
        List of column names to be normalized.
    dataDF : pd.DataFrame
        DataFrame containing the data to be normalized.
    dataMinMaxDataFrame : pd.DataFrame
        DataFrame with columns 'min' and 'max' containing the minimum and maximum values 
        for each column in `columnNameList`.

    Returns:
    --------
    pd.DataFrame
        A new DataFrame with the specified columns normalized.
    """
    scaledData = dataDF.copy()
    for col in columnNameList:
        den = dataMinMaxDataFrame["max"][col] - dataMinMaxDataFrame["min"][col]
        if den == 0.0:
            scaledData[col] = 0.0
        else:
            scaledData[col] = (scaledData[col] - dataMinMaxDataFrame["min"][col]) / den
    return scaledData


def deNormalizeColumns(columnNameList, dataDF, dataMinMaxDataFrame):
    """
    Reverts the normalization of specified columns in a DataFrame to their original scale.

    This function denormalizes the specified columns of a DataFrame by using the minimum 
    and maximum values provided in `dataMinMaxDataFrame`. If the difference between the 
    max and min values is zero, the column is set to the max value.

    Parameters:
    -----------
    columnNameList : list
        List of column names to be denormalized.
    dataDF : pd.DataFrame
        DataFrame containing the data to be denormalized.
    dataMinMaxDataFrame : pd.DataFrame
        DataFrame with columns 'min' and 'max' containing the minimum and maximum values 
        for each column in `columnNameList`.

    Returns:
    --------
    pd.DataFrame
        A new DataFrame with the specified columns denormalized to their original scale.
    """
    unscaledData = dataDF.copy()
    for col in columnNameList:
        den = dataMinMaxDataFrame["max"][col] - dataMinMaxDataFrame["min"][col]
        if den == 0.0:
            unscaledData[col] = dataMinMaxDataFrame["max"][col]
        else:
            unscaledData[col] = unscaledData[col] * den + dataMinMaxDataFrame["min"][col]
    return unscaledData


def meanCenterData(dataDF):
    """
    This function subtracts the mean of each column from the corresponding column values 
    in the DataFrame, effectively centering the data around zero.

    Parameters:
    -----------
    dataDF : pd.DataFrame
        DataFrame containing the data to be mean-centered.

    Returns:
    --------
    tuple:
        - pd.DataFrame: A new DataFrame with the mean-centered data.
        - pd.Series: A Series containing the mean of each column from the original DataFrame.
    """
    DF = dataDF.copy()
    meanSeries = DF.mean()
    DF = DF - meanSeries
    return DF, meanSeries


def findPCs(input_data_block, input_columnlist):
    """
    Computes the Principal Components (PCs) for a given dataset.

    This function normalizes the data, mean-centers it, and then computes the 
    eigenvalues and eigenvectors using Singular Value Decomposition (SVD) to identify 
    the Principal Components that capture the desired variance (default 95%).

    Parameters:
    -----------
    input_data_block : pd.DataFrame
        DataFrame containing the data for which PCs are to be computed.
    input_columnlist : list
        List of columns to be used for the PC computation.

    Returns:
    --------
    tuple:
        - np.ndarray: Retained eigenvectors corresponding to the selected Principal Components.
        - np.ndarray: Retained eigenvalues corresponding to the selected Principal Components.
    """
    dataInputMinMaxDataFrame = computeBounds(input_data_block)
    scaledData = NormalizeColumns(input_columnlist, input_data_block, dataInputMinMaxDataFrame)
    datainput, inputMean = meanCenterData(scaledData)
    A = datainput[input_columnlist].values
    _, s, vh = np.linalg.svd(A)
    eVectors = vh.T
    eValues = s**2 / (A.shape[0] - 1.0)
    variancePercent = eValues / sum(eValues)
    varianceCaptured = 0.0
    percentageVarianceExplainedPC = 0.95
    for i in range(len(eValues)):
        varianceCaptured += variancePercent[i]
        if varianceCaptured >= percentageVarianceExplainedPC:
            break
    retainedInputEigenVectors = eVectors[:, 0:(i+1)]
    retainedInputEigenValues = eValues[0:(i+1)]
    return retainedInputEigenVectors, retainedInputEigenValues


def findPCs2(input_data_block, input_columnlist, dataInputMinMaxDataFrame):
    """
    Computes the Principal Components (PCs) for a given dataset using a precomputed 
    min-max DataFrame for normalization.

    This function normalizes the data using a precomputed min-max DataFrame, 
    mean-centers it, and then computes the eigenvalues and eigenvectors using Singular 
    Value Decomposition (SVD) to identify the Principal Components that capture the 
    desired variance (default 95%).

    Parameters:
    -----------
    input_data_block : pd.DataFrame
        DataFrame containing the data for which PCs are to be computed.
    input_columnlist : list
        List of columns to be used for the PC computation.
    dataInputMinMaxDataFrame : pd.DataFrame
        DataFrame with columns 'min' and 'max' containing the minimum and maximum values 
        for each column in `input_columnlist`.

    Returns:
    --------
    tuple:
        - np.ndarray: Retained eigenvectors corresponding to the selected Principal Components.
        - np.ndarray: Retained eigenvalues corresponding to the selected Principal Components.
    """
    scaledData = NormalizeColumns(input_columnlist, input_data_block, dataInputMinMaxDataFrame)
    datainput, inputMean = meanCenterData(scaledData)
    A = datainput[input_columnlist].values
    _, s, vh = np.linalg.svd(A)
    eVectors = vh.T
    eValues = s**2 / (A.shape[0] - 1.0)
    variancePercent = eValues / sum(eValues)
    varianceCaptured = 0.0
    percentageVarianceExplainedPC = 0.95
    for i in range(len(eValues)):
        varianceCaptured += variancePercent[i]
        if varianceCaptured >= percentageVarianceExplainedPC:
            break
    retainedInputEigenVectors = eVectors[:, 0:(i+1)]
    retainedInputEigenValues = eValues[0:(i+1)]
    return retainedInputEigenVectors, retainedInputEigenValues

def projection2ReducedSpace(X, retainedInputEigenVectors):
    """
    Projects data into a reduced-dimensional space using retained Principal Components.

    This function projects the original data matrix `X` into a lower-dimensional space 
    by multiplying it with the retained eigenvectors (Principal Components), resulting 
    in the reduced-space representation of the data.

    Parameters:
    -----------
    X : np.ndarray
        The original data matrix to be projected into a reduced-dimensional space.
    retainedInputEigenVectors : np.ndarray
        The eigenvectors corresponding to the selected Principal Components.

    Returns:
    --------
    np.ndarray
        The data projected into the reduced-dimensional space.
    """
    return np.dot(X, retainedInputEigenVectors)


def estimateModelParameters(input_data_block, output_data_block, retainedInputEigenVectors, inputFieldsList, outputFieldsList, delayedoutputfieldlist,idx,config):
    """
    Estimates model parameters using Principal Component Regression (PCR) and stores the output as an Excel file.

    Parameters:
    -----------
    input_data_block : pd.DataFrame
    output_data_block : pd.DataFrame
    retainedInputEigenVectors : np.ndarray
    inputFieldsList : list
    outputFieldsList : list
    delayedoutputfieldlist : list
    output_filename : str, optional
        Filename for the Excel output (default is "model_output.xlsx").

    Returns:
    --------
    dict
        Dictionary containing model coefficients, means, min-max values, and other model information.
    """
    model_info = {}
    

    dataInputMinMaxDataFrame = computeBounds(input_data_block)
    dataOutputMinMaxDataFrame = computeBounds(output_data_block)
    
    scaledInputData = NormalizeColumns(inputFieldsList, input_data_block, dataInputMinMaxDataFrame)
    scaledOutputData = NormalizeColumns(outputFieldsList, output_data_block, dataOutputMinMaxDataFrame)
    
    datainput, inputMean = meanCenterData(scaledInputData)
    dataoutput, outputMean = meanCenterData(scaledOutputData)
    

    inputScores = projection2ReducedSpace(datainput[inputFieldsList].values, retainedInputEigenVectors)
    output = dataoutput[outputFieldsList].values
    
    m1 = np.dot(inputScores.T, inputScores)
    m1 = np.linalg.pinv(m1)
    m2 = np.dot(m1, inputScores.T)
    modelCoefficientsInReducedSpace = np.dot(m2, output)
    modelCoefficientsInOriginalSpace = np.dot(retainedInputEigenVectors, modelCoefficientsInReducedSpace)

    modelCoefficientsInOriginalSpace_df = pd.DataFrame(modelCoefficientsInOriginalSpace, columns=outputFieldsList, index=inputFieldsList)
    modelCoefficientsInOriginalSpace_list = modelCoefficientsInOriginalSpace.tolist()
    model_coefficients_dict = modelCoefficientsInOriginalSpace_df.to_dict()
    inputMean = inputMean.to_dict()
    outputMean = outputMean.to_dict()
    model_info['coeff'] = modelCoefficientsInOriginalSpace_list
    model_info['inputMean'] = inputMean
    model_info['outputMean'] = outputMean
    model_info['input_features'] = inputFieldsList
    model_info['output_cols'] = outputFieldsList
    model_info['input_min_max'] = dataInputMinMaxDataFrame.to_dict()
    model_info['output_min_max'] = dataOutputMinMaxDataFrame.to_dict()
    model_info['delayed_output_list'] = delayedoutputfieldlist
    model_info["model_coef_df"] = model_coefficients_dict 
    input_properties = {}
    input_properties = {}
    for feature in inputFieldsList:
        if feature in dataInputMinMaxDataFrame.index:
            min_val = dataInputMinMaxDataFrame.loc[feature, 'min']
            max_val = dataInputMinMaxDataFrame.loc[feature, 'max']
            mean_val = inputMean.get(feature, None)
            input_properties[feature] = {'min': min_val, 'max':max_val,'mean':mean_val}

    output_properties = {}
    for feature in outputFieldsList:
        if feature in dataOutputMinMaxDataFrame.index:
            min_val = dataOutputMinMaxDataFrame.loc[feature, 'min']
            max_val = dataOutputMinMaxDataFrame.loc[feature, 'max']
            mean_val = outputMean.get(feature, None)
            output_properties[feature] = {'min': min_val, 'max':max_val,'mean':mean_val}
    model_info['input'] = input_properties
    model_info['output'] = output_properties
    
    return model_info
  
def predict(datainput,model_info):
    """
    Generates predictions using the provided model information.

    Parameters:
    -----------
    datainput : pd.DataFrame
        Input data for which predictions are to be made.
    model_info : dict
        Dictionary containing model coefficients, input features, means, and other relevant information.

    Returns:
    --------
    pd.DataFrame
        Predicted values for the output columns.
    """
    print(model_info['input_features'])
    print(datainput.head())
    data = copy.deepcopy(datainput[model_info['input_features']])
    data = data.reindex(model_info['input_features'],axis=1)
    data = NormalizeColumns(model_info['input_features'],data,model_info['input_min_max'])
    
    data = data -pd.Series(model_info['inputMean'])[model_info['input_features']]
    prediction = np.dot(data.values,model_info['coeff'])
    prediction = pd.DataFrame(prediction,columns=model_info['output_cols'])
    prediction = prediction + pd.Series(model_info['outputMean'])[model_info['output_cols']]
    prediction = deNormalizeColumns( model_info['output_cols'], prediction,model_info['output_min_max'])
    return prediction


def  optimize(optX,prevX,model_info,inputFieldsList2):
    """
    Performs optimization based on the model coefficients and input data.

    Parameters:
    -----------
    optX : pd.Series
        The target output values to be optimized.
    prevX : pd.Series
        The previous input values for comparison.
    model_info : dict
        Dictionary containing model coefficients, input features, means, and other relevant information.
    inputFieldsList2 : list
        List of input fields used for optimization.

    Returns:
    --------
    pd.Series
        Optimized input values.
    np.ndarray
        Beta coefficients from the model.
    """
    # modified on 20.09.2016
    optX1 = optX.reindex(model_info['output_cols'],axis = 0)
    prevX1 = prevX.reindex(model_info['delayed_output_list'],axis = 0)  
    optX1 = NormalizeColumns(model_info['output_cols'],optX1,model_info['output_min_max'])
    prevX1 = NormalizeColumns(model_info['delayed_output_list'],prevX1,model_info['input_min_max'])
    #optX1 = optX1 - model_info['outputMean'][model_info['output_cols']]
    #prevX1 = prevX1 -model_info['inputMean'][model_info['delayed_output_list']] 


    # Adjusting optX1 using model_info['outputMean']
    optX1 = optX1 - pd.Series({col: model_info['outputMean'].get(col, 0) for col in model_info['output_cols']})

    # Adjusting prevX1 using model_info['inputMean']
    prevX1 = prevX1 - pd.Series({col: model_info['inputMean'].get(col, 0) for col in model_info['delayed_output_list']})
        
    optX1 = optX1.values
    prevX1 = prevX1.values     
    modelMatrix = np.array(model_info['coeff']).T
    nSandProp = len(model_info['delayed_output_list'])
    nAll = len(model_info['input_features'])
    # nAdditives = len(inputFieldsList2)
    nAdditives =  nAll-nSandProp
    alpha = modelMatrix[:,:nSandProp]
    beta = modelMatrix[:,nSandProp:nAll]        
    # creating the weighting matrix
    # modified on 20.09.2016
    mSandProp = len(model_info['output_cols'])
    W = np.identity(mSandProp)
    P = 2.0 * np.dot(np.dot(beta.T,W),beta)    
    q1temp = np.dot(prevX1,alpha.T)
    q1 = np.dot(np.dot(q1temp,W),beta)
    q2 = np.dot(np.dot(optX1,W),beta)
    q = 2.0 * (q1 -q2)
    G = np.vstack((-np.identity(nAdditives),np.identity(nAdditives)))
   
    h1 = np.zeros((nAdditives)) -pd.Series(model_info['inputMean'])[inputFieldsList2].reindex(inputFieldsList2, axis = 0)
    h2 = np.ones((nAdditives)) - pd.Series(model_info['inputMean'])[inputFieldsList2].reindex(inputFieldsList2, axis = 0)
    h = np.hstack((-h1.T,h2.T))

    #delta =(model_info['input_min_max']["max"]) -(model_info['input_min_max']["min"])
    max_values = pd.Series(model_info['input_min_max']['max'])
    min_values = pd.Series(model_info['input_min_max']['min'])

    delta = max_values - min_values

    A = delta[inputFieldsList2].reindex(inputFieldsList2,axis = 0).values
    b = 1 - sum(pd.Series(model_info['inputMean'])[inputFieldsList2] * A)-sum(pd.Series(model_info['input_min_max']["min"])[inputFieldsList2])       
    P_matrix = matrix(P,tc='d')
    q_matrix = matrix(q,tc='d')
    G_matrix = matrix(G,tc='d')
    h_matrix = matrix(h,tc='d')
    A_matrix = matrix(A,tc ='d')
    A_matrix = A_matrix.T
    b_matrix = matrix(b,tc='d')
    solvers.options["show_progress"] = True
    solvers.options["maxiters"] = 2000

    result = solvers.qp(P_matrix,q_matrix,G_matrix,h_matrix, A_matrix,b_matrix )
    rs = pd.Series(result['x'],index=inputFieldsList2)
    rs = rs + pd.Series(model_info['inputMean'])[inputFieldsList2]
    rs = deNormalizeColumns(inputFieldsList2, rs, pd.DataFrame(model_info['input_min_max']).loc[inputFieldsList2])
    return rs,beta


def additives_simulation(test_data,manipulative_var_list,model_info,optimum,Mixture_capacity):
    """
    It predict additives dosage based on previous shift input data
    
    params:
        test_data: Test data 
        dtype: DataFrame
        
        manipulative_var: Manipulative variables to achieve desired target
        dtype: list
        
        model_info: Model information 
        dtype: Dictionary
        
        optimum: optimal target
        dtype: Series
    """
    if not issubclass(type(test_data), pd.DataFrame):
        raise TypeError("test_data should be Pandas DataFrame")
    if not issubclass(type(manipulative_var_list),list):
        raise TypeError("manipulative_var_list should be a list")
    
    if not issubclass(type(model_info), dict):
        raise TypeError("model_info should be dictionary")
    
    if not issubclass(type(optimum), pd.Series):
        raise TypeError("model_info should be Series")
    
    
    additives=np.zeros((len(test_data),len(manipulative_var_list)))
    prev_input=np.zeros((len(test_data),len(model_info['delayed_output_list'])))
    delayedoutputfieldlist=model_info['delayed_output_list']
    print(delayedoutputfieldlist)
    outputFieldsList=model_info['output_cols']
    for i in range(len(test_data)):
        
        pk_1=test_data[delayedoutputfieldlist].iloc[i]
        prevX = pd.Series(pk_1,index=delayedoutputfieldlist)
        
        
    # with new rejecttion
        optX=optimum.reindex(index=outputFieldsList)
 
        rs,beta=optimize(optX,prevX,model_info,manipulative_var_list)
        additives[i]=rs
        prev_input[i]=prevX
        
    prev_input=pd.DataFrame(prev_input,columns=delayedoutputfieldlist)
    # prev_input=prev_input[outputFieldsList]
    predicted_additive=pd.DataFrame(additives,columns=manipulative_var_list)
    predicted_additive=np.matrix(pd.DataFrame(additives,columns=manipulative_var_list))
    
    addtive_prediclist =[item.split("_")[0]+"_"+"predicted" for item in manipulative_var_list]
    predicted_additive=pd.DataFrame(additives,columns=addtive_prediclist)
    predicted_additive.index=range(len(predicted_additive))

#Additives in kg for group all
    predicted_additive_column=[]
    for i in range(len(addtive_prediclist)):
        modify_str_all=addtive_prediclist[i]+str('_kg')
        predicted_additive_column.append(modify_str_all)
        predicted_additive.insert(1,modify_str_all,(predicted_additive[addtive_prediclist[i]]*Mixture_capacity/predicted_additive['return sand_predicted']))    
    
    predicted_additive=predicted_additive[predicted_additive_column]
    predicted_additives=pd.concat([predicted_additive,prev_input],axis=1)

   
    return predicted_additives,predicted_additive_column


def additive_refinement(df,optimum,adjustment_dict,additve_to_adjust):
    """
    Parameters
    ----------
    df : Dataframe
        dataframe contains optimized value in kg
    optimum : Series
        optimal value corrosponding to sand parameters
    adjustment_dict : Dictionary
        It contains refinment coeff,weights, offset,min and max bound as a key of dictionary
    additve_to_adjust : list
        List of additives which we want to adjust
        
    Returns
    -----
    copy_actual_predicted : DataFrame
      dataframe with adjusted value of additives

    """
    if not issubclass(type(df), pd.DataFrame):
        raise TypeError("df should be Pandas DataFrame")
    
    if not issubclass(type(optimum), pd.Series):
        raise TypeError("optimum should be Pandas series")
    
    if not issubclass(type(adjustment_dict), dict):
        raise TypeError("adjustment_dict should be dictionary")
    
    if not issubclass(type(additve_to_adjust), list):
        raise TypeError("adjustment_dict should be a list")
    
    expected_refinement_key=['param','ref_coeff','weight','prediction_bound','offset']
    
    copy_actual_predicted=copy.deepcopy(df)
     
    for item in additve_to_adjust:
        refinement_data=adjustment_dict[item]
        
        if len(set(expected_refinement_key)-set(list(refinement_data.keys()))) !=0:
            raise Exception("Refinement keys are missing")
            
        param_handle=refinement_data["param"]
        
        ref_coeff,weights=refinement_data['ref_coeff'],refinement_data['weight']
        
        if len(param_handle)!= len(ref_coeff):
            raise Exception("Length of adjustment handle/param and refinment coefficient should be same")
        if len(param_handle)!= len(weights):
            raise Exception("Length of adjustment handle/param and weights should be same")
        #if sum(weights)!=1:
            #raise Exception("Sum of weights should be equal to 1")
        
        
        for i, idx in enumerate(param_handle):
            refinement_on_additives=(refinement_data["weight"][i]*(optimum[idx]-copy_actual_predicted[idx]))/refinement_data["ref_coeff"][i]
            col_name=item+"_"+"predicted"+"_"+"kg"
            copy_actual_predicted[col_name]=copy_actual_predicted[col_name]+refinement_on_additives
        copy_actual_predicted[col_name]=copy_actual_predicted[col_name]+refinement_data["offset"]
        copy_actual_predicted[col_name] = np.where(copy_actual_predicted[col_name] > adjustment_dict[item]['prediction_bound']["max"], adjustment_dict[item]['prediction_bound']["max"], copy_actual_predicted[col_name])
        copy_actual_predicted[col_name] = np.where(copy_actual_predicted[col_name] < adjustment_dict[item]['prediction_bound']["min"], adjustment_dict[item]['prediction_bound']["min"], copy_actual_predicted[col_name])
    return copy_actual_predicted