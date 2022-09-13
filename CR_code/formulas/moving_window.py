# -*- coding: utf-8 -*-
"""
Created on Mon Jul  4 08:24:49 2022

@author: eduardo.scheffer
"""

import numpy as np
import pandas as pd
import math as math
from tqdm import tqdm # progress bar
from scipy.stats import percentileofscore, hmean

def moving_window(measurement, data_returns, step, window, fund_late_start):

    n_funds = data_returns.shape[1]
    n_dates = data_returns.shape[0]
    n_windows = math.floor((n_dates - window)/step)+1
    
    output_df = pd.DataFrame(np.zeros((n_funds, 3)), index = list(data_returns.columns), columns = ['median', 'min', 'max'])   
    
    aux_df = pd.DataFrame(np.zeros((n_windows, n_funds)), columns = data_returns.columns)
    
    i = 0
    
    if measurement == 'return':
        
        
        for i in tqdm(range(n_windows)):
            start = i*step
            end = start + window
            aux_df.iloc[i,:] = data_returns.iloc[start:end, :].sum()
            
        for fund in list(fund_late_start.index):         
            aux_df.loc[0:fund_late_start.loc[fund,'index'],fund] = aux_df.loc[fund_late_start.loc[fund,'index']:,fund].median()
            
        output_df['median'] = aux_df.median()
        output_df['min'] = aux_df.min()
        output_df['max'] = aux_df.max()
        
        #Delete indicators that have insuficient track record
        output_df['median'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['median'])
        output_df['min'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['min'])
        output_df['max'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['max'])
        
        return output_df['median'], output_df['min'], output_df['max']
        
    elif measurement == 'sharpe':
        
        for i in tqdm(range(n_windows)):
            start = i*step
            end = start + window
            returns = data_returns.iloc[start:end, :].sum()
            volatility = data_returns.iloc[start:end, :].std()*np.sqrt(252)
            aux_df.iloc[i,:] = (returns**(252/window) - returns.iloc[0]**(252/window))/volatility
                       
        for fund in list(fund_late_start.index):         
            aux_df.loc[0:fund_late_start.loc[fund,'index'],fund] = aux_df.loc[fund_late_start.loc[fund,'index']:,fund].median()
           
        output_df['median'] = aux_df.median()
        output_df['min'] = aux_df.min()
        output_df['max'] = aux_df.max()
        
        #Delete indicators that have insuficient track record
        output_df['median'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['median'])
        output_df['min'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['min'])
        output_df['max'] = np.where(data_returns.iloc[0:3,:].isna().sum()>=2, 0, output_df['max'])
    
        return output_df['median'], output_df['min'], output_df['max']
    
    elif measurement == 'consistency':
        
        # Use only funds with sufficient historic 
        data_returns2 = data_returns[data_returns.columns[~data_returns.columns.isin(list(fund_late_start.index))]] 
        output_df2 = pd.DataFrame(np.zeros((data_returns2.shape[1], 1)), index = list(data_returns2.columns), columns = ['hmean'])   
        aux_df2 = pd.DataFrame(np.zeros((n_windows, len(data_returns2.columns))), columns = data_returns2.columns)
               
        for i in tqdm(range(n_windows)):
            start = i*step
            end = start + window
            returns = data_returns2.iloc[start:end, :].sum()
            returns_sorted = returns.sort_values(ascending=False)
            returns_perc = returns.apply(lambda x: percentileofscore(returns_sorted, x))
            
            aux_df2.iloc[i,:] = returns_perc/100
                           
        output_df2['hmean'] = hmean(aux_df2) 
        
        # Now normalize the harmonic mean using percentiles
        output_df2_sorted = output_df2['hmean'].sort_values(ascending=False)
        output_df2_percent = output_df2['hmean'].apply(lambda x: percentileofscore(output_df2_sorted, x))
        output_df2['hmean'] = output_df2_percent/100
        
        for col in list(aux_df2.columns):
            output_df.loc[col, 'median'] = output_df2.loc[col, 'hmean'] 
            
        for col in list(aux_df2.columns):
            aux_df[col] = aux_df2[col]
            
        return output_df['median'], aux_df2