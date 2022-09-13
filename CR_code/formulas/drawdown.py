# -*- coding: utf-8 -*-
"""
Created on Wed Jun 29 22:49:10 2022

@author: eduardo.scheffer
"""
import numpy as np
import pandas as pd
from tqdm import tqdm # progress bar

def event_drawdown(data_returns, min_period):
    
    # aux variables
    n_funds = data_returns.shape[1]
    n_dates = data_returns.shape[0]
    aux_df = pd.DataFrame(np.zeros((n_funds, 8)), index = list(data_returns.columns), 
                          columns = ['acc_returns', 'water_mark', 'max_drawdown', 'drawdown', 'minimum', 'time_back_maxDD', 'time_back', 'flag'])
    
    def update_stats(aux_df, d, min_period):
        for i in aux_df.index:
            if aux_df._get_value(i,'flag') == 0: # Only calculates statistics if it hasn't already recovered from max drawdown
                acc_return = aux_df._get_value(i,'acc_returns')
                water_mark = aux_df._get_value(i,'water_mark')
                time_back = aux_df._get_value(i,'time_back')
                time_back_maxDD = aux_df._get_value(i,'time_back_maxDD')
                max_drawdown = aux_df._get_value(i,'max_drawdown')
                drawdown = aux_df._get_value(i,'drawdown')
                
                if acc_return >= water_mark:
                    aux_df._set_value(i,'water_mark', acc_return)
                    aux_df._set_value(i,'minimum', acc_return)
                    if time_back > time_back_maxDD:
                       aux_df._set_value(i,'time_back_maxDD', time_back)
                    
                    aux_df._set_value(i,'time_back', 0)
                    if d > min_period:
                        aux_df._set_value(i,'flag', 1)
                        break
                else:
                    aux_df._set_value(i,'time_back', time_back + 1)
                    if acc_return < water_mark: 
                        aux_df._set_value(i,'minimum', acc_return)
                        
                    drawdown = acc_return - water_mark
                    aux_df._set_value(i,'drawdown', drawdown)
                    if drawdown < max_drawdown:
                        aux_df._set_value(i,'max_drawdown', drawdown)
        
        return aux_df
                             
    for d in tqdm(range(1,n_dates)): # dates loop
        aux_df['acc_returns'] = aux_df['acc_returns'] + data_returns.iloc[d,:]
        aux_df = update_stats(aux_df, d, min_period)
    
    '''
    aux_df['time_back_maxDD'] = aux_df['time_back_maxDD'].astype(int).astype(str)
    aux_df['time_back'] = aux_df['time_back'].astype(int).astype(str)
    
    for i in aux_df.index: # If fund is still hasn`t recovered from max drawdown
        max_drawdown = aux_df._get_value(i,'max_drawdown')
        drawdown = aux_df._get_value(i,'drawdown')
        time_back = str(int(aux_df._get_value(i,'time_back')))
        if drawdown < max_drawdown:
            aux_df._set_value(i,'time_back_maxDD', time_back+'*')
    '''        
    return aux_df['max_drawdown'], aux_df['time_back_maxDD']
    

def max_drawdown(data_returns):
    
    # aux variables
    n_funds = data_returns.shape[1]
    n_dates = data_returns.shape[0]
    aux_df = pd.DataFrame(np.zeros((n_funds, 7)), index = list(data_returns.columns), 
                          columns = ['acc_returns', 'water_mark', 'max_drawdown', 'drawdown', 'minimum', 'time_back_maxDD', 'time_back'])
    
    def update_stats(aux_df):
        for i in aux_df.index:
            acc_return = aux_df._get_value(i,'acc_returns')
            water_mark = aux_df._get_value(i,'water_mark')
            time_back = aux_df._get_value(i,'time_back')
            time_back_maxDD = aux_df._get_value(i,'time_back_maxDD')
            max_drawdown = aux_df._get_value(i,'max_drawdown')
            drawdown = aux_df._get_value(i,'drawdown')
            
            if acc_return >= water_mark:
                aux_df._set_value(i,'water_mark', acc_return)
                aux_df._set_value(i,'minimum', acc_return)
                if time_back > time_back_maxDD:
                   aux_df._set_value(i,'time_back_maxDD', time_back)
                aux_df._set_value(i,'time_back', 0)
                
            else:
                aux_df._set_value(i,'time_back', time_back + 1)
                if acc_return < water_mark: 
                    aux_df._set_value(i,'minimum', acc_return)
                   
                drawdown = acc_return - water_mark

                aux_df._set_value(i,'drawdown', drawdown)
                if drawdown < max_drawdown:
                    aux_df._set_value(i,'max_drawdown', drawdown)
        
        return aux_df
        
    for d in tqdm(range(1,n_dates)): # dates loop
        aux_df['acc_returns'] = aux_df['acc_returns'] + data_returns.iloc[d,:]
        aux_df = update_stats(aux_df)
            
    '''    
    aux_df['time_back_maxDD'] = aux_df['time_back_maxDD'].astype(int).astype(str)
    aux_df['time_back'] = aux_df['time_back'].astype(int).astype(str)
    
    for i in aux_df.index: # If fund is still hasn`t recovered from max drawdown
        max_drawdown = aux_df._get_value(i,'max_drawdown')
        drawdown = aux_df._get_value(i,'drawdown')
        time_back = aux_df._get_value(i,'time_back')
        if drawdown < max_drawdown:
            aux_df._set_value(i,'time_back_maxDD', time_back+'*')
    '''
        
    return aux_df['max_drawdown'], aux_df['time_back_maxDD']
