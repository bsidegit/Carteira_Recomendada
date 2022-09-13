# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 08:08:21 2022

@author: eduardo.scheffer
"""

import pandas as pd
import pyodbc

def index_prices_database(benchmark_list):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bswm-db.database.windows.net'
    database = 'bswm'
    username = 'bswm-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
    
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
    
    
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    
    indices = pd.read_sql_query("SELECT * FROM Tbl_Indices", con=conn)
    indices_valores = pd.read_sql_query("SELECT * FROM Tbl_IndicesValores ORDER BY DtRef ASC", con=conn)
    #Indices = pd.read_sql_query("SELECT * FROM Tbl_Indices", con=conn)
    #IndicesValores = pd.read_sql_query("SELECT * FROM Tbl_IndicesValores", con=conn)
    
    benchmarksId = list(indices[(indices['NomeIndice'].isin(benchmark_list))]['IdIndice'])
    benchmark_list = list(indices[(indices['NomeIndice'].isin(benchmark_list))]['NomeIndice'])
    
    dates = list(pd.to_datetime(indices_valores['DtRef']).drop_duplicates())
    
    indices_select = pd.DataFrame(columns = benchmarksId, index = dates)
    indices_select.index.name = None
    
    for col in benchmarksId:
        valores = indices_valores[(indices_valores["IdIndice"] == col)][['DtRef', 'Valor']]
        set(valores['DtRef'].dt.date)
        valores.set_index('DtRef', inplace=True, drop=True)
        valores.columns = [col]
        indices_select.loc[list(valores.index), col] = valores[col]
     
    #cotas_select = cotas_select.iloc[:-1,:]

    indices_select.columns = benchmark_list
    indices_select = indices_select.iloc[len(indices_select)-15*252:len(indices_select)-3,:]
     
    return indices_select