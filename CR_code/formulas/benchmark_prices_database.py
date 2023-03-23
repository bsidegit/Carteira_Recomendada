# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 08:08:21 2022

@author: eduardo.scheffer
"""

import pandas as pd
import pyodbc
import datetime as dt

def benchmark_prices_database(benchmark_list, date_first, date_last):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bside-srv.database.windows.net'
    database = 'bswm'
    username = 'bside-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
    
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
    
    
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    
    indices = pd.read_sql_query("SELECT * FROM Tbl_Indices", con=conn)
        
    index_str = ''
    for i in list(indices.index):
        if indices.loc[i, 'NomeIndice'] in benchmark_list:
            index_str = index_str + 'IdIndice =' + str(indices.loc[i, 'IdIndice']) + ' OR '
            
    index_str = index_str[:-3]
        
    str_query = "SELECT * FROM Tbl_IndicesValores WHERE ("+index_str+") ORDER BY DtRef ASC"

    values_index = pd.read_sql_query(str_query, con=conn)

    values_index = values_index.pivot(index = 'DtRef', columns ='IdIndice' , values = 'Valor')
    
    #-30 to get the IPCA since the first date (monthly data) - these dates will be deleted after obtaining daily IPCA rates
    date_first = date_first-dt.timedelta(days=32)
    values_index = values_index[((values_index.index >= date_first) & (values_index.index <= date_last))]
    
    # Change IdIndice for index name
    dict_map = dict(zip(indices['IdIndice'], indices['NomeIndice']))
    benchmark_names = list(values_index.columns)
    benchmark_names = [dict_map.get(item, item) for item in benchmark_names]
    values_index.columns = benchmark_names
    
    return values_index