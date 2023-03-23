# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 08:08:21 2022

@author: eduardo.scheffer
"""

import pandas as pd
import pyodbc

def fixed_income_prices_database(fixedIncome_list, date_first, date_last):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bside-srv.database.windows.net'
    database = 'bswm'
    username = 'bside-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
     
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
     
     
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
     
    fixedIncome = pd.read_sql_query("SELECT * FROM Tbl_CadastroRendaFixa", con=conn)
    
    fixedIncome_str = ''
    for i in list(fixedIncome.index):
        if fixedIncome.loc[i, 'ISIN'] in fixedIncome_list:
            fixedIncome_str = fixedIncome_str + 'IdAtivo =' + str(fixedIncome.loc[i, 'IdAtivo']) + ' OR '
            
    fixedIncome_str = fixedIncome_str[:-3]
        
    str_query = "SELECT * FROM Tbl_PrecosRF WHERE ("+fixedIncome_str+") ORDER BY DtRef ASC"

    values_fixedIncome = pd.read_sql_query(str_query, con=conn)

    values_fixedIncome = values_fixedIncome.pivot(index = 'DtRef', columns ='IdAtivo' , values = 'Preco')
    
    values_fixedIncome = values_fixedIncome[((values_fixedIncome.index >= date_first) & (values_fixedIncome.index <= date_last))]
    
    # Change IdIndice for index name
    dict_map = dict(zip(fixedIncome['IdAtivo'], fixedIncome['ISIN']))
    fixedIncome_ISIN = list(values_fixedIncome.columns)
    fixedIncome_ISIN = [dict_map.get(item, item) for item in fixedIncome_ISIN]
    values_fixedIncome.columns = fixedIncome_ISIN
    
    return values_fixedIncome