# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 08:08:21 2022

@author: eduardo.scheffer
"""

import pandas as pd
import pyodbc

def stock_prices_database(stocks_list, date_first, date_last):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bside-srv.database.windows.net'
    database = 'bswm'
    username = 'bside-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
     
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
     
     
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
     
    PrecoAcoes = pd.read_sql_query("SELECT * FROM Tbl_PrecosAcoes", con=conn)
    PrecoAcoes = PrecoAcoes.drop(columns = ['Moeda'])
     
    PrecoAcoes = PrecoAcoes.pivot(index = 'DtRef', columns ='Ticker' , values = 'Preco')
    PrecoAcoes = PrecoAcoes[((PrecoAcoes.index >= date_first) & (PrecoAcoes.index <= date_last))]
     
    # Filter only funds needed
    PrecoAcoes = PrecoAcoes.loc[:,PrecoAcoes.columns.isin(stocks_list)]
     
    return PrecoAcoes