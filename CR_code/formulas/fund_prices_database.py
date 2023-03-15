# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 21:59:37 2022

@author: eduardo.scheffer
"""

import pandas as pd
import pyodbc


def fund_prices_database(cnpj_list, date_first, date_last):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bside-srv.database.windows.net'
    database = 'bswm'
    username = 'bside-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
    
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
    
    
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    CadastroFundos = pd.read_sql_query("SELECT * FROM Tbl_CadastroFundos", con=conn)
    CotasPL = pd.read_sql_query("SELECT * FROM Tbl_CotasPL", con=conn)
    CotasPL = CotasPL.drop(columns = ['PL', 'Status'])
    
    dict_map = dict(zip(CadastroFundos['IdFundo'], CadastroFundos['CNPJ']))
    
    CotasPL = CotasPL.pivot(index = 'DtRef', columns ='IdFundo' , values = 'Cota')
    CotasPL = CotasPL[((CotasPL.index >= date_first) & (CotasPL.index <= date_last))]
    
    # Change IdFundo for CNPJ
    fund_list = list(CotasPL.columns)
    fund_list = [dict_map.get(item, item) for item in fund_list]
    CotasPL.columns = fund_list
    
    # Filter only funds needed
    CotasPL = CotasPL.loc[:,CotasPL.columns.isin(cnpj_list)]
    
    return CotasPL