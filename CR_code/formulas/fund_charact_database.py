# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 18:46:34 2022

@author: eduardo.scheffer
"""

import pandas as pd
import numpy as np
import pyodbc
from tqdm import tqdm # progress bar
import datetime as dt


def fund_charact_database(cnpj_list):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bswm-db.database.windows.net'
    database = 'bswm'
    username = 'bswm-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
    
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
    
    
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    
    CadastroFundos = pd.read_sql_query("SELECT * FROM Tbl_CadastroFundos", con=conn)
    fund_caract = CadastroFundos[(CadastroFundos['CNPJ'].isin(cnpj_list))][['CNPJ', 'ConvResgate', 'LiqResgate', 'ClasseBSide', 'Estratégia', 'Geografia', 'Moeda']]
    fund_caract['Liquidez (D+)'] = fund_caract['ConvResgate']+fund_caract['LiqResgate']
    fund_caract = fund_caract.drop(['ConvResgate', 'LiqResgate'], axis = 1)
    
    return fund_caract