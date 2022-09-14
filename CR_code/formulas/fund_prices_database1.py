# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 21:59:37 2022

@author: eduardo.scheffer
"""

import pandas as pd
import numpy as np
import pyodbc
from tqdm import tqdm # progress bar
import datetime as dt


def fund_prices_database(cnpj_list):

    ''' 1) SET UP DATA BASE CONNECTOR ----------------------------------------------------------------------------'''

    server = 'bswm-db.database.windows.net'
    database = 'bswm'
    username = 'bswm-sa'
    password = 'BatataPalha123!'   
    driver= '{ODBC Driver 18 for SQL Server}'
    
    conn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+',1433;Database='+database+';Uid='+username+';Pwd='+ password)
    
    
    ''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    
    CadastroFundos = pd.read_sql_query("SELECT * FROM Tbl_CadastroFundos", con=conn)
    CotasPL = pd.read_sql_query("SELECT * FROM Tbl_CotasPL", con=conn)
    #Indices = pd.read_sql_query("SELECT * FROM Tbl_Indices", con=conn)
    #IndicesValores = pd.read_sql_query("SELECT * FROM Tbl_IndicesValores", con=conn)
    
    cnpj = list(CadastroFundos[(CadastroFundos['CNPJ'].isin(cnpj_list))]['CNPJ'])
    IdFundo = list(CadastroFundos[(CadastroFundos['CNPJ'].isin(cnpj_list))]['IdFundo'].apply(str))
    
    # get the dates of the fund with the longest historic:
    aux_df = CotasPL.groupby(by="IdFundo").count()
    longest = aux_df.index[aux_df["Cota"] == aux_df.max()[0]].tolist()[0]
    dates = pd.to_datetime(list(CotasPL[(CotasPL["IdFundo"] == longest)]["DtRef"]))
    
    cotas_select = pd.DataFrame(columns = IdFundo, index = dates)
    cotas_select.index.name = None
    fund_age = pd.DataFrame(data = np.zeros(len(cnpj)), columns = ['Age (months)'], index = IdFundo)    
    fund_startDt = pd.DataFrame(data = np.zeros(len(cnpj)), columns = ['Dates'], index = IdFundo)                       
    
    
    for col in tqdm(IdFundo):
        cotas = CotasPL[(CotasPL["IdFundo"] == int(col))][['DtRef', 'Cota']]
        set(cotas['DtRef'].dt.date)
        cotas.set_index('DtRef', inplace=True, drop=True)
        cotas.columns = [col]
        cotas_select.loc[list(cotas.index), col] = cotas[col]
        
        startDt = pd.to_datetime(list(CadastroFundos[(CadastroFundos["IdFundo"] == int(col))]["DtPrimeiraCota"]))
        fund_age.loc[col,'Age (months)'] = ((dt.date.today().year - startDt.year) * 12 + (dt.date.today().month - startDt.month))[0]
        fund_startDt.loc[col,'Dates'] = CadastroFundos[(CadastroFundos["IdFundo"] == int(col))]["DtPrimeiraCota"].iloc[0]
        
    cotas_select = cotas_select.iloc[:-1,:]
    cotas_select.columns = cnpj
    
    fund_age = fund_age.set_index([cnpj])
    fund_startDt = fund_startDt.set_index([cnpj])
    
    #for i in range(len(fund_startDt)):
     #   fund_startDt.iloc[i] = fund_startDt.iloc[i,0].date[0]
     

    cotas_select = cotas_select.iloc[98:,:] # Only dates after 2019-06-30
     
    
    return cotas_select, fund_age, fund_startDt