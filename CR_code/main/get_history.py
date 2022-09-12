# -*- coding: utf-8 -*-
"""
Created on Tue May 31 19:16:47 2022
@author: eduardo.scheffer

Use: get daily return of a sample of selected funds
"""

''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import numpy as np
import pandas as pd

import datetime as dt
from tqdm import tqdm # progress bar

import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

import matplotlib.pyplot as plt
import seaborn as sns

import warnings 
warnings.filterwarnings('ignore')


''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 

path_read = 'I:/GESTAO/5) Produtos/5.2) Fundos Abertos/Análises/5.2.1) Comps/Comparables.xlsx'
cnpj_list = pd.read_excel(path_read, sheet_name = "MM universe").dropna(how='all',axis='columns').iloc[1:,1:].to_numpy().flatten()
cnpj_list = list(cnpj_list)

cnpj_list = [str(x) for x in cnpj_list] 
for i in range(len(cnpj_list)):
    if len(cnpj_list[i]) == 13:
        cnpj_list[i] = '0'+ cnpj_list[i]


''' 2) IMPORT HISTORICAL DATA ---------------------------------------------------------------------------------''' 

period = 252*2.5 # lenth of past time series to import

date_last = dt.datetime.today() - dt.timedelta(days=5)
year_last = date_last.year 
month_last = date_last.month 

date_first = dt.datetime.today() - dt.timedelta(days=5+period)
year_first = date_first.year 
month_first = date_first.month 

fund_shares = {} # dictionary of dataframes (one for each fund)

readCVM = pd.DataFrame()
for p in range (year_last - year_first):
    x = str(year)
    zip_file.open
    url = 'http://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/DADOS/HIST/inf_diario_fi_{}.zip'.format(x)
    aux = pd.read_csv(url, compression = 'zip', sep=";", usecols = ["CNPJ_FUNDO", "DT_COMPTC", "VL_QUOTA"])
    readCVM.append(aux)    
    
    data = yyyy+mm # Ano e mes para substituir no url
    
    cvm_m = consulta_cvm(data)
    if m == 1:
        firstD = cvm_m['DT_COMPTC'].iloc[0] # Verifica qual é o primeiro dia do ano
        lastD = cvm_m['DT_COMPTC'].iloc[len(cvm_m)-1] # Verifica qual é o primeiro dia do ano
        cvm = cvm_m
    else:
        cvm = pd.concat([cvm,cvm_m])
    

    yyyy = str(year)
    if m < 10: mm = '0'+str(m)
    else: mm = str(m)
    
    data = yyyy+mm # Ano e mes para substituir no url
    
    cvm_m = consulta_cvm(data)
    if m == 1:
        firstD = cvm_m['DT_COMPTC'].iloc[0] # Verifica qual é o primeiro dia do ano
        lastD = cvm_m['DT_COMPTC'].iloc[len(cvm_m)-1] # Verifica qual é o primeiro dia do ano
        cvm = cvm_m
    else:
        cvm = pd.concat([cvm,cvm_m])

#path_read = 'I:/GESTAO/5.2) Fundos Abertos/Análises/FI_code/get_history.py'
#wb = openpyxl.load_workbook(abs_path+'/output/output.xlsx')

#sheet = 'MM universe'