# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:07:00 2022

@author: eduardo.scheffer
"""

''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import pandas as pd
import numpy as np
import math as math
import xlwings as xw
import win32com.client
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

import datetime as dt

#import matplotlib.pyplot as plt
#import seaborn as sns

import warnings 
warnings.filterwarnings('ignore')

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", 20)

import sys
import os
import inspect

parent_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))))
sys.path.append(parent_path+'/CR_code/formulas')


from  fund_prices_database import fund_prices_database
#from  fund_charact_database import fund_charact_database
from  benchmark_prices_database import benchmark_prices_database
from drawdown import event_drawdown, max_drawdown
from moving_window import moving_window

''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 

sheet_portfolio = "Criação de Portfólio"

print("Getting portfolio from Excel...")    
excel_path = parent_path + '/Carteira Recomendada.xlsm'
wb = openpyxl.load_workbook(excel_path, data_only=True)
worksheet = wb[sheet_portfolio]
date_first = worksheet['AB2'].value
date_last = worksheet['AB3'].value

portfolio = pd.read_excel(excel_path, sheet_name = sheet_portfolio, header=1).iloc[4:,2:].dropna(how='all',axis=0).dropna(how='all',axis=1) 
portfolio = portfolio.dropna(how='all',axis=0).iloc[:-1,:] # Delete TOTAL line and beyond
portfolio = portfolio.rename(columns=portfolio.iloc[0]).iloc[1:,:] # Turn first row into column headers
portfolio.reset_index(inplace=True,drop=True)
portfolio['R$'] = portfolio['R$'].astype(float)
portfolio['% do PL'] = portfolio['% do PL'].astype(float)
portfolio['Ativo'] = portfolio['Ativo'].astype(str)

columns = list(portfolio.columns)[2:]
columns = ["column_0"] + ["column_1"] + columns
portfolio.columns = columns

classe = ""
estrategia = ""
for i in range(portfolio.shape[0]):
    if portfolio.iloc[i,1] == "y":
        classe = portfolio.iloc[i,2]
    if portfolio.iloc[i,0] == "x" and portfolio.iloc[i,1] == "x":
        estrategia = portfolio.iloc[i,2]
    else:
        portfolio._set_value(i,'Classe', classe)
        portfolio._set_value(i,'Estratégia', estrategia)
        portfolio._set_value(i,'Estratégia', estrategia)

portfolio = portfolio[(portfolio['column_0'] != "x") & (portfolio['column_1'] != "x") & (portfolio['column_1'] != "y")]    

portfolio = portfolio[['Ativo','CNPJ', 'R$', '% do PL', 'Benchmark', '% Benchmark', 'Benchmark +', 'Liquidez (D+)', 'Classe', 'Estratégia']]
dict_CNPJ = dict(zip(portfolio['CNPJ'], portfolio['Ativo']))

benchmark_list = ['CDI','SELIC', 'Ibovespa','IHFA','IPCA','IMA-B','IMA-B 5', 'PTAX', 'SP500', 'Prévia IPCA']

cnpj_list = list(portfolio[(portfolio['CNPJ']!="-")]['CNPJ'].to_numpy().flatten().astype(float))


''' 3) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''

print("Getting fund prices...")
fund_prices = fund_prices_database(cnpj_list, date_first - dt.timedelta(days=62), date_last) # -62 to be able to get IPCA values 
fund_prices.rename(columns=dict_CNPJ, inplace=True)

print("Getting benchmark prices...")
benchmark_prices = benchmark_prices_database(benchmark_list, date_first - dt.timedelta(days=62), date_last) 


''' 4) MANIPULATE PORTFOLIO CATHEGORICAL COLUMNS ----------------------------------------------------------------------------------------------------------'''

print("Getting fixed income rates...")
# FI Rates end benchmark
portfolio.loc[((portfolio['Ativo'].str.contains('CDI', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'CDI'
portfolio.loc[((portfolio['Ativo'].str.contains('IPCA', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'IPCA'
portfolio.loc[((portfolio['Ativo'].str.contains('IMA-B', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'IMA-B'

portfolio['Rates'] = portfolio['Ativo'].str.split(r"(",expand=True)[1].str.replace("CDI","").str.replace("IPCA","").str.replace("IMA-B","").str.replace("IMAB","").str.replace(" ","").str.replace(")","").str.replace("+","")
portfolio.loc[((portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-")),'% Benchmark'] = portfolio[((portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']
portfolio.loc[((portfolio['CNPJ'] == "-") & (portfolio['% Benchmark'].isna())), 'Benchmark +'] = portfolio[((~portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']

portfolio = portfolio.drop(['Rates'], axis = 1)

portfolio['Benchmark +'] = portfolio['Benchmark +'].fillna(0)
portfolio['Benchmark +'] = portfolio['Benchmark +'].astype('str').str.rstrip('%').astype('float') / 100.0
portfolio['% Benchmark'] = portfolio['% Benchmark'].fillna(0)
portfolio['Benchmark'] = portfolio['% Benchmark'].fillna("-")
portfolio['% Benchmark'] = portfolio['% Benchmark'].astype('str').str.rstrip('%').astype('float') / 100.0


''' 5) GET ASSETS AND BENCHMARKS DAILY RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

print("Calculating fund returns...")
# Fund daily returns:
fund_prices.fillna(method='ffill', inplace=True)
fund_lnReturns = np.log(fund_prices.astype('float')) - np.log(fund_prices.astype('float').shift(1))
fund_lnReturns = fund_lnReturns.iloc[1:,:]
fund_lnReturns = fund_lnReturns[((fund_lnReturns.index>=date_first) & (fund_lnReturns.index<=date_last))]

print("Calculating benchmark returns...")
# Delete weekends and Brazilian holidays
benchmark_prices.loc[:,'IPCA'].fillna(method='ffill', inplace=True)
benchmark_prices.loc[:,'SP500'].fillna(method='ffill', inplace=True)
benchmark_prices = benchmark_prices[(~benchmark_prices['CDI'].isna())]
# Fill empty prices (different calendars)
benchmark_prices.fillna(method='ffill', inplace=True)
benchmark_index = list(benchmark_prices.loc[:,'CDI'].index)

# Get number of workdays
benchmark_prices['IPCA_m/yyyy'] = [str(i) + "/"+ str(j) for i, j in zip(list(benchmark_prices.index.month), list(benchmark_prices.index.year))] #get new date column
work_days = benchmark_prices[(~benchmark_prices['CDI'].isna())].groupby(['IPCA_m/yyyy'])['CDI'].count().to_frame() # get work days (DU) for each month
work_days.rename(columns={'CDI':'DU'}, inplace=True)
benchmark_prices = pd.merge(benchmark_prices, work_days, how="left", on=['IPCA_m/yyyy'])
benchmark_prices.index = list(benchmark_index)
benchmark_prices = benchmark_prices.drop(['IPCA_m/yyyy'], axis = 1)

# Benchmark daily returns given as prices:
benchmark_lnReturns = np.log(benchmark_prices.astype('float')) - np.log(benchmark_prices.astype('float').shift(1))
benchmark_lnReturns = benchmark_lnReturns.drop(['DU'], axis = 1)

# CDI, SELIC and Previa IPCA daily returns (given as rates, not prices)
benchmark_lnReturns.loc[:,'CDI'] = (1+benchmark_prices.loc[:,'CDI'].astype('float')/100)**(1/252)-1
benchmark_lnReturns.loc[:,'SELIC'] = (1+benchmark_prices.loc[:,'SELIC'].astype('float')/100)**(1/252)-1
benchmark_lnReturns.loc[:,'Prévia IPCA'] = (1+benchmark_prices.loc[:,'Prévia IPCA'].astype('float')/100)**(1/benchmark_prices.loc[:, "DU"])-1

# IPCA
benchmark_lnReturns.loc[:, "IPCA"] = (1+benchmark_lnReturns.loc[:,'IPCA'])**(1/benchmark_prices.loc[:, "DU"])-1 # calculate daily IPCA rates
benchmark_lnReturns.loc[:, "IPCA"] = benchmark_lnReturns.loc[:, "IPCA"].replace(0,np.nan)

# Use forecasted IPCA as missing IPCA
benchmark_lnReturns['IPCA_m/yyyy'] = [str(i) + "/"+ str(j) for i, j in zip(list(benchmark_lnReturns.index.month), list(benchmark_lnReturns.index.year))] #get new date column
IPCA_est  = benchmark_lnReturns.loc[:,['Prévia IPCA', 'IPCA_m/yyyy']] # Create dataframe to get last Prévia IPCA
last_month = list(benchmark_lnReturns.loc[:, "IPCA"].dropna().index)[-1]
IPCA_est = IPCA_est[(IPCA_est.index > last_month)] # Select only rows greater than last IPCA data
IPCA_est = IPCA_est[(IPCA_est['IPCA_m/yyyy'] != IPCA_est.iloc[0,1])] # Select only rows of later months

# Get latests forecasts of IPCA
for aux_index in IPCA_est.index:
    date_day = aux_index.day
    if date_day <= 15:
        date_month = aux_index.month
        date_year = aux_index.year
        IPCA_est.loc[aux_index, 'IPCA_m/yyyy'] = str(date_month) + "/"+ str(date_year)
    else:
        date_month = (dt.datetime(aux_index.year, aux_index.month, date_day)+ dt.timedelta(days=16)).month
        date_year = (dt.datetime(aux_index.year, aux_index.month, date_day)+ dt.timedelta(days=16)).year
        IPCA_est.loc[aux_index, 'IPCA_m/yyyy'] = str(date_month) + "/"+ str(date_year)

IPCA_est = IPCA_est.drop_duplicates(subset=['IPCA_m/yyyy'], keep='last')
benchmark_index = list(benchmark_lnReturns.index)
benchmark_lnReturns = pd.merge(benchmark_lnReturns, IPCA_est, how="left", on=['IPCA_m/yyyy'])
benchmark_lnReturns.index = list(benchmark_index)
benchmark_lnReturns.loc[~benchmark_lnReturns['Prévia IPCA_y'].isna(),'IPCA'] = benchmark_lnReturns['Prévia IPCA_y'] # Fill missing IPCA with forecasted value

benchmark_lnReturns.fillna(method='ffill', inplace=True)
benchmark_lnReturns = benchmark_lnReturns.drop(columns = ['IPCA_m/yyyy', 'Prévia IPCA_x', 'Prévia IPCA_y'])

benchmark_lnReturns = benchmark_lnReturns[((benchmark_lnReturns.index>=date_first) & (benchmark_lnReturns.index<=fund_lnReturns.index[-1]))]

''' 6) CALCULATE PORTFOLIO --------------------------------------------------------------------------------------------------------------------------------'''

portfolio = portfolio.set_index('Ativo')
portfolio_returns = pd.DataFrame(index = benchmark_lnReturns.index, columns = list(portfolio.index))

# Get asset returns
for i in portfolio_returns.columns:
    if portfolio.loc[i,'CNPJ'] != "-":
        portfolio_returns[i] = fund_lnReturns.loc[portfolio_returns.index, i]
    elif portfolio.loc[i,'% Benchmark'] != 0:
        portfolio_returns[i] = benchmark_lnReturns.loc[portfolio_returns.index, portfolio.loc[i,'Benchmark']] * portfolio.loc[i,'% Benchmark']
    elif portfolio.loc[i,'Benchmark +'] != 0:
        if portfolio.loc[i,'Benchmark'] != "-":
            portfolio_returns[i] = (1+benchmark_lnReturns.loc[portfolio_returns.index, portfolio.loc[i,'Benchmark']]) * (1+portfolio.loc[i,'% Benchmark']) - 1
        else:
            benchmark_lnReturns.loc[portfolio_returns.index, portfolio.loc[i,'Benchmark']]
