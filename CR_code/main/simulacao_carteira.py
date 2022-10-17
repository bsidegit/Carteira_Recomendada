# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:07:00 2022

@author: eduardo.scheffer
"""

''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import pandas as pd
import numpy as np
import xlwings as xw
#import win32com.client
import openpyxl
#from openpyxl.utils.dataframe import dataframe_to_rows
#import matplotlib.pyplot as plt

import datetime as dt
from dateutil.relativedelta import relativedelta

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
from  benchmark_prices_database import benchmark_prices_database

''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 

sheet_portfolio = "Criação de Portfólio"

print("Getting portfolio from Excel...")    
excel_path = parent_path + '/Carteira Recomendada.xlsm'
wb = openpyxl.load_workbook(excel_path, data_only=True)
worksheet = wb[sheet_portfolio]
date_first = worksheet['AB2'].value
date_last = worksheet['AB3'].value
benchmark = worksheet['AB4'].value
amount = worksheet['M3'].value


portfolio = pd.read_excel(excel_path, sheet_name = sheet_portfolio, header=1).iloc[4:,2:].dropna(how='all',axis=0).dropna(how='all',axis=1) 
portfolio = portfolio.dropna(how='all',axis=0).iloc[:-1,:] # Delete TOTAL line and beyond
portfolio = portfolio.rename(columns=portfolio.iloc[0]).iloc[1:,:] # Turn first row into column headers
portfolio.reset_index(inplace=True,drop=True)
portfolio['R$'] = portfolio['R$'].astype(float)
portfolio['% do PL'] = portfolio['% do PL'].astype(float)
portfolio['% do PL'] = portfolio['% do PL'].astype(float)
portfolio['Liquidez (D+)'] = portfolio['Liquidez (D+)'].replace(["-"], np.nan)
portfolio['Liquidez (D+)'] = portfolio['Liquidez (D+)'].astype(float)

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
portfolio['Benchmark +'] = portfolio['Benchmark +'].astype('str').str.rstrip('%').astype('float')
portfolio['% Benchmark'] = portfolio['% Benchmark'].fillna(0)
portfolio['% Benchmark'] = portfolio['% Benchmark'].astype('str').str.rstrip('%').astype('float')
portfolio['Benchmark'] = portfolio['Benchmark'].fillna("-")

portfolio.loc[(portfolio['CNPJ'] == "-"),'Benchmark +'] = portfolio.loc[(portfolio['CNPJ'] == "-"),'Benchmark +'] / 100
portfolio.loc[(portfolio['CNPJ'] == "-"),'% Benchmark'] = portfolio.loc[(portfolio['CNPJ'] == "-"),'% Benchmark'] / 100

''' 5) GET ASSETS AND BENCHMARKS DAILY RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

print("Calculating fund returns...")
# Fund daily returns:
fund_prices.fillna(method='ffill', inplace=True)
fund_lnReturns = fund_prices.astype('float') / fund_prices.astype('float').shift(1) - 1
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
benchmark_prices_freeze = benchmark_prices.copy() # line created to validation purposes only

# Benchmark daily returns given as prices:
benchmark_lnReturns = benchmark_prices.astype('float') / benchmark_prices.astype('float').shift(1) - 1
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

''' 6) CALCULATE PORTFOLIO RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

portfolio = portfolio.set_index('Ativo')
assets_returns = pd.DataFrame(index = benchmark_lnReturns.index, columns = list(portfolio.index))
print("Calculating portfolio returns...")

# Get asset returns
for i in assets_returns.columns:
    
    if portfolio.loc[i,'CNPJ'] != "-": # Fund returns
        assets_returns[i] = fund_lnReturns.loc[assets_returns.index, i]
        
        if portfolio.loc[i,'% Benchmark'] != 0: # Fund with % Benchmark proxy
           assets_returns.loc[assets_returns[i].isna(), i] = benchmark_lnReturns.loc[assets_returns[i].isna(), portfolio.loc[i,'Benchmark']] * portfolio.loc[i,'% Benchmark']
        
        elif portfolio.loc[i,'Benchmark +'] != 0:
            if portfolio.loc[i,'Benchmark'] != "-": # Fund with Benchmark+ proxy
                assets_returns.loc[assets_returns[i].isna(), i] = (1+benchmark_lnReturns.loc[assets_returns[i].isna(), portfolio.loc[i,'Benchmark']]) * ((1+portfolio.loc[i,'Benchmark +'])**(1/252)) - 1
            
            else: # Fund with prefixed proxy
                assets_returns.loc[assets_returns[i].isna(), i] = (1+portfolio.loc[i,'Benchmark +'])**(1/252) - 1
                
    elif portfolio.loc[i,'% Benchmark'] != 0: # Fixed income % Benchmark returns
        assets_returns[i] = benchmark_lnReturns.loc[assets_returns.index, portfolio.loc[i,'Benchmark']] * portfolio.loc[i,'% Benchmark']
    
    elif portfolio.loc[i,'Benchmark +'] != 0:
        if portfolio.loc[i,'Benchmark'] != "-": # Fixed income Benchmark+ returns
            assets_returns[i] = (1+benchmark_lnReturns.loc[assets_returns.index, portfolio.loc[i,'Benchmark']]) * ((1+portfolio.loc[i,'Benchmark +'])**(1/252)) - 1
        
        else: # Fixed income prefixed returns
            assets_returns[i] = (1+portfolio.loc[i,'Benchmark +'])**(1/252) - 1

# Get weighted performance for each asset
for i in assets_returns.index:
    assets_returns.loc[i,:] = assets_returns.loc[i,:].mul(portfolio['% do PL']) # performance attribution per product
            
# PORTFOLIO
portfolio_return = assets_returns.sum(axis=1) # Daily returns

portfolio_acc = portfolio_return.copy()
portfolio_acc.name = 'Portfolio Cliente'
for i in range(portfolio_acc.shape[0]-1):
    portfolio_acc[i+1] = (1 + portfolio_acc[i+1]) * (1+ portfolio_acc[i]) - 1 # Accumulated returns

# ASSET STRATEGIES
strategy_returns = pd.DataFrame(index = benchmark_lnReturns.index , columns = list(portfolio['Estratégia'].unique()))
for i in strategy_returns.columns:
    asset_group = []
    for j in assets_returns.columns:
        if portfolio.loc[j,'Estratégia'] == i:
            asset_group = asset_group + [j]        
    strategy_returns[i] = assets_returns[asset_group].sum(axis=1)  # Daily returns
    
strategy_acc = strategy_returns .copy()   
for i in range(strategy_acc.shape[0]-1):
    strategy_acc.iloc[i+1,:] = (1 + strategy_acc.iloc[i+1,:]) * (1 + strategy_acc.iloc[i,:]) - 1 # Accumulated returns

    
# ASSET CLASSES
class_returns = pd.DataFrame(index = benchmark_lnReturns.index , columns = list(portfolio['Classe'].unique()))        
for i in class_returns.columns:
    asset_group = []
    for j in assets_returns.columns:
        if portfolio.loc[j,'Classe'] == i:
            asset_group = asset_group + [j]
    class_returns[i] = assets_returns[asset_group].sum(axis=1)    # Daily returns
    
class_acc = class_returns.copy()    
for i in range(class_acc.shape[0]-1):
    class_acc.iloc[i+1,:] = (1 + class_acc.iloc[i+1,:]) * (1 + class_acc.iloc[i,:]) - 1 # Accumulated returns
    
''' 7) CALCULATE PERFORMANCE METRICS --------------------------------------------------------------------------------------------------------------------------------'''
print("Calculating performance metrics...")

# Performance attribution
class_perf_attr = class_acc.iloc[-1,:]
strategy_perf_attr = strategy_acc.iloc[-1,:]

# Portfolio vs. Benchmark: Return, Vol, Sharpe
if benchmark == "": benchmark = "CDI" 

portf_vs_bench_1 = pd.DataFrame(columns = ["Rentabilidade Anualizada", "Volatilidade Anualizada", "Sharpe"],
                                  index = ["Portfólio Sugerido", benchmark])


portf_vs_bench_1.iloc[0,0] = portfolio_acc[-1]
portf_vs_bench_1.iloc[1,0] = (benchmark_lnReturns[benchmark]+1).to_numpy().prod() - 1
portf_vs_bench_1.iloc[0,1] = np.sqrt(252)*np.std(portfolio_return)
portf_vs_bench_1.iloc[1,1] = np.sqrt(252)*np.std(benchmark_lnReturns[benchmark])
portf_vs_bench_1.iloc[0,2] = ((1+portf_vs_bench_1.iloc[0,0])**(252/benchmark_lnReturns.shape[0]) - 
                              ((benchmark_lnReturns['CDI']+1).to_numpy().prod())**(252/benchmark_lnReturns.shape[0])) / portf_vs_bench_1.iloc[0,1]

# Portfolio vs. Benchmark: Returns
portf_vs_bench_2 = pd.DataFrame(columns = ["Mes", "Ano", "6 meses", "12 meses", "2 anos", "Saldo Inicial ("+date_first.strftime("%d/%m/%Y")+")", "Saldo Final ("+date_last.strftime("%d/%m/%Y")+")"],
                                  index = ["Portfólio Sugerido", benchmark])

date_MTD = dt.date(portfolio_return.index[-1].year, portfolio_return.index[-1].month, 1)
date_YTD = dt.date(portfolio_return.index[-1].year, 1, 1)
date_6M = portfolio_return.index[-1] - relativedelta(months=6)
date_12M = portfolio_return.index[-1] - relativedelta(months=12)
date_24M = portfolio_return.index[-1] - relativedelta(months=24)

portf_vs_bench_2.iloc[0,0] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_MTD)]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[1,0] = (1+benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date > date_MTD)][benchmark]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[0,1] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_YTD)]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[1,1] = (1+benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >date_YTD)][benchmark]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[0,2] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_6M)]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[1,2] = (1+benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date > date_6M)][benchmark]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[0,3] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_12M)]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[1,3] = (1+benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date > date_12M)][benchmark]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[0,4] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_24M)]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[1,4] = (1+benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date > date_24M)][benchmark]).to_numpy().prod() - 1
portf_vs_bench_2.iloc[0,5] = amount
portf_vs_bench_2.iloc[1,5] = "-"
portf_vs_bench_2.iloc[0,6] = amount * portf_vs_bench_1.iloc[0,0]
portf_vs_bench_2.iloc[1,6] = "-"


# Portfolio vs. Benchmark: Monthly results
portf_vs_bench_aux = portfolio_return.copy()
portf_vs_bench_aux.name = "Portfólio Sugerido"
portf_vs_bench_aux = pd.concat([portf_vs_bench_aux, benchmark_lnReturns[benchmark]], axis=1)
portf_vs_bench_aux.index = portf_vs_bench_aux.index.strftime("%Y-%m")

portf_vs_bench_3 = pd.DataFrame(index = np.unique(list(portf_vs_bench_aux.index)), columns = ["Portfólio Sugerido", benchmark])
for i in portf_vs_bench_3.index:
    for j in portf_vs_bench_3.columns:
        portf_vs_bench_3.loc[i,j] = (1 + portf_vs_bench_aux[(portf_vs_bench_aux.index == i)][j]).to_numpy().prod() - 1

portf_vs_bench_3['%'+benchmark] = portf_vs_bench_3["Portfólio Sugerido"] / portf_vs_bench_3[benchmark]


# Portfolio vs. Benchmark: Statistics
portf_vs_bench_4 = pd.DataFrame(columns = ["Meses Positivos", "Meses Negativos", "Maior Retorno", "Menor Retorno", "Acima do CDI (meses)", "Abaixo do CDI (meses)"],
                                  index = ["Portfólio Sugerido", benchmark])

portf_vs_bench_4.iloc[0,0] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]>=0)]["Portfólio Sugerido"].count()
portf_vs_bench_4.iloc[1,0] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]>=0)][benchmark].count()
portf_vs_bench_4.iloc[0,1] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]<0)]["Portfólio Sugerido"].count()
portf_vs_bench_4.iloc[1,1] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]<0)][benchmark].count()
portf_vs_bench_4.iloc[0,2] = portf_vs_bench_3["Portfólio Sugerido"].max()
portf_vs_bench_4.iloc[1,2] = portf_vs_bench_3[benchmark].max()
portf_vs_bench_4.iloc[0,3] = portf_vs_bench_3["Portfólio Sugerido"].min()
portf_vs_bench_4.iloc[1,3] = portf_vs_bench_3[benchmark].min()

if benchmark == "CDI":
    portf_vs_bench_4.iloc[0,4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]>=portf_vs_bench_3["CDI"])]["Portfólio Sugerido"].count()
    portf_vs_bench_4.iloc[0,5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]<portf_vs_bench_3["CDI"])]["Portfólio Sugerido"].count()
    portf_vs_bench_4.iloc[1,4] = 0
    portf_vs_bench_4.iloc[1,5] = 0
else:
    retorno_CDI_M = benchmark_lnReturns["CDI"]
    retorno_CDI_M.name = "retorno_CDI_M"
    retorno_CDI_M.index = retorno_CDI_M.index.strftime("%Y-%m")
    retorno_CDI_M = retorno_CDI_M.groupby(level=0,axis=0).sum()
    portf_vs_bench_4.iloc[0,4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]>=retorno_CDI_M)]["Portfólio Sugerido"].count()
    portf_vs_bench_4.iloc[0,5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Sugerido"]<retorno_CDI_M)]["Portfólio Sugerido"].count()
    portf_vs_bench_4.iloc[1,4] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]>=retorno_CDI_M)][benchmark].count()
    portf_vs_bench_4.iloc[1,5] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]<retorno_CDI_M)][benchmark].count()


# Portfolio vs. Benchmark: Volatility
portf_vs_bench_5 = pd.DataFrame(columns = ["Mes", "Ano", "6 meses", "12 meses", "2 anos"],
                                  index = ["Portfólio Sugerido", benchmark])
    
portf_vs_bench_5.iloc[0,0] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_MTD)])
portf_vs_bench_5.iloc[1,0] = np.sqrt(252)*np.std(benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >= date_MTD)][benchmark])
portf_vs_bench_5.iloc[0,1] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_YTD)])
portf_vs_bench_5.iloc[1,1] = np.sqrt(252)*np.std(benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >= date_YTD)][benchmark])
portf_vs_bench_5.iloc[0,2] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_6M)])
portf_vs_bench_5.iloc[1,2] = np.sqrt(252)*np.std(benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >= date_6M)][benchmark])
portf_vs_bench_5.iloc[0,3] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_12M)])
portf_vs_bench_5.iloc[1,3] = np.sqrt(252)*np.std(benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >= date_12M)][benchmark])
portf_vs_bench_5.iloc[0,4] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_24M)])
portf_vs_bench_5.iloc[1,4] = np.sqrt(252)*np.std(benchmark_lnReturns[(pd.to_datetime(benchmark_lnReturns.index).date >= date_24M)][benchmark])

volatility = pd.concat([portfolio_return, benchmark_lnReturns[benchmark]], axis=1)
volatility.columns = ["Portfólio Sugerido", benchmark]

for i in range(volatility.shape[0]-21): # Rolling 21-days volatility
    volatility.iloc[i+21,0] = np.sqrt(252)*np.std(portfolio_return[i:i+21]) 
    volatility.iloc[i+21,1] = np.sqrt(252)*np.std(benchmark_lnReturns[benchmark][i:i+21])

volatility = volatility.iloc[21:,:]

# Portfolio vs. Benchmark: Drawdown
portf_vs_bench_6 = pd.DataFrame(columns = ["Drawdown Máximo", "Mes", "Ano", "6 meses", "12 meses", "2 anos"],
                                  index = ["Portfólio Sugerido", benchmark])

drawdown = pd.DataFrame(np.zeros((portfolio_return.shape[0], 2)), columns = ["Portfólio Sugerido", benchmark], index = portfolio_return.index)

drawdown_acc_returns = pd.concat([portfolio_acc, benchmark_lnReturns[benchmark]], axis=1) # Initialize accumulated returns to calculate drawdowns
drawdown_acc_returns.columns = ["Portfólio Sugerido", benchmark]
for i in range(drawdown.shape[0]-1):
    drawdown_acc_returns[benchmark][i+1] = (1+drawdown_acc_returns[benchmark][i+1]) * (1+drawdown_acc_returns[benchmark][i])-1 # Accumulated benchmark returns

drawdown_MTD = drawdown[(pd.to_datetime(drawdown.index).date >= date_MTD)]
drawdown_YTD = drawdown[(pd.to_datetime(drawdown.index).date >= date_YTD)]
drawdown_6M = drawdown[(pd.to_datetime(drawdown.index).date >= date_6M)]
drawdown_12M= drawdown[(pd.to_datetime(drawdown.index).date >= date_12M)]
drawdown_24M = drawdown[(pd.to_datetime(drawdown.index).date >= date_24M)]
                        
# Calculate drawdown series (entire period)
for i in range(drawdown.shape[0]-1):
    drawdown.iloc[i+1,:] = (1+drawdown.iloc[i,:]) * (1+drawdown_acc_returns.iloc[i+1,:]) / (1+drawdown_acc_returns.iloc[i,:]) - 1
    drawdown.iloc[i+1,:].values[drawdown.iloc[i+1,:].values > 0] = 0

# Calculate drawdown series (MTD)
for i in range(drawdown_MTD.shape[0]-1):
    drawdown_MTD.iloc[i+1,:] = (1+drawdown_MTD.iloc[i,:]) * (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_MTD)].iloc[i+1,:]) / (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_MTD)].iloc[i,:]) - 1
    drawdown_MTD.iloc[i+1,:].values[drawdown_MTD.iloc[i+1,:].values > 0] = 0

# Calculate drawdown series (YTD)
for i in range(drawdown_YTD.shape[0]-1):
    drawdown_YTD.iloc[i+1,:] = (1+drawdown_YTD.iloc[i,:]) * (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_YTD)].iloc[i+1,:]) / (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_YTD)].iloc[i,:]) - 1
    drawdown_YTD.iloc[i+1,:].values[drawdown_YTD.iloc[i+1,:].values > 0] = 0

# Calculate drawdown series (6M)
for i in range(drawdown_6M.shape[0]-1):
    drawdown_6M.iloc[i+1,:] = (1+drawdown_6M.iloc[i,:]) * (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_6M)].iloc[i+1,:]) / (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_6M)].iloc[i,:]) - 1
    drawdown_6M.iloc[i+1,:].values[drawdown_6M.iloc[i+1,:].values > 0] = 0

# Calculate drawdown series (12M)
for i in range(drawdown_12M.shape[0]-1):
    drawdown_12M.iloc[i+1,:] = (1+drawdown_12M.iloc[i,:]) * (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_12M)].iloc[i+1,:]) / (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_12M)].iloc[i,:]) - 1
    drawdown_12M.iloc[i+1,:].values[drawdown_12M.iloc[i+1,:].values > 0] = 0
# Calculate drawdown series (24M)
for i in range(drawdown_24M.shape[0]-1):
    drawdown_24M.iloc[i+1,:] = (1+drawdown_24M.iloc[i,:]) * (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_24M)].iloc[i+1,:]) / (1+drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_24M)].iloc[i,:]) - 1 
    drawdown_24M.iloc[i+1,:].values[drawdown_24M.iloc[i+1,:].values > 0] = 0      

portf_vs_bench_6.iloc[0,0] = "{:.2%}".format(min(drawdown.iloc[:,0])) + "("+ drawdown[(drawdown.iloc[:,0] == min(drawdown.iloc[:,0]))].index.strftime('%d/%m/%Y') + ")"
portf_vs_bench_6.iloc[1,0] = "{:.2%}".format(min(drawdown.iloc[:,1])) + "("+ drawdown[(drawdown.iloc[:,1] == min(drawdown.iloc[:,1]))].index.strftime('%d/%m/%Y') + ")"
portf_vs_bench_6.iloc[0,1] = min(drawdown_MTD.iloc[:,0])
portf_vs_bench_6.iloc[1,1] = min(drawdown_MTD.iloc[:,1])
portf_vs_bench_6.iloc[0,2] = min(drawdown_YTD.iloc[:,0])
portf_vs_bench_6.iloc[1,2] = min(drawdown_YTD.iloc[:,1])
portf_vs_bench_6.iloc[0,3] = min(drawdown_6M.iloc[:,0])
portf_vs_bench_6.iloc[1,3] = min(drawdown_6M.iloc[:,1])
portf_vs_bench_6.iloc[0,4] = min(drawdown_12M.iloc[:,0])
portf_vs_bench_6.iloc[1,4] = min(drawdown_12M.iloc[:,1])
portf_vs_bench_6.iloc[0,5] = min(drawdown_24M.iloc[:,0])
portf_vs_bench_6.iloc[1,5] = min(drawdown_24M.iloc[:,1])


''' 8) PRINT TO EXCEL --------------------------------------------------------------------------------------------------------------------------------'''

print("Writing on Excel...")

# Create workbook object (try to opeen an existing one, if it doesn`t exist, create one)
try:
    wb = xw.Book(excel_path)
except:
    wb = xw.Book()
    wb.save(excel_path)
    wb = xw.Book(excel_path)

sNamList = [sh.name for sh in wb.sheets]

            
# Print Benchmark Prices:
output_sheet = 'Bench Prices'        
data = benchmark_prices_freeze

if not output_sheet in sNamList:
    sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data

# Print Benchmark Returns:
output_sheet = 'Bench Returns'        
data = benchmark_lnReturns

if not output_sheet in sNamList:
    sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Funds Prices:
output_sheet = 'Fund Prices'        
data = fund_prices

if not output_sheet in sNamList:  sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data

# Print Funds Returns:
output_sheet = 'Fund Returns'        
data = fund_lnReturns

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Assets Returns:
output_sheet = 'Assets Returns'        
data = assets_returns

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio Returns:
output_sheet = 'Portfolio Returns'        
data = portfolio_return

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio Acc Returns:
output_sheet = 'Portfolio AccReturns'        
data = portfolio_acc

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Strategy Returns:
output_sheet = 'Strategy Returns'        
data = strategy_returns

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Strategy Acc Returns:
output_sheet = 'Strategy AccReturns'        
data = strategy_acc

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Class Returns:
output_sheet = 'Class Returns'        
data = class_returns

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Class Acc Returns:
output_sheet = 'Class AccReturns'        
data = class_acc

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio Volatility:
output_sheet = 'Volatility'        
data = volatility

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data

drawdown
# Print Portfolio Drawdown:
output_sheet = 'Drawdown'        
data = drawdown

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 1:
output_sheet = 'Port_vs_Bench 1'        
data = portf_vs_bench_1

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 2:
output_sheet = 'Port_vs_Bench 2'        
data = portf_vs_bench_2

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 3:
output_sheet = 'Port_vs_Bench 3'        
data = portf_vs_bench_3

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 4:
output_sheet = 'Port_vs_Bench 4'        
data = portf_vs_bench_4

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 5:
output_sheet = 'Port_vs_Bench 5'        
data = portf_vs_bench_5

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data


# Print Portfolio vs Bench 6:
output_sheet = 'Port_vs_Bench 6'        
data = portf_vs_bench_6

if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
else: sheet = wb.sheets[output_sheet]

sheet.clear_contents() # Delete old data
sheet.range((2, 1)).value = data

print("Completed.")