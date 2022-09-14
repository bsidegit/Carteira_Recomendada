# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:07:00 2022

@author: eduardo.scheffer
"""

''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import pandas as pd
import numpy as np
import math as math

import pyodbc
from tqdm import tqdm # progress bar
import datetime as dt

import datetime as dt

import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

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
from  fund_charact_database import fund_charact_database
from  benchmark_prices_database import benchmark_prices_database
from drawdown import event_drawdown, max_drawdown
from moving_window import moving_window
#from  fund_prices_cvm import fund_prices_cvm

''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 

sheet_portfolio = "Portfólio Sugerido"

excel_path = parent_path + '/Carteira Recomendada.xlsm'
wb = openpyxl.load_workbook(excel_path, data_only=True)
worksheet = wb[sheet_portfolio]
date_first = worksheet['AB5'].value
date_last = worksheet['AB7'].value

portfolio = pd.read_excel(excel_path, sheet_name = sheet_portfolio, header=1).iloc[8:,2:].dropna(how='all',axis=0).dropna(how='all',axis=1) 
portfolio = portfolio[(portfolio.iloc[:,0].isna()) & (portfolio.iloc[:,1].isna())].iloc[:,2:] # Delete assets classes
portfolio = portfolio.rename(columns=portfolio.iloc[0]).iloc[1:,:] # Turn first row into column headers
portfolio = portfolio[['Ativo','CNPJ', 'R$', '% do PL', 'Benchmark', '% Benchmark', 'Benchmark +', 'Liquidez (D+)']]
portfolio['R$'] = portfolio['R$'].astype(float)
portfolio['% do PL'] = portfolio['% do PL'].astype(float)
portfolio['Benchmark +'] = portfolio['Benchmark +'].astype(float)
portfolio['Ativo'] = portfolio['Ativo'].astype(str)

portfolio.rename(columns={'Liquidez (D+)':'Liq RF'}, inplace=True)

benchmark_list = ['CDI','SELIC', 'Ibovespa','IHFA','IPCA','IMA-B','IMA-B 5', 'PTAX', 'SP500', 'Prévia IPCA']


#funds_name = list(funds_list.iloc[:,0:1].to_numpy().flatten())


cnpj_list = list(portfolio[(portfolio['CNPJ']!="-")]['CNPJ'].to_numpy().flatten().astype(float))


''' 2) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''

benchmark_prices = benchmark_prices_database(benchmark_list, date_first, date_last)
fund_prices = fund_prices_database(cnpj_list, date_first, date_last)  

fund_charact = fund_charact_database(cnpj_list) 

portfolio = pd.merge(portfolio, fund_charact, how='outer', on='CNPJ')
portfolio.rename(columns={'ClasseBSide':'Classe'}, inplace=True)

''' 4) MANIPULATE CATHEGORICAL COLUMNS ------------------------------------------------------------------------------------------------------------'''

# FX and geography
portfolio['Moeda'] = portfolio['Moeda'].fillna('BRL')
portfolio.loc[(portfolio['CNPJ']!="-"), 'Geografia'] = 'Global' # Se for fundo, assume exposição global
portfolio['Geografia'] = portfolio['Geografia'].fillna('Brasil') # Se for RF, assume exposição Brasil

# Class
portfolio.loc[(portfolio['Ativo'].str.contains('Título', regex=False)),'Classe'] = 'Renda Fixa'
portfolio.loc[(portfolio['Ativo'].str.contains('NTN', regex=False)),'Classe'] = 'Renda Fixa'
portfolio.loc[(portfolio['Ativo'].str.contains('LFT', regex=False)),'Classe'] = 'Renda Fixa'

# Strategy
portfolio.loc[(portfolio['Ativo'].str.contains('CDI', regex=True)), 'Estratégia'] = 'Pós-fixado'
portfolio.loc[(portfolio['Ativo'].str.contains('IPCA', regex=True)), 'Estratégia'] = 'Inflação'
portfolio['Estratégia'] = portfolio['Estratégia'].fillna('Pré-fixado')

# FI Rates end benchmark

portfolio.loc[((portfolio['Ativo'].str.contains('CDI', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'CDI'
portfolio.loc[((portfolio['Ativo'].str.contains('IPCA', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'IPCA'

portfolio['Rates'] = portfolio['Ativo'].str.split(r"(",expand=True)[1].str.replace("CDI","").str.replace("IPCA","").str.replace(" ","").str.replace(")","").str.replace("+","")
portfolio.loc[((portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-")),'% Benchmark'] = portfolio[((portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']
portfolio.loc[((portfolio['CNPJ'] == "-") & (portfolio['% Benchmark'].isna())), 'Benchmark +'] = portfolio[((~portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']

portfolio = portfolio.drop(['Rates'], axis = 1)

# FI Liquidity
portfolio.loc[(portfolio['CNPJ'] == "-"), 'Liquidez (D+)'] = portfolio[(portfolio['CNPJ'] == "-")]['Liq RF']
portfolio = portfolio.drop(['Liq RF'], axis = 1)


''' 5) GET INDEX DAILY RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

# Fund daily returns:
fund_lnReturns = np.log(fund_prices.astype('float')) - np.log(fund_prices.astype('float').shift(1))
fund_lnReturns = fund_lnReturns.iloc[1:,:]

#AJUSTAR FERIADOS/FDS NO BRASIL E EUA
#fund_prices = fund_prices[(~fund_prices['CDI'].isna())] # Delete weekends 

# CDI, SELIC and Previa IPCA daily returns (given as rates, not prices)
benchmark_lnReturns = np.log(benchmark_prices.astype('float')) - np.log(benchmark_prices.astype('float').shift(1))
benchmark_lnReturns.loc[:,'CDI'] = (1+benchmark_prices.loc[:,'CDI'].astype('float')/100)**(1/252)-1
benchmark_lnReturns.loc[:,'SELIC'] = (1+benchmark_prices.loc[:,'SELIC'].astype('float')/100)**(1/252)-1
benchmark_lnReturns.loc[:,'Prévia IPCA'] = (1+benchmark_prices.loc[:,'Prévia IPCA'].astype('float')/100)**(1/252)-1

# IPCA daily returns (given as rates once a month, and forecasted for last month)
benchmark_lnReturns['IPCA_m/yyyy'] = [str(i) + "/"+ str(j) for i, j in zip(list(benchmark_lnReturns.index.month), list(benchmark_lnReturns.index.year))]

benchmark_IPCA = benchmark_lnReturns.loc[:,['IPCA','IPCA_m/yyyy']]
benchmark_IPCA.loc[:,'IPCA'] = benchmark_prices['IPCA']
benchmark_IPCA = benchmark_IPCA.dropna()
benchmark_IPCA.loc[:,'IPCA'] = np.log(benchmark_IPCA.loc[:,'IPCA'].astype('float')).shift(-1) - np.log(benchmark_IPCA.loc[:,'IPCA'].astype('float'))
benchmark_IPCA = benchmark_IPCA.dropna()

benchmark_index = list(benchmark_lnReturns.index)
benchmark_lnReturns = pd.merge(benchmark_lnReturns, benchmark_IPCA, how="left", on=['IPCA_m/yyyy'])
benchmark_lnReturns.index = list(benchmark_index)
benchmark_lnReturns.loc[:,'IPCA_x'] = benchmark_lnReturns.loc[:,'IPCA_y']
benchmark_lnReturns = benchmark_lnReturns.drop(['IPCA_y','IPCA_m/yyyy' ], axis = 1)
benchmark_lnReturns.rename(columns={'IPCA_x':'IPCA'}, inplace=True)

benchmark_lnReturns.loc[:,'IPCA'] = (1+benchmark_lnReturns.loc[:,'IPCA'].astype('float'))**(1/252)-1
benchmark_lnReturns['IPCA'] = benchmark_lnReturns['IPCA'].fillna(benchmark_lnReturns['Prévia IPCA'].iat[-1]) # Assume IPCA for the month yet to be calculated is equal to last forecast "Prévia IPCA"

benchmark_lnReturns = benchmark_lnReturns[(benchmark_lnReturns.index >= date_first)] 

# Delete weekends/holidays (based on nan in CDI column):



