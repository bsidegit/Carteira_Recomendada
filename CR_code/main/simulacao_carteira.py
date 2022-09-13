# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:07:00 2022

@author: eduardo.scheffer
"""

''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import pandas as pd
import numpy as np
import math as math

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
from  index_prices_database import index_prices_database
from drawdown import event_drawdown, max_drawdown
from moving_window import moving_window
#from  fund_prices_cvm import fund_prices_cvm

''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 

sheet_portfolio = "Portfólio Sugerido"

excel_path = parent_path + '/Carteira Recomendada.xlsm'

portfolio = pd.read_excel(excel_path, sheet_name = sheet_portfolio, header=1).iloc[8:,2:].dropna(how='all',axis=0).dropna(how='all',axis=1) 
portfolio = portfolio[(portfolio.iloc[:,0].isna()) & (portfolio.iloc[:,1].isna())].iloc[:,2:] # Delete assets classes
portfolio = portfolio.rename(columns=portfolio.iloc[0]).iloc[1:,:] # Turn first row into column headers
portfolio = portfolio[['Ativo','CNPJ', 'R$', '% do PL', 'Benchmark', '% Benchmark', 'Benchmark +']]
portfolio['R$'] = portfolio['R$'].astype(float)
portfolio['% do PL'] = portfolio['% do PL'].astype(float)
portfolio['Benchmark +'] = portfolio['Benchmark +'].astype(float)
portfolio['Ativo'] = portfolio['Ativo'].astype(str)

benchmark_list = portfolio[(~portfolio['Benchmark'].isna())]['Benchmark'].unique()

#funds_name = list(funds_list.iloc[:,0:1].to_numpy().flatten())


cnpj_list = list(portfolio[(portfolio['CNPJ']!="-")]['CNPJ'].to_numpy().flatten().astype(float))


''' 3) IMPORT BENCHMARKS AND FUNDS PRICES AND CHARACTERISTICS --------------------------------------------------------------------------------------'''

#benchmark_list = list(["CDI", "IHFA", "Ibovespa"])
   
benchmark_prices = index_prices_database(benchmark_list)    
fund_price = fund_prices_database(cnpj_list) 

fund_charact = fund_charact_database(cnpj_list) 
portfolio = pd.merge(portfolio, fund_charact, how='outer', on='CNPJ')
portfolio.rename(columns={'ClasseBSide':'Classe'}, inplace=True)

''' 4) MANIPULATE CATHEGORICAL COLUMNS -----------------------------------------------------------------------------------------'''

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

# Liquidity
