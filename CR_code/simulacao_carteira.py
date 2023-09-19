# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:07:00 2022

@author: eduardo.scheffer
"""

#Compiler PyInstaller (change the paths!!! Works for Eduardo): pyinstaller --onedir --add-binary "C:\\Users\eduardo.scheffer\\.conda\\envs\\bside_clean\\Lib\\site-packages\\pywin32_system32\\pythoncom39.dll;." --paths C:\Users\\eduardo.scheffer\\.conda\\envs\\bside_clean\\Lib\\site-packages run_CR.py

# To compile executable file:
# 1. Delete older files from local paste (build, _pycache_, CR_code, dist)
# 2. Copy new "CR_code"  paste on local folder
# 3. Add "." before "from formulas..." as from .formulas..." below in the "simulacao_carteira.py" file copy in the local folder
# 4. Comment main_code() line in the end of file to avoid double execution
# 5. Check excel file_name
# 6. Go to Anaconda Prompt (using the chosen environment >>"conda activate bside_clean") up to the local folder (>>cd [local address])
# 7. Paste and run the Compiler PyInstaller code line above
# 8. Copy and paste the new generated "dist" folder into the original intranet folder (I:\GESTAO\3) Carteira Recomendada\Carteira Recomendada - Pack Simulação)

# Files needed to run the file standalone: "dist" folder, run_CR, "RUN_Excel" (Windows Batch File)


''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------''' 
import pandas as pd
import numpy as np
import xlwings as xw
import win32com.client
#import openpyxl
#from openpyxl.utils.dataframe import dataframe_to_rows

import datetime as dt
from dateutil.relativedelta import relativedelta

import warnings 
warnings.filterwarnings('ignore')

import sys
import os
import inspect

parent_path = os.path.dirname(os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))))
sys.path.append(parent_path+'/CR_code/formulas')

from  formulas.fund_prices_database import fund_prices_database
from  formulas.benchmark_prices_database import benchmark_prices_database
from  formulas.stock_prices_database import stock_prices_database
from  formulas.fixed_income_prices_database import fixed_income_prices_database

def main_code():
    
    ''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------''' 
    print("Initiating simulation...") 
    print("Getting portfolio from Excel...")     
    
    file_name = 'Carteira Recomendada - MFO.xlsm'
    sheet_name = "Criação de Portfólio"

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.Workbooks(file_name)
        ws = wb.Worksheets(sheet_name)
        flag_MFO = 1
        # File is already open
    except Exception as e:
        print(f"The file '{file_name}' is not open. Opening 'Carteira Recomendada - AAI.xlsm' instead.")
        file_name = 'Carteira Recomendada - AAI.xlsm'
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks(file_name)
        ws = wb.Worksheets(sheet_name)
        flag_MFO = 0
    
    date_first = ws.Range('X2').Value
    date_last = ws.Range('X3').Value
    benchmark = ws.Range('X4').Value
    amount = ws.Range('N3').Value
    if flag_MFO == 1:
        taxa_gestao = ws.Range('N4').Value
    
    date_first = pd.Timestamp(date_first.timestamp(), unit = 's')
    date_last = pd.Timestamp(date_last.timestamp(), unit = 's')
    
    ExcRng = ws.UsedRange()
    raw_data = ExcRng[7:]
    portfolio = pd.DataFrame(data = raw_data[0:]).dropna(how='all',axis=0).dropna(how='all',axis=1)
    
    '''
    #Using openpyxl:
    excel_path = parent_path + '/Carteira Recomendada - MFO 3.xlsm'
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    worksheet = wb[sheet_portfolio]
    
    # Find out how to get these ones using win32com
    date_first = worksheet['AB2'].value
    date_last = worksheet['AB3'].value
    benchmark = worksheet['AB4'].value
    amount = worksheet['M3'].value
    taxa_gestao = worksheet['J3'].value
    
    portfolio = pd.read_excel(excel_path, sheet_name = sheet_portfolio, header=1).iloc[4:,2:].dropna(how='all',axis=0).dropna(how='all',axis=1) 
    portfolio = portfolio.dropna(how='all',axis=0).iloc[:-1,:] # Delete TOTAL line and beyond
    ''' 
    
    portfolio = portfolio.rename(columns=portfolio.iloc[0]).iloc[1:,1:]
    portfolio = portfolio.replace('', np.nan)
    portfolio = portfolio.dropna(how='all',subset = ['Ativo'], axis=0)
    portfolio.reset_index(inplace=True,drop=True)
    portfolio['R$ / $'] = portfolio['R$ / $'].astype(float)
    portfolio['% do PL'] = portfolio['% do PL'].astype(float)
    portfolio['Liquidez/ Duration'] = portfolio['Liquidez/ Duration'].replace(["-"], np.nan).astype(float)
    
    columns = list(portfolio.columns)[2:]
    columns = ["column_0"] + ["column_1"] + columns
    portfolio.columns = columns
    
    portfolio = portfolio.dropna(how='all',axis=0).iloc[:-1,:] # Delete TOTAL line

    classe = ""
    estrategia = ""
    for i in range(portfolio.shape[0]):
        if portfolio.iloc[i,1] == "y":
            classe = portfolio.iloc[i,2]
        if (portfolio.iloc[i,0] == "y" or portfolio.iloc[i,0] == "x") and portfolio.iloc[i,1] == "x":
            estrategia = portfolio.iloc[i,2]
        else:
            portfolio._set_value(i,'Classe', classe)
            portfolio._set_value(i,'Estratégia', estrategia)
            
    # Adjust strategy names
    portfolio.loc[((portfolio['Classe']=='Renda Variável') & (portfolio['Estratégia'] == "Brasil")), 'Estratégia'] = 'RV Brasil'
    portfolio.loc[((portfolio['Classe']=='Renda Variável') & (portfolio['Estratégia'] == "Internacional")), 'Estratégia'] = 'RV Internacional'
    portfolio.loc[((portfolio['Classe']=='Renda Fixa') & (portfolio['Estratégia'] == "Internacional")), 'Estratégia'] = 'RF Internacional'
    
    
    portfolio = portfolio[(portfolio['column_0'] != "x") & (portfolio['column_1'] != "x") & (portfolio['column_1'] != "y")]    
               
    portfolio = portfolio[['Ativo','CNPJ', 'Veículo', 'R$ / $', '% do PL', 'Benchmark', '% Benchmark', 'Benchmark +', 'Liquidez/ Duration', 'Classe', 'Estratégia']]
    
    portfolio_ativos = [sub.replace(' FIC', '').replace(' de ', ' ').replace(' LP', '').replace(' MM', '').replace(' Access', '')
                              .replace(' Feeder', '').replace(' FIA', '').replace(' FIM', '').replace(' FIM', '').replace(' RF', '')
                              .replace(' FIRF', '').replace(' CP', '').replace(' IE', '').replace(' Long And Short', ' LS')
                              .replace(' Renda Fixa', ' RF').replace('Ibovespa', 'IBOV')
                              .replace(' Debêntures Incentivadas', ' Deb Inc') for sub in portfolio['Ativo']]
    portfolio['Ativo'] = portfolio_ativos
    
    dict_CNPJ = dict(zip(portfolio['CNPJ'], portfolio['Ativo']))
    
    resposta = input("Gostaria de utilizar marcação a mercado para títulos públicos? (S/N): ")
    if resposta.lower() == 's':
        flag_fixedIncome_MtM = 1
    elif resposta.lower() == 'n':
        flag_fixedIncome_MtM = 0
    else:
        print("Resposta inválida. Por favor, responda com 'S' para Sim ou 'N' para Não.")
        

    ID_ativos = portfolio[((portfolio['CNPJ']!="-"))]['CNPJ']
    cnpj_list = []
    fixedIncome_list = []
    stocks_list = []

    for id in ID_ativos:
        if isinstance(id, (int, float)):  # Check if the ID is a cnpj (fund)
            cnpj_list.append(id)
        elif isinstance(id, str) and id.startswith("BRSTINCNTB"): #  Check if ID is ISIN of a NTN-B
            fixedIncome_list.append(id)
        elif id in portfolio['Ativo'].values: # Check if the ID is of a stock or listed fund
            stocks_list.append(id)

    #cnpj_list = cnpj_list.to_numpy().flatten().astype(float)

    benchmark_list = portfolio['Benchmark'].dropna().unique().tolist()
    for indice in ['CDI', 'IPCA', 'Prévia IPCA','SP500', 'SELIC', 'IBOV']:
        if indice not in benchmark_list:
            benchmark_list.append(indice)
    
    additional_benchmarks = portfolio['Benchmark'].dropna().unique().tolist() # Add values from portfolio['Benchmark']

    # Extend benchmark_list with additional_benchmarks
    benchmark_list.extend(additional_benchmarks)

    # Get unique values in benchmark_list
    benchmark_list = list(set(benchmark_list))
    print("Done.")
    
    ''' 3) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''
    
    print("Getting fund prices...")
    fund_prices = fund_prices_database(cnpj_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5))
    fund_prices.rename(columns=dict_CNPJ, inplace=True)
    new_columns = list(portfolio[((portfolio['CNPJ'].isin(cnpj_list)) & (~portfolio['Ativo'].isin(list(fund_prices.columns))))]['Ativo'])
    fund_prices[[new_columns]] = np.nan
    
    # Adjusting repeated assets in prices dataframe
    fund_list = portfolio[(portfolio['CNPJ']!="-")]['Ativo']
    fund_list1 = []
    fund_list2 = []
    for x in fund_list:
        if x in fund_list1:
            fund_list2.append(x)
        if x not in fund_list:
            fund_list1.append(x)
    fund_prices2 = fund_prices.loc[:,fund_prices.columns.isin(fund_list2)]
    fund_list2 = list(fund_prices2.columns)
    fund_list2 = [x + " " for x in fund_list2]
    fund_prices2.columns = fund_list2
    
    fund_prices = pd.concat([fund_prices2, fund_prices], axis = 1)
    
    # Adjusting repeated assets in portfolio dataframe
    portfolio_list = portfolio['Ativo']
    portfolio_list1 = []
    portfolio_list2 = []
    for x in portfolio_list:
        if x in portfolio_list1:
            portfolio_list2.append(x)
        if x not in portfolio_list1:
            portfolio_list1.append(x)
            
    portfolio2 = portfolio.loc[portfolio['Ativo'].isin(portfolio_list2), :]
    portfolio2 = portfolio2.drop_duplicates(subset=["Ativo"], keep='last', inplace=False)
    portfolio1 = portfolio.drop_duplicates(subset=["Ativo"], keep='first', inplace=False)
    
    portfolio_list2 = list(portfolio2['Ativo'])
    portfolio_list2 = [x + " " for x in portfolio_list2]
    portfolio2['Ativo'] = portfolio_list2
    
    portfolio = pd.concat([portfolio2, portfolio1], axis = 0)
    
    print("Done.")  
    
    if flag_fixedIncome_MtM != 0 and len(fixedIncome_list)>0:
        print("NTN-B prices...")
        fixedIncome_prices = fixed_income_prices_database(fixedIncome_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5))
        fixedIncome_prices.rename(columns=dict_CNPJ, inplace=True)
        new_columns = list(portfolio[((portfolio['CNPJ'].isin(fixedIncome_list)) & (~portfolio['Ativo'].isin(list(fixedIncome_prices.columns))))]['Ativo'])
        fixedIncome_prices[[new_columns]] = np.nan 
        print("Done.")
    
    if len(stocks_list)>0:
        print("Getting stock prices...")
        stock_prices = stock_prices_database(stocks_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5)) 
        new_columns = list(portfolio[((portfolio['CNPJ'].isin(stocks_list)) & (~portfolio['Ativo'].isin(list(stock_prices.columns))))]['Ativo'])
        stock_prices[[new_columns]] = np.nan 
        print("Done.")
    
    print("Getting benchmark prices...")
    benchmark_prices = benchmark_prices_database(benchmark_list, date_first - dt.timedelta(days=62), date_last + dt.timedelta(days=5)) 
    print("Done.")
    
    ''' 4) MANIPULATE PORTFOLIO CATHEGORICAL COLUMNS ----------------------------------------------------------------------------------------------------------'''
    
    portfolio['Benchmark +'] = portfolio['Benchmark +'].fillna(0)
    portfolio['% Benchmark'] = portfolio['% Benchmark'].fillna(0)
    portfolio['Benchmark'] = portfolio['Benchmark'].fillna("-")
    portfolio['Benchmark'].replace("PRÉ", "-", inplace=True)
    
    ''' Below section not needed anymore, as fixed income already comes with benchmark inputed with gross-up rates:
    if portfolio[((portfolio['CNPJ']=="-"))].shape[0] - len(stocks_list) > 0:
        print("Getting fixed income rates...")
        # FI Rates end benchmark
        portfolio.loc[((portfolio['Ativo'].str.contains('CDI', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'CDI'
        portfolio.loc[((portfolio['Ativo'].str.contains('SELIC', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'SELIC'
        portfolio.loc[((portfolio['Ativo'].str.contains('IPCA', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'IPCA'
        portfolio.loc[((portfolio['Ativo'].str.contains('IMA-B', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'IMA-B'
        portfolio.loc[((portfolio['Ativo'].str.contains('USD', regex=True)) & (portfolio['CNPJ'] == "-")), 'Benchmark'] = 'PTAX'
        
        portfolio['Rates'] = portfolio['Ativo'].str.split(r"(",expand=True)[1].str.replace("CDI","").str.replace("SELIC","").str.replace("IPCA","").str.replace("IMA-B","").str.replace("IMAB","").str.replace("USD","").str.replace(" ","").str.replace(")","").str.replace("+","")
        portfolio.loc[((portfolio['Ativo'].str.contains(' CDI', regex=True) | portfolio['Ativo'].str.contains(' SELIC', regex=True)) & (portfolio['CNPJ'] == "-")),'% Benchmark'] = portfolio[((portfolio['Ativo'].str.contains(' CDI', regex=True) | portfolio['Ativo'].str.contains(' SELIC', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']
        portfolio.loc[((portfolio['CNPJ'] == "-") & portfolio['% Benchmark'] == 0), 'Benchmark +'] = portfolio[((~portfolio['Ativo'].str.contains(' CDI', regex=True)) & (portfolio['CNPJ'] == "-"))]['Rates']
        
        portfolio = portfolio.drop(['Rates'], axis = 1)
        
        portfolio['Benchmark +'] = portfolio['Benchmark +'].fillna(0)
        portfolio['% Benchmark'] = portfolio['% Benchmark'].fillna(0)
        portfolio['Benchmark'] = portfolio['Benchmark'].fillna("-")
        
        portfolio['Benchmark +'] = portfolio['Benchmark +'].astype('str').str.rstrip('%').astype('float')
        portfolio['% Benchmark'] = portfolio['% Benchmark'].astype('str').str.rstrip('%').astype('float')
                
        portfolio.loc[(portfolio['CNPJ'] == "-"),'Benchmark +'] = portfolio.loc[(portfolio['CNPJ'] == "-"),'Benchmark +'] / 100
        portfolio.loc[(portfolio['CNPJ'] == "-"),'% Benchmark'] = portfolio.loc[(portfolio['CNPJ'] == "-"),'% Benchmark'] / 100
        
    '''   
    print("Done.")

    ''' 5) GET ASSETS AND BENCHMARKS DAILY RETURNS --------------------------------------------------------------------------------------------------------------------------------'''
    
    print("Calculating fund returns...")
    # Fund daily returns:
    fund_prices.fillna(method='ffill', inplace=True)
    fund_Returns = fund_prices.astype('float') / fund_prices.astype('float').shift(1) - 1
    fund_Returns.iloc[0:1,:].fillna(0, inplace=True)
    fund_Returns = fund_Returns[((fund_Returns.index>=date_first) & (fund_Returns.index<=date_last))]
    print(fund_prices)
    print("Done.")
    
    if flag_fixedIncome_MtM == 1:
        print("Calculating fixed income returns...")
        # Fixed income daily returns:
        fixedIncome_prices.fillna(method='ffill', inplace=True)
        fixedIncome_Returns = fixedIncome_prices.astype('float') / fixedIncome_prices.astype('float').shift(1) - 1
        fixedIncome_Returns.iloc[0:1,:].fillna(0, inplace=True)
        fixedIncome_Returns = fixedIncome_Returns[((fixedIncome_Returns.index>=date_first) & (fixedIncome_Returns.index<=date_last))]
        print("Done.")
    
    if len(stocks_list) > 0:
        print("Calculating stock returns...")
        # Stock daily returns:
        stock_prices.fillna(method='ffill', inplace=True)
        stock_Returns = stock_prices.astype('float') / stock_prices.astype('float').shift(1) - 1
        stock_Returns.iloc[0:1,:].fillna(0, inplace=True)
        stock_Returns = stock_Returns[((stock_Returns.index>=date_first) & (stock_Returns.index<=date_last))]
        print("Done.")
    
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
    benchmark_Returns = benchmark_prices.astype('float') / benchmark_prices.astype('float').shift(1) - 1
    benchmark_Returns = benchmark_Returns.drop(['DU'], axis = 1)
    
    # CDI, SELIC and Previa IPCA daily returns (given as rates, not prices)
    benchmark_Returns.loc[:,'CDI'] = (1+benchmark_prices.loc[:,'CDI'].astype('float')/100)**(1/252)-1
    benchmark_Returns.loc[:,'SELIC'] = (1+benchmark_prices.loc[:,'SELIC'].astype('float')/100)**(1/252)-1
    benchmark_Returns.loc[:,'Prévia IPCA'] = (1+benchmark_prices.loc[:,'Prévia IPCA'].astype('float')/100)**(1/benchmark_prices.loc[:, "DU"])-1
    
    # IPCA
    benchmark_Returns.loc[:, "IPCA"] = (1+benchmark_Returns.loc[:,'IPCA'])**(1/benchmark_prices.loc[:, "DU"])-1 # calculate daily IPCA rates
    benchmark_Returns.loc[:, "IPCA"] = benchmark_Returns.loc[:, "IPCA"].replace(0,np.nan)
    
    # Use forecasted IPCA as missing IPCA
    benchmark_Returns['IPCA_m/yyyy'] = [str(i) + "/"+ str(j) for i, j in zip(list(benchmark_Returns.index.month), list(benchmark_Returns.index.year))] #get new date column
    IPCA_est  = benchmark_Returns.loc[:,['Prévia IPCA', 'IPCA_m/yyyy']] # Create dataframe to get last Prévia IPCA
    last_month = list(benchmark_Returns.loc[:, "IPCA"].dropna().index)[-1]
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
    benchmark_index = list(benchmark_Returns.index)
    benchmark_Returns = pd.merge(benchmark_Returns, IPCA_est, how="left", on=['IPCA_m/yyyy'])
    benchmark_Returns.index = list(benchmark_index)
    benchmark_Returns.loc[~benchmark_Returns['Prévia IPCA_y'].isna(),'IPCA'] = benchmark_Returns['Prévia IPCA_y'] # Fill missing IPCA with forecasted value
    
    benchmark_Returns.fillna(method='ffill', inplace=True)
    benchmark_Returns = benchmark_Returns.drop(columns = ['IPCA_m/yyyy', 'Prévia IPCA_x', 'Prévia IPCA_y'])
    
    benchmark_Returns = benchmark_Returns[((benchmark_Returns.index>=date_first) & (benchmark_Returns.index<=fund_Returns.index[-1]))]
    benchmark_Returns.iloc[0,:] = 0
    print("Done.")
    
    
    ''' 6) CALCULATE PORTFOLIO RETURNS --------------------------------------------------------------------------------------------------------------------------------'''
    
    portfolio = portfolio.set_index('Ativo')
    assets_returns = pd.DataFrame(index = benchmark_Returns.index, columns = list(portfolio.index))
    print("Calculating portfolio returns...")
    
    # Get asset returns
    portfolio['% Benchmark'] = pd.to_numeric(portfolio['% Benchmark'], errors='coerce')
    portfolio['Benchmark +'] = pd.to_numeric(portfolio['Benchmark +'], errors='coerce')
    

    for asset in assets_returns.columns:
    # Check if the asset is in fund_Returns
        if asset in fund_Returns.columns:
            assets_returns[asset] = fund_Returns.loc[assets_returns.index, asset]

        elif len(stocks_list)>0:
            if asset in stock_Returns.columns: # Fixed income prices if MtM is being used
                assets_returns[asset] = stock_Returns.loc[assets_returns.index, asset]
        
        elif flag_fixedIncome_MtM == 1 and len(fixedIncome_list)>0:
            if asset in fixedIncome_Returns.columns: # Fixed income prices if MtM is being used
                assets_returns[asset] = fixedIncome_Returns.loc[assets_returns.index, asset]
        
        #Fill benchmark returns for assets not included above or to those dates in which they fund have no prices
        if portfolio.loc[asset, '% Benchmark'] != 0: 
            if portfolio.loc[asset, 'Benchmark +'] != 0: # Fund with % Benchmark proxy and Benchmark + proxy
                assets_returns.loc[assets_returns[asset].isna(), asset] = (1+benchmark_Returns.loc[assets_returns[asset].isna(), portfolio.loc[asset,'Benchmark']] * portfolio.loc[asset,'% Benchmark']) * ((1+portfolio.loc[asset,'Benchmark +'])**(1/252)) - 1
            else: #Fund with % Benchmark proxy
                assets_returns.loc[assets_returns[asset].isna(), asset] = benchmark_Returns.loc[assets_returns[asset].isna(), portfolio.loc[asset,'Benchmark']] * portfolio.loc[asset,'% Benchmark']
            
        elif portfolio.loc[asset, 'Benchmark +'] != 0: # Fund with Benchmark + proxy
            if portfolio.loc[asset,'Benchmark'] != "-" and portfolio.loc[asset,'Benchmark'] != 0: # Fund with Benchmark+ proxy
                assets_returns.loc[assets_returns[asset].isna(), asset] = (1+benchmark_Returns.loc[assets_returns[asset].isna(), portfolio.loc[asset,'Benchmark']]) * ((1+portfolio.loc[asset,'Benchmark +'])**(1/252)) - 1
            else: # Fund with prefixed proxy
                assets_returns.loc[assets_returns[asset].isna(), asset] = (1+portfolio.loc[asset,'Benchmark +'])**(1/252) - 1


    # Get weighted performance for each asset
    assets_returns_W = assets_returns.copy()
    for i in assets_returns_W.index:
        assets_returns_W.loc[i,:] = assets_returns_W.loc[i,:].mul(portfolio['% do PL']) # performance attribution per product
    
    # PORTFOLIO
    portfolio_return = assets_returns_W.sum(axis=1) # Daily returns
    if flag_MFO == 1:
        portfolio_return = portfolio_return + ((1-taxa_gestao)**(1/252)-1) # Subtract management fee
    
    portfolio_acc = portfolio_return.copy()
    portfolio_acc.name = 'Portfólio Modelo'
    portfolio_acc = pd.concat([portfolio_acc, benchmark_Returns[benchmark]], axis=1)
    for i in range(portfolio_acc.shape[0]-1):
        portfolio_acc.iloc[i+1,:] = (1 + portfolio_acc.iloc[i+1,:]) * (1+ portfolio_acc.iloc[i,:])  - 1 # Accumulated returns
    
    if benchmark == 'CDI': 
        retorno_cdi_acc = portfolio_acc.loc[:,benchmark]
    else:
        retorno_cdi_acc = benchmark_Returns["CDI"].copy()
        for i in range(len(retorno_cdi_acc)-1):
            retorno_cdi_acc.iloc[i+1] = (1 + retorno_cdi_acc.iloc[i+1]) * (1 + retorno_cdi_acc.iloc[i]) - 1 # Accumulated returns
        
    # ASSET STRATEGIES
    
    strategy_list = list(portfolio['Estratégia'].unique())
    strategy_weights = pd.Series(data = strategy_list, index = strategy_list)
    strategy_weights = strategy_weights.apply(lambda x: portfolio[(portfolio['Estratégia'] == x)]['% do PL'].sum()).astype(float)
    
    strategy_listAll = ['Pós-fixado', 'Pré-fixado', 'Inflação', 'RF Internacional', 'Macro', 'Descorrelacionados', 'RV Brasil', 'RV Internacional']
    strategy_weightsAll = pd.Series(data = strategy_listAll, index = strategy_listAll)
    strategy_weightsAll = strategy_weightsAll.apply(lambda x: portfolio[(portfolio['Estratégia'] == x)]['% do PL'].sum()).astype(float)
        
    strategy_returns =  pd.DataFrame(index = benchmark_Returns.index , columns = strategy_listAll) # Daily returns:   
    for i in strategy_returns.columns:
        asset_group = []
        for j in assets_returns_W.columns:
            if portfolio.loc[j,'Estratégia'] == i:
                asset_group = asset_group + [j]        
        strategy_returns[i] = assets_returns_W[asset_group].sum(axis=1)  # Daily returns
        
    strategy_attr = strategy_returns.copy() # Attribution:
    
    strategy_acc = strategy_returns.copy() # Accumulated returns (normalized, standalone):
    strategy_acc = strategy_acc.mul(strategy_weightsAll**(-1)) 
    for i in range(strategy_acc.shape[0]-1):
        strategy_acc.iloc[i+1,:] = (1 + strategy_acc.iloc[i+1,:]) * (1 + strategy_acc.iloc[i,:]) - 1 # Accumulated returns (normalized for each strategy)
        strategy_attr.iloc[i+1,:] =  strategy_returns.iloc[i+1,:] * (1 + portfolio_acc.iloc[i,0])
    
    
    strategy_columns = [sub.replace('Pós-fixado', 'RF Pós').replace('Pré-fixado', 'RF Pré').replace('Inflação', 'RF Inflação').replace('RF Internacional', 'RF Intl.')
                                .replace('Macro', 'MM Macro').replace('Descorrelacionados', 'MM Descorr.').replace('RV Brasil', 'RV BR')
                                .replace('RV Internacional', 'RV Intl.') for sub in strategy_listAll]      
    
    strategy_acc.columns = strategy_columns
    strategy_attr.columns = strategy_columns
    strategy_acc['CDI'] = retorno_cdi_acc
    
    strategy_weightsAll.index = strategy_columns
    
    strategy_weights.index = [sub.replace('Pós-fixado', 'RF Pós').replace('Pré-fixado', 'RF Pré').replace('Inflação', 'RF Inflação')
                                .replace('Macro', 'MM Macro').replace('Descorrelacionados', 'MM Descorrelacionados') for sub in strategy_list]  
        
    # ASSET CLASSES
    class_list = list(portfolio['Classe'].unique())
    class_weights = pd.Series(data = class_list, index = class_list)
    class_weights = class_weights.apply(lambda x: portfolio[(portfolio['Classe'] == x)]['% do PL'].sum()).astype(float)
    
    class_listAll = ['Renda Fixa', 'Multimercado', 'Renda Variável']  
    class_weightsAll = pd.Series(data = class_listAll, index = class_listAll)
    class_weightsAll = class_weightsAll.apply(lambda x: portfolio[(portfolio['Classe'] == x)]['% do PL'].sum()).astype(float)
    
    class_returns = pd.DataFrame(index = benchmark_Returns.index , columns = class_listAll)  # Daily returns:         
    for i in class_returns.columns:
        asset_group = []
        for j in assets_returns_W.columns:
            if portfolio.loc[j,'Classe'] == i:
                asset_group = asset_group + [j]
        class_returns[i] = assets_returns_W[asset_group].sum(axis=1)    # Daily returns
            
    class_attr = class_returns.copy() # Attribution:
        
    class_acc = class_returns.copy() # Accumulated returns (normalized, standalone):
    class_acc = class_acc.mul(class_weightsAll**(-1)) 
    for i in range(class_acc.shape[0]-1):
        class_acc.iloc[i+1,:] = (1 + class_acc.iloc[i+1,:]) * (1 + class_acc.iloc[i,:]) - 1 # Accumulated returns (normalized for each strategy)
        class_attr.iloc[i+1,:] =  class_returns.iloc[i+1,:] * (1 + portfolio_acc.iloc[i,0])
    
    
    class_columns = [sub.replace('Renda Fixa', 'Renda Fixa (RF)').replace('Multimercado', 'Multimercado (MM)')
                     .replace('Renda Variável', 'Renda Variável (RV)') for sub in class_listAll]
        
    class_acc.columns = class_columns
    class_attr.columns = class_columns
    class_acc['CDI'] = retorno_cdi_acc
    
    class_weightsAll.index = class_columns
    
    class_weights.index = [sub.replace('Renda Fixa', 'Renda Fixa (RF)').replace('Multimercado', 'Multimercado (MM)')
                     .replace('Renda Variável', 'Renda Variável (RV)') for sub in class_list]
    
    # VEHICLE
    vehicle = pd.Series(data = list(portfolio['Veículo'].unique()), index = list(portfolio['Veículo'].unique()))
    vehicle = vehicle.apply(lambda x: portfolio[(portfolio['Veículo'] == x)]['% do PL'].sum()).astype(float)
    
    
    vehicle_list = [sub.replace('F. Excl.', 'Fundo Exclusivo').replace('C. Adm.', 'Carteira Administrada')
                     .replace('Offsh.', 'Offshore') for sub in list(vehicle.index)] 
    vehicle.index = vehicle_list    
        
    print("Done.")
    
    ''' 7) CALCULATE PERFORMANCE METRICS --------------------------------------------------------------------------------------------------------------------------------'''
    print("Calculating performance metrics...")
    
    # Performance attribution
    class_perf_attr = class_attr.sum(axis=0)
    strategy_perf_attr = strategy_attr.sum(axis=0)
    if flag_MFO == 1:
        class_perf_attr['Taxa de Gestão'] = (1-taxa_gestao)**(class_attr.shape[0]/252)-1
        strategy_perf_attr['Taxa de Gestão'] = (1-taxa_gestao)**(strategy_attr.shape[0]/252)-1
    
    class_perf_attr['Total'] = class_perf_attr.sum()
    
    strategy_perf_attr['Total'] = strategy_perf_attr.sum()
    
    
    # Portfolio vs. Benchmark: Return, Vol, Sharpe
    if benchmark == "": benchmark = "CDI" 
    
    portf_vs_bench_1 = pd.DataFrame(columns = ["Rentabilidade Acumulada", "Rentabilidade Anualizada", "Volatilidade Anualizada", "Sharpe"],
                                      index = ["Portfólio Modelo", benchmark])
    
    portf_vs_bench_1.iloc[0,0] = portfolio_acc.iloc[-1,0]
    portf_vs_bench_1.iloc[1,0] = portfolio_acc.iloc[-1,1]
    portf_vs_bench_1.iloc[0,1] = (1+portfolio_acc.iloc[-1,0])**(252/portfolio_acc.shape[0])-1
    portf_vs_bench_1.iloc[1,1] = (1+portfolio_acc.iloc[-1,1])**(252/portfolio_acc.shape[0])-1
    portf_vs_bench_1.iloc[0,2] = np.sqrt(252)*np.std(portfolio_return)
    portf_vs_bench_1.iloc[1,2] = np.sqrt(252)*np.std(benchmark_Returns[benchmark])
    portf_vs_bench_1.iloc[0,3] = (portf_vs_bench_1.iloc[0,1] - (((benchmark_Returns['CDI']+1).to_numpy().prod())**(252/benchmark_Returns.shape[0])-1)) / portf_vs_bench_1.iloc[0,2]
    portf_vs_bench_1.iloc[1,3] = (portf_vs_bench_1.iloc[1,1] - (((benchmark_Returns['CDI']+1).to_numpy().prod())**(252/benchmark_Returns.shape[0])-1)) / portf_vs_bench_1.iloc[1,2]
    
    # Portfolio vs. Benchmark: Returns
    portf_vs_bench_2 = pd.DataFrame(columns = ["Mes", "Ano", "6 meses", "12 meses", "2 anos", "Saldo Inicial ("+date_first.strftime("%d/%m/%Y")+")", "Saldo Final ("+portfolio_acc.index[-1].strftime("%d/%m/%Y")+")"],
                                      index = ["Portfólio Modelo", benchmark])
    
    date_MTD = dt.date(portfolio_return.index[-1].year, portfolio_return.index[-1].month, 1)
    date_YTD = dt.date(portfolio_return.index[-1].year, 1, 1)
    date_6M = portfolio_return.index[-1] - relativedelta(months=6)
    date_12M = portfolio_return.index[-1] - relativedelta(months=12)
    date_24M = portfolio_return.index[-1] - relativedelta(months=24)
    
    portf_vs_bench_2.iloc[0,0] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_MTD)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1,0] = (1+benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_MTD)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0,1] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_YTD)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1,1] = (1+benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >date_YTD)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0,2] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_6M)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1,2] = (1+benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_6M)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0,3] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_12M)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1,3] = (1+benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_12M)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0,4] = (1+portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_24M)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1,4] = (1+benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_24M)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0,5] = amount
    portf_vs_bench_2.iloc[1,5] = amount
    portf_vs_bench_2.iloc[0,6] = amount * (1+portfolio_acc.iloc[-1,0])
    portf_vs_bench_2.iloc[1,6] = amount * (1+portfolio_acc.iloc[-1,1])
    
    
    # Portfolio vs. Benchmark: Monthly results
    portf_vs_bench_aux = portfolio_return.copy()
    portf_vs_bench_aux.name = "Portfólio Modelo"
    portf_vs_bench_aux = pd.concat([portf_vs_bench_aux, benchmark_Returns[benchmark]], axis=1)
    
    portf_vs_bench_aux.index = pd.to_datetime(portf_vs_bench_aux.index)
    portf_vs_bench_3 = portf_vs_bench_aux.groupby(pd.Grouper(freq='M')).apply(lambda x: (1 + x).prod() - 1)
    portf_vs_bench_3.index = portf_vs_bench_3.index.strftime("%Y-%m")
        
    portf_vs_bench_3['%'+benchmark] = portf_vs_bench_3.apply(lambda row: '-' if row[benchmark] < 0 else row["Portfólio Modelo"] / row[benchmark], axis=1)

    
    portf_vs_bench_3.index =  [sub.replace('-01', 'Jan.').replace('-02', 'Fev.').replace('-03', 'Mar.').replace('-04', 'Abr.').replace('-05', 'Mai.')
                                .replace('-06', 'Jun.').replace('-07', 'Jul.').replace('-08', 'Ago.').replace('-09', 'Set.').replace('-10', 'Out.')
                               .replace('-11', 'Nov.').replace('-12', 'Dez.') for sub in list(portf_vs_bench_3.index)]
    
    
    # Portfolio vs. Benchmark: Statistics
    portf_vs_bench_4 = pd.DataFrame(columns = ["Meses\nPositivos", "Meses\nNegativos", "Maior Retorno\nMensal", "Menor Retorno\nMensal", "Acima do CDI\n(meses)", "Abaixo do CDI\n(meses)"],
                                      index = ["Portfólio Modelo", benchmark])
    
    portf_vs_bench_4.iloc[0,0] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]>=0)]["Portfólio Modelo"].count()
    portf_vs_bench_4.iloc[1,0] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]>=0)][benchmark].count()
    portf_vs_bench_4.iloc[0,1] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]<0)]["Portfólio Modelo"].count()
    portf_vs_bench_4.iloc[1,1] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]<0)][benchmark].count()
    portf_vs_bench_4.iloc[0,2] = portf_vs_bench_3["Portfólio Modelo"].max()
    portf_vs_bench_4.iloc[1,2] = portf_vs_bench_3[benchmark].max()
    portf_vs_bench_4.iloc[0,3] = portf_vs_bench_3["Portfólio Modelo"].min()
    portf_vs_bench_4.iloc[1,3] = portf_vs_bench_3[benchmark].min()
    
    if benchmark == "CDI":
        portf_vs_bench_4.iloc[0,4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]>=portf_vs_bench_3[benchmark])]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[0,5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]<portf_vs_bench_3[benchmark])]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1,4] = 0
        portf_vs_bench_4.iloc[1,5] = 0
    else:
        retorno_CDI_M_aux = benchmark_Returns["CDI"].copy()
        retorno_CDI_M_aux.name = "retorno_CDI_M" 
        retorno_CDI_M_aux.index = pd.to_datetime(portf_vs_bench_aux.index)
        retorno_CDI_M = retorno_CDI_M_aux.groupby(pd.Grouper(freq='M')).apply(lambda x: (1 + x).prod() - 1)
        retorno_CDI_M.index = retorno_CDI_M.index.strftime("%Y-%m")
      
        retorno_CDI_M.index = [sub.replace('-01', 'Jan.').replace('-02', 'Fev.').replace('-03', 'Mar.').replace('-04', 'Abr.').replace('-05', 'Mai.')
                                    .replace('-06', 'Jun.').replace('-07', 'Jul.').replace('-08', 'Ago.').replace('-09', 'Set.').replace('-10', 'Out.')
                                   .replace('-11', 'Nov.').replace('-12', 'Dez.') for sub in list(retorno_CDI_M.index)]
        
        #retorno_CDI_M = pd.Series(index = np.unique(list(retorno_CDI_M.index)))
        
        portf_vs_bench_4.iloc[0,4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]>=retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[0,5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"]<retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1,4] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]>=retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1,5] = portf_vs_bench_3[(portf_vs_bench_3[benchmark]<retorno_CDI_M)]["Portfólio Modelo"].count()
    
    
    # Portfolio vs. Benchmark: Volatility
    portf_vs_bench_5 = pd.DataFrame(columns = ["Mes", "Ano", "6 meses", "12 meses", "2 anos"],
                                      index = ["Portfólio Modelo", benchmark])
        
    portf_vs_bench_5.iloc[0,0] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_MTD)])
    portf_vs_bench_5.iloc[1,0] = np.sqrt(252)*np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_MTD)][benchmark])
    portf_vs_bench_5.iloc[0,1] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_YTD)])
    portf_vs_bench_5.iloc[1,1] = np.sqrt(252)*np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_YTD)][benchmark])
    portf_vs_bench_5.iloc[0,2] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_6M)])
    portf_vs_bench_5.iloc[1,2] = np.sqrt(252)*np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_6M)][benchmark])
    portf_vs_bench_5.iloc[0,3] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_12M)])
    portf_vs_bench_5.iloc[1,3] = np.sqrt(252)*np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_12M)][benchmark])
    portf_vs_bench_5.iloc[0,4] = np.sqrt(252)*np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_24M)])
    portf_vs_bench_5.iloc[1,4] = np.sqrt(252)*np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_24M)][benchmark])
    
    volatility = pd.concat([portfolio_return, benchmark_Returns[benchmark]], axis=1)
    volatility.columns = ["Portfólio Modelo", benchmark]
    
    for i in range(volatility.shape[0]-21): # Rolling 21-days volatility
        volatility.iloc[i+21,0] = np.sqrt(252)*np.std(portfolio_return[i:i+21]) 
        volatility.iloc[i+21,1] = np.sqrt(252)*np.std(benchmark_Returns[benchmark][i:i+21])
    
    volatility = volatility.iloc[21:,:]
    
    # Portfolio vs. Benchmark: Drawdown
    portf_vs_bench_6 = pd.DataFrame(columns = ["Mes", "Ano", "6 meses", "12 meses", "2 anos", "Drawdown Máximo", "Data", "Tempo de Recuperação"],
                                      index = ["Portfólio Modelo", benchmark])
    
    drawdown = pd.DataFrame(np.zeros((portfolio_return.shape[0], 2)), columns = ["Portfólio Modelo", benchmark], index = portfolio_return.index)
    
    drawdown_acc_returns = portfolio_acc.copy()
    
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
    
    portf_vs_bench_6.iloc[0,0] = min(drawdown_MTD.iloc[:,0])
    portf_vs_bench_6.iloc[1,0] = min(drawdown_MTD.iloc[:,1])
    portf_vs_bench_6.iloc[0,1] = min(drawdown_YTD.iloc[:,0])
    portf_vs_bench_6.iloc[1,1] = min(drawdown_YTD.iloc[:,1])
    portf_vs_bench_6.iloc[0,2] = min(drawdown_6M.iloc[:,0])
    portf_vs_bench_6.iloc[1,2] = min(drawdown_6M.iloc[:,1])
    portf_vs_bench_6.iloc[0,3] = min(drawdown_12M.iloc[:,0])
    portf_vs_bench_6.iloc[1,3] = min(drawdown_12M.iloc[:,1])
    portf_vs_bench_6.iloc[0,4] = min(drawdown_24M.iloc[:,0])
    portf_vs_bench_6.iloc[1,4] = min(drawdown_24M.iloc[:,1])
    

    if min(drawdown.iloc[:,0]) == 0: 
        portf_vs_bench_6.iloc[0,5] = 0
        portf_vs_bench_6.iloc[0,6] = "-"
        portf_vs_bench_6.iloc[0,7] = "-"
    
    else: 
        portf_vs_bench_6.iloc[0,5] = min(drawdown.iloc[:,0])
        portf_vs_bench_6.iloc[0,6] = list(drawdown[(drawdown.iloc[:,0] == min(drawdown.iloc[:,0]))].index.strftime('%d/%m/%Y'))[0]
        aux_time_recover1 = drawdown.loc[:drawdown.iloc[:,0].idxmin()].iloc[:,0]
        aux_time_recover1 = aux_time_recover1[(aux_time_recover1.index > max(aux_time_recover1[(aux_time_recover1 == 0)].index))]
        aux_time_recover2 = drawdown.loc[drawdown.iloc[:,0].idxmin():].iloc[:,0]
        
        if len(aux_time_recover2[(aux_time_recover2 == 0)].index == 0) > 0: 
            aux_time_recover2 = aux_time_recover2[(aux_time_recover2.index <= min(aux_time_recover2[(aux_time_recover2 == 0)].index))]
            portf_vs_bench_6.iloc[0,7] =  len(aux_time_recover1) + len(aux_time_recover2) - 1
            
        else:    # has not yet recovered
            time_recover = len(aux_time_recover2)
            portf_vs_bench_6.iloc[0,7] =  str(time_recover) + '(+...) '
        
    if min(drawdown.iloc[:,1]) == 0: 
        portf_vs_bench_6.iloc[1,5] = 0
        portf_vs_bench_6.iloc[1,6] = "-"
        portf_vs_bench_6.iloc[1,7] = "-"
    else: 
       portf_vs_bench_6.iloc[1,5] = min(drawdown.iloc[:,1])
       portf_vs_bench_6.iloc[1,6] = list(drawdown[(drawdown.iloc[:,1] == min(drawdown.iloc[:,1]))].index.strftime('%d/%m/%Y'))[0]
       aux_time_recover1 = drawdown.loc[:drawdown.iloc[:,1].idxmin()].iloc[:,1]
       aux_time_recover1 = aux_time_recover1[(aux_time_recover1.index > max(aux_time_recover1[(aux_time_recover1 == 0)].index))]
       aux_time_recover2 = drawdown.loc[drawdown.iloc[:,1].idxmin():].iloc[:,1]
       
       if len(aux_time_recover2[(aux_time_recover2 == 0)].index == 0) > 0: 
           aux_time_recover2 = aux_time_recover2[(aux_time_recover2.index <= min(aux_time_recover2[(aux_time_recover2 == 0)].index))]
           portf_vs_bench_6.iloc[1,7] =  len(aux_time_recover1) + len(aux_time_recover2) - 1
           
       else:    # has not yet recovered
           time_recover = len(aux_time_recover2)
           portf_vs_bench_6.iloc[1,7] =  str(time_recover) + '(+...) '
            
    
    # Correlation Matrix
    asset_group = [benchmark]
    for j in assets_returns.columns:
        if portfolio.loc[j,'CNPJ']!="-" or np.std(assets_returns[j])*np.sqrt(252)>0.01:
            asset_group = asset_group + [j]
            
    assets_returns2 = pd.concat([assets_returns, benchmark_Returns[benchmark]], axis=1)
    first_column = assets_returns2.pop(benchmark)        
    assets_returns2.insert(0,benchmark, first_column)
    assets_returns2 = assets_returns2[asset_group]
                    
    correlation = assets_returns2[asset_group].corr()
    
    for i in range(correlation.shape[0]):
        for j in range(correlation.shape[1]):
            if i<j: correlation.iloc[i,j] = "-"
            
    rows_correl = list(correlation.index)
    if len(rows_correl)<=26:
        alphabet = list(map(chr, range(65, 91)))[0:len(rows_correl)]
    else:
        alphabet1 = list(map(chr, range(65, 91)))[0:26]
        alphabet2 = list(map(chr, range(65, 91)))[0:len(rows_correl)-26]
        alphabet2 = ["A" + letter for letter in alphabet2]
        alphabet = alphabet1 + alphabet2
    columns_correl = list('('+ a + ") " for a in alphabet)
    rows_correl =  [ x + y for x, y in zip(columns_correl, rows_correl)]
    correlation.columns = columns_correl
    correlation.index = rows_correl
    print("Done.")
        
    ''' 8) PRINT TO EXCEL --------------------------------------------------------------------------------------------------------------------------------'''
    
    # Create workbook object (try to opeen an existing one, if it doesn`t exist, create one)
    try:
        wb = xw.Book(file_name)
    except:
        wb = xw.Book()
        wb.save(file_name)
        wb = xw.Book(file_name)
    
    sNamList = [sh.name for sh in wb.sheets]
    wb.app.calculation = 'manual'
    
    # Write Portfolio vs. Benchmark comparisons:
    output_sheet = 'Performance Measurement'        
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    print("Writing on Excel...(1/15)")
    sheet.range('4:6').clear_contents() # Delete old data
    sheet.range((4, 2)).value = portf_vs_bench_1
    
    print("Writing on Excel...(2/15)")
    sheet.range('11:13').clear_contents()
    sheet.range((11, 2)).value = portf_vs_bench_2
    
    print("Writing on Excel...(3/15)")
    sheet.range('18:21').clear_contents() 
    sheet.range((18, 2)).value = portf_vs_bench_3.T
    
    print("Writing on Excel...(4/15)")
    sheet.range('26:28').clear_contents() 
    sheet.range((26, 2)).value = portf_vs_bench_4
    
    print("Writing on Excel...(5/15)")
    sheet.range('33:35').clear_contents() 
    sheet.range((33, 2)).value = portf_vs_bench_5
    
    print("Writing on Excel...(6/15)")
    sheet.range('40:43').clear_contents()
    sheet.range((40, 2)).value = portf_vs_bench_6
    
    print("Writing on Excel...(7/15)")
    sheet.range((48, 2)).value = class_perf_attr
    
    print("Writing on Excel...(8/15)")
    sheet.range((58, 2)).value = strategy_perf_attr
    
    
    print("Writing on Excel...(9/15)")
    # Write Total per Class, Strategy and Vehicle:
    output_sheet = 'Alocation'        
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.range('5:18').clear_contents() # Delete old data
    sheet.range((5, 2)).value = class_weights
    sheet.range((5, 5)).value = strategy_weightsAll
    sheet.range((5, 8)).value = vehicle
    sheet.range((5, 11)).value = class_weightsAll
    
    print("Writing on Excel...(10/15)")
    # Write Portfolio and Benchmark Accumulated Returns:
    output_sheet = 'AccReturns'        
    
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.clear_contents() # Delete old data
    sheet.range((2, 2)).value = portfolio_acc
    
    print("Writing on Excel...(11/15)")
    # Write Volatility Moving Windows:
    output_sheet = 'MovingVol'        
    data = volatility
    
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.clear_contents() # Delete old data
    sheet.range((2, 2)).value = data
    
    print("Writing on Excel...(12/15)")
    # Write Correlation Matrix:
    output_sheet = 'Sim_4'        
    data = correlation
    
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.range('4:40').clear_contents() 
    sheet.range((6, 2)).value = data
    
    print("Writing on Excel...(13/15)")
    # Write Drawdown:
    output_sheet = 'Drawdown'        
    data = drawdown
    
    if not output_sheet in sNamList:
        sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.clear_contents() # Delete old data
    sheet.range((2, 2)).value = data
    
    print("Writing on Excel...(14/15)")
    # Write Assets Returns:
    output_sheet = 'Assets Returns'        
    data = assets_returns
    
    if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.clear_contents() # Delete old data
    sheet.range((2, 1)).value = data
    
    print("Writing on Excel...(15/15)")
    # Write Strategy Acc Returns:
    output_sheet = 'Strategy AccReturns'        
    data = strategy_acc
    
    if not output_sheet in sNamList: sheet = wb.sheets.add(output_sheet)
    else: sheet = wb.sheets[output_sheet]
    
    sheet.clear_contents() # Delete old data
    sheet.range((2, 1)).value = data
    
    
    wb.app.calculation = 'automatic'
    
    print("Completed.")

main_code()

