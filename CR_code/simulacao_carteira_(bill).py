''' 1) IMPORT MAIN LIBRARIES --------------------------------------------------------------------------------'''
import pandas as pd
import numpy as np
import xlwings as xw
import win32com.client
import datetime as dt
from dateutil.relativedelta import relativedelta
import warnings
import sys
import os
import inspect

warnings.filterwarnings('ignore')

parent_path = os.path.dirname(os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))))
sys.path.append(parent_path + '/CR_code/formulas')

from formulas.fund_prices_database import fund_prices_database
from formulas.benchmark_prices_database import benchmark_prices_database
from formulas.stock_prices_database import stock_prices_database
from formulas.fixed_income_prices_database import fixed_income_prices_database

def ler_tabela_nomeada(caminho_arquivo, nome_tabela):
    try:
        print(f"Tentando abrir o arquivo: {caminho_arquivo}")
        app = xw.App(visible=False)
        wb = xw.Book(caminho_arquivo)
        print("Arquivo aberto com sucesso.")
        
        tabela = wb.sheets['Rebalance'].range(nome_tabela).options(pd.DataFrame, header=1, index=False).value
        print("Tabela 'Table1' carregada com sucesso.")
        
        app.quit()
        return tabela
    except Exception as e:
        print(f"Erro ao carregar a tabela nomeada '{nome_tabela}': {e}")
        app.quit()
        return None

def ajustar_nomes_colunas(df):
    df.columns = [
        'CLASSE', 'ESTRATÉGIA', 'SUB-ESTRATÉGIA', 'ATIVO', 'CNPJ', 'LIQUIDEZ/DURATION', 'INÍCIO DO FUNDO', 'ABERTO P/ CAPTAÇÃO?',
        'PL_ATUAL', '%PL_ATUAL', '%ESTRATÉGIA_ATUAL', 'PL_NOVO', '%PL_NOVO', '%ESTRATÉGIA_NOVO', 'VALOR', 'BENCH', '%BENCH', 'BENCH+'
    ]
    return df

''' 2) READ SELECTED FUNDS ------------------------------------------------------------------------------------'''

def read_selected_funds():
    print("Initiating simulation...") 
    print("Getting portfolio from Excel...")     
    
    file_path = r'Y:\GESTAO\3) Carteira Recomendada\Carteira Recomendada - Bill.xlsm'
    nome_tabela = 'Table1'

    try:
        portfolio = ler_tabela_nomeada(file_path, nome_tabela)
        if portfolio is None or portfolio.empty:
            raise ValueError("Falha ao carregar os dados ou dados estão vazios.")
    except Exception as e:
        print(f"Erro ao ler a tabela: {e}")
        return None, None, None, None, None, None, None, None, None, None, None, None

    print(f"Dimensões dos dados carregados: {portfolio.shape}")
    print("Dados carregados e limpos:")
    print(portfolio.to_string())

    portfolio = ajustar_nomes_colunas(portfolio)

    print(f"Columns available in the table 'Table1':")
    print(list(portfolio.columns))

    # Converting columns to appropriate types
    portfolio['PL_ATUAL'] = portfolio['PL_ATUAL'].astype(float)
    portfolio['%PL_ATUAL'] = portfolio['%PL_ATUAL'].astype(float)
    portfolio['%ESTRATÉGIA_ATUAL'] = portfolio['%ESTRATÉGIA_ATUAL'].astype(float)
    portfolio['PL_NOVO'] = portfolio['PL_NOVO'].astype(float)
    portfolio['%PL_NOVO'] = portfolio['%PL_NOVO'].astype(float)
    portfolio['%ESTRATÉGIA_NOVO'] = portfolio['%ESTRATÉGIA_NOVO'].astype(float)
    portfolio['VALOR'] = portfolio['VALOR'].astype(float)
    portfolio['%BENCH'] = portfolio['%BENCH'].astype(float)
    portfolio['BENCH+'] = portfolio['BENCH+'].astype(float)

    # Filtering out empty rows
    portfolio = portfolio.dropna(how='all', subset=['ATIVO'])

    # Getting other necessary data
    date_first = dt.datetime(2022, 12, 21)
    date_last = dt.datetime(2024, 4, 21)
    benchmark = 'CDI'
    amount = 5000000
    taxa_gestao = 0.00

    # Flags and lists for later processing
    flag_MFO = 1
    flag_fixedIncome_MtM = 0
    cnpj_list = portfolio['CNPJ'].dropna().unique().tolist()
    fixedIncome_list = []
    stocks_list = []
    benchmark_list = ['CDI', 'IPCA', 'SELIC', 'Ibovespa']
    dict_CNPJ = dict(zip(portfolio['CNPJ'], portfolio['ATIVO']))

    # Solicitar a marcação a mercado para títulos públicos (only portuguese comment to avoid confusion, English: 'Request mark-to-market for public securities')
    resposta = input("Gostaria de utilizar marcação a mercado para títulos públicos? (S/N): ")
    if resposta.lower() == 's':
        flag_fixedIncome_MtM = 1
    elif resposta.lower() == 'n':
        flag_fixedIncome_MtM = 0
    else:
        print("Resposta inválida. Por favor, responda com 'S' para Sim ou 'N' para Não.")
        return None, None, None, None, None, None, None, None, None, None, None, None

    return portfolio, date_first, date_last, benchmark, amount, taxa_gestao, flag_MFO, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ

''' 3) IMPORT FUND PRICES --------------------------------------------------------------------------------------'''

def import_fund_prices(portfolio, date_first, date_last, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ):
    if len(cnpj_list) > 0:
        print("Getting fund prices...")
        fund_prices = fund_prices_database(cnpj_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5))
        fund_prices.rename(columns=dict_CNPJ, inplace=True)
        new_columns = list(portfolio[((portfolio['CNPJ'].isin(cnpj_list)) & (~portfolio['ATIVO'].isin(list(fund_prices.columns))))]['ATIVO'])
        fund_prices[new_columns] = np.nan
        
        # Adjusting repeated assets in prices dataframe
        fund_list = portfolio[(portfolio['CNPJ']!="-")]['ATIVO']
        fund_list1 = []
        fund_list2 = []
        for x in fund_list:
            if x in fund_list1:
                fund_list2.append(x)
            if x not in fund_list1:
                fund_list1.append(x)
        fund_prices2 = fund_prices.loc[:,fund_prices.columns.isin(fund_list2)]
        fund_list2 = list(fund_prices2.columns)
        fund_list2 = [x + " " for x in fund_list2]
        fund_prices2.columns = fund_list2
        
        fund_prices = pd.concat([fund_prices2, fund_prices], axis=1)
        
        # Adjusting repeated assets in portfolio dataframe
        portfolio_list = portfolio['ATIVO']
        portfolio_list1 = []
        portfolio_list2 = []
        for x in portfolio_list:
            if x in portfolio_list1:
                portfolio_list2.append(x)
            if x not in portfolio_list1:
                portfolio_list1.append(x)
                
        portfolio2 = portfolio.loc[portfolio['ATIVO'].isin(portfolio_list2), :]
        portfolio2 = portfolio2.drop_duplicates(subset=["ATIVO"], keep='last', inplace=False)
        portfolio1 = portfolio.drop_duplicates(subset=["ATIVO"], keep='first', inplace=False)
        portfolio.drop(columns=['Ativo'], inplace=True, errors='ignore')
        
        portfolio_list2 = list(portfolio2['ATIVO'])
        portfolio_list2 = [x + " " for x in portfolio_list2]
        portfolio2['ATIVO'] = portfolio_list2
        
        portfolio = pd.concat([portfolio2, portfolio1], axis=0)
        
        print("Done.")  
    
    fixedIncome_prices = None  # Initialize fixedIncome_prices to avoid UnboundLocalError
    
    if flag_fixedIncome_MtM != 0 and len(fixedIncome_list) > 0:
        print("NTN-B prices...")
        fixedIncome_prices = fixed_income_prices_database(fixedIncome_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5))
        fixedIncome_prices.rename(columns=dict_CNPJ, inplace=True)
        new_columns = list(portfolio[((portfolio['CNPJ'].isin(fixedIncome_list)) & (~portfolio['ATIVO'].isin(list(fixedIncome_prices.columns))))]['ATIVO'])
        fixedIncome_prices[new_columns] = np.nan 
        print("Done.")
    
    stock_prices = None  # Initialize stock_prices to avoid UnboundLocalError
    
    if len(stocks_list) > 0:
        print("Getting stock prices...")
        stock_prices = stock_prices_database(stocks_list, date_first - dt.timedelta(days=5), date_last + dt.timedelta(days=5)) 
        new_columns = list(portfolio[((portfolio['CNPJ'].isin(stocks_list)) & (~portfolio['ATIVO'].isin(list(stock_prices.columns))))]['ATIVO'])
        stock_prices[new_columns] = np.nan 
        print("Done.")
    
    print("Getting benchmark prices...")
    benchmark_prices = benchmark_prices_database(benchmark_list, date_first - dt.timedelta(days=62), date_last + dt.timedelta(days=5)) 
    print("Done.")
    
    return portfolio, fund_prices, fixedIncome_prices, stock_prices, benchmark_prices

''' 4) MANIPULATE PORTFOLIO CATEGORICAL COLUMNS ----------------------------------------------------------------------------------------------------------'''

def manipulate_portfolio_categorical_columns(portfolio):
    print("Completing proxy columns...")
    
    # Checking if the expected columns are present in the DataFrame
    expected_columns = ['BENCH+', '%BENCH', 'BENCH']
    for column in expected_columns:
        if column not in portfolio.columns:
            portfolio[column] = np.nan  # Keep as NaN if not present
    
    print("Colunas disponíveis no DataFrame 'portfolio':", portfolio.columns)

    portfolio['Ativo'] = portfolio['ATIVO']
    portfolio['BENCH+'] = portfolio['BENCH+'].fillna(np.nan)
    portfolio['%BENCH'] = portfolio['%BENCH'].fillna(np.nan)
    portfolio['BENCH'] = portfolio['BENCH'].fillna(np.nan)
    
    print("Done.")
    return portfolio



''' 5) GET ASSETS AND BENCHMARKS DAILY RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

def get_assets_and_benchmarks_daily_returns(portfolio, date_first, date_last, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ, fund_prices, fixedIncome_prices, stock_prices, benchmark_prices):
    
    fund_Returns = None
    fixedIncome_Returns = None
    stock_Returns = None

    if len(cnpj_list) > 0:
        print("Calculating fund returns...")
        # Fund daily returns:
        fund_prices.fillna(method='ffill', inplace=True)
        fund_Returns = fund_prices.astype('float') / fund_prices.astype('float').shift(1) - 1
        fund_Returns.iloc[0:1,:].fillna(0, inplace=True)
        fund_Returns = fund_Returns[((fund_Returns.index>=date_first) & (fund_Returns.index<=date_last))]
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
    
    if 'IPCA_m/yyyy' in benchmark_prices.columns:
        benchmark_prices = benchmark_prices.drop(['IPCA_m/yyyy'], axis=1)
    
    # Benchmark daily returns given as prices:
    benchmark_Returns = benchmark_prices.astype('float') / benchmark_prices.astype('float').shift(1) - 1
    benchmark_Returns = benchmark_Returns.drop(['DU'], axis = 1)
    
    # CDI, SELIC and Previa IPCA daily returns (given as rates, not prices)
    benchmark_Returns.loc[:,'CDI'] = (1+benchmark_prices.loc[:,'CDI'].astype('float')/100)**(1/252)-1
    benchmark_Returns.loc[:,'SELIC'] = (1+benchmark_prices.loc[:,'SELIC'].astype('float')/100)**(1/252)-1
    
    if 'Prévia IPCA' in benchmark_prices.columns:
        benchmark_Returns.loc[:,'Prévia IPCA'] = (1+benchmark_prices.loc[:,'Prévia IPCA'].astype('float')/100)**(1/benchmark_prices.loc[:, "DU"])-1
    
    # IPCA
    benchmark_Returns.loc[:, "IPCA"] = (1+benchmark_Returns.loc[:,'IPCA'])**(1/benchmark_prices.loc[:, "DU"])-1 # calculate daily IPCA rates
    benchmark_Returns.loc[:, "IPCA"] = benchmark_Returns.loc[:, "IPCA"].replace(0,np.nan)
    
    if 'Prévia IPCA' in benchmark_prices.columns:
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
    
    if 'Prévia IPCA_y' in benchmark_Returns.columns:
        benchmark_Returns = benchmark_Returns.drop(columns = ['IPCA_m/yyyy', 'Prévia IPCA_x', 'Prévia IPCA_y'])
    elif 'IPCA_m/yyyy' in benchmark_Returns.columns:
        benchmark_Returns = benchmark_Returns.drop(columns = ['IPCA_m/yyyy'])
    
    benchmark_Returns = benchmark_Returns[((benchmark_Returns.index>=date_first) & (benchmark_Returns.index<=date_last))]
    benchmark_Returns.iloc[0,:] = 0
    print("Done.")

    return fund_Returns, fixedIncome_Returns, stock_Returns, benchmark_Returns


''' 6) CALCULATE PORTFOLIO RETURNS --------------------------------------------------------------------------------------------------------------------------------'''

def calculate_portfolio_returns(portfolio, assets_returns, fund_Returns, stock_Returns, fixedIncome_Returns, benchmark_Returns, cnpj_list, stocks_list, fixedIncome_list, flag_MFO, taxa_gestao, benchmark, flag_fixedIncome_MtM):
    portfolio = portfolio.set_index('ATIVO')
    assets_returns = pd.DataFrame(index=benchmark_Returns.index, columns=list(portfolio.index))
    print("Calculating portfolio returns...")

    portfolio['%BENCH'] = pd.to_numeric(portfolio['%BENCH'], errors='coerce')
    portfolio['BENCH+'] = pd.to_numeric(portfolio['BENCH+'], errors='coerce')

    for asset in assets_returns.columns:
        # Check if the asset is in fund_Returns
        if asset in fund_Returns.columns:
            assets_returns[asset] = fund_Returns.loc[assets_returns.index, asset]
        elif asset in stock_Returns.columns:
            assets_returns[asset] = stock_Returns.loc[assets_returns.index, asset]
        elif flag_fixedIncome_MtM == 1 and asset in fixedIncome_Returns.columns:
            assets_returns[asset] = fixedIncome_Returns.loc[assets_returns.index, asset]

        # Fill benchmark returns for assets not included above or to those dates in which they have no prices
        if portfolio.loc[asset, '%BENCH'] != 0:
            if portfolio.loc[asset, 'BENCH+'] != 0:
                assets_returns.loc[assets_returns[asset].isna(), asset] = (
                    1 + benchmark_Returns.loc[assets_returns[asset].isna(), portfolio.loc[asset, 'BENCH']] *
                    portfolio.loc[asset, '%BENCH']) * (
                    1 + portfolio.loc[asset, 'BENCH+']) ** (1 / 252) - 1
            else:
                assets_returns.loc[assets_returns[asset].isna(), asset] = benchmark_Returns.loc[
                    assets_returns[asset].isna(), portfolio.loc[asset, 'BENCH']] * portfolio.loc[asset, '%BENCH']
        elif portfolio.loc[asset, 'BENCH+'] != 0:
            assets_returns.loc[assets_returns[asset].isna(), asset] = (1 + portfolio.loc[asset, 'BENCH+']) ** (1 / 252) - 1
        elif portfolio.loc[asset, 'BENCH'] != "-":
            assets_returns.loc[assets_returns[asset].isna(), asset] = benchmark_Returns.loc[
                assets_returns[asset].isna(), portfolio.loc[asset, 'BENCH']]

    # Get weighted performance for each asset
    assets_returns_W = assets_returns.multiply(portfolio['%PL_NOVO'], axis=1)

    # PORTFOLIO
    portfolio_return = assets_returns_W.sum(axis=1)  # Daily returns
    if flag_MFO == 1:
        portfolio_return = portfolio_return + ((1 - taxa_gestao) ** (1 / 252) - 1)  # Subtract management fee

    portfolio_acc = portfolio_return.add(1).cumprod().sub(1)
    portfolio_acc = pd.concat([portfolio_acc, benchmark_Returns[benchmark].add(1).cumprod().sub(1)], axis=1)
    portfolio_acc.columns = ['Portfólio Modelo', benchmark]

    # ASSET STRATEGIES
    strategy_list = portfolio['ESTRATÉGIA'].unique()
    strategy_weights = portfolio.groupby('ESTRATÉGIA')['%PL_NOVO'].sum()

    strategy_returns = pd.DataFrame(index=benchmark_Returns.index, columns=strategy_list)  # Daily returns:
    for strategy in strategy_list:
        asset_group = portfolio[portfolio['ESTRATÉGIA'] == strategy].index
        strategy_returns[strategy] = assets_returns_W[asset_group].sum(axis=1)  # Daily returns

    strategy_attr = strategy_returns.divide(strategy_weights, axis=1)
    strategy_acc = strategy_returns.add(1).cumprod().sub(1)

    # ASSET CLASSES
    class_list = portfolio['CLASSE'].unique()
    class_weights = portfolio.groupby('CLASSE')['%PL_NOVO'].sum()

    class_returns = pd.DataFrame(index=benchmark_Returns.index, columns=class_list)  # Daily returns:
    for class_type in class_list:
        asset_group = portfolio[portfolio['CLASSE'] == class_type].index
        class_returns[class_type] = assets_returns_W[asset_group].sum(axis=1)  # Daily returns

    class_attr = class_returns.divide(class_weights, axis=1)
    class_acc = class_returns.add(1).cumprod().sub(1)

    print("Done.")

    return assets_returns, portfolio_return, portfolio_acc, strategy_weights, strategy_returns, strategy_attr, strategy_acc, class_weights, class_returns, class_attr, class_acc

''' 7) CALCULATE PERFORMANCE METRICS --------------------------------------------------------------------------------------------------------------------------------'''

def calculate_correlation_matrix(assets_returns, benchmark_Returns, benchmark, portfolio, cnpj_list):
    asset_group = [benchmark]
    cnpj_list_str = [str(cnpj) for cnpj in cnpj_list]  # Convert CNPJ list to string for comparison
    for j in assets_returns.columns:
        if str(portfolio.loc[portfolio['ATIVO'] == j, 'CNPJ'].values[0]) in cnpj_list_str or np.std(assets_returns[j]) * np.sqrt(252) > 0.01:
            asset_group.append(j)

    assets_returns2 = pd.concat([assets_returns, benchmark_Returns[benchmark]], axis=1)
    first_column = assets_returns2.pop(benchmark)
    assets_returns2.insert(0, benchmark, first_column)
    assets_returns2 = assets_returns2[asset_group]

    correlation = assets_returns2.corr()
    for i in range(correlation.shape[0]):
        for j in range(correlation.shape[1]):
            if i < j:
                correlation.iloc[i, j] = "-"

    rows_correl = list(correlation.index)
    if len(rows_correl) <= 26:
        alphabet = list(map(chr, range(65, 91)))[0:len(rows_correl)]
    else:
        alphabet1 = list(map(chr, range(65, 91)))[0:26]
        alphabet2 = list(map(chr, range(65, 91)))[0:len(rows_correl) - 26]
        alphabet2 = ["A" + letter for letter in alphabet2]
        alphabet = alphabet1 + alphabet2

    columns_correl = list('(' + a + ") " for a in alphabet)
    rows_correl = [x + y for x, y in zip(columns_correl, rows_correl)]

    correlation.columns = columns_correl
    correlation.index = rows_correl
    print("Done.")

    return correlation

def calculate_performance_metrics(class_attr, strategy_attr, flag_MFO, taxa_gestao, portfolio_acc, portfolio_return, benchmark, date_first, amount, benchmark_Returns, assets_returns, portfolio, cnpj_list):
    print("Calculating performance metrics...")
    # Performance attribution
    class_perf_attr = class_attr.sum(axis=0)
    strategy_perf_attr = strategy_attr.sum(axis=0)
    if flag_MFO == 1:
        class_perf_attr['Taxa de Gestão'] = (1 - taxa_gestao)**(class_attr.shape[0] / 252) - 1
        strategy_perf_attr['Taxa de Gestão'] = (1 - taxa_gestao)**(strategy_attr.shape[0] / 252) - 1

    class_perf_attr['Total'] = class_perf_attr.sum()
    strategy_perf_attr['Total'] = strategy_perf_attr.sum()

    # Portfolio vs. Benchmark: Return, Vol, Sharpe
    if benchmark == "":
        benchmark = "CDI"

    portf_vs_bench_1 = pd.DataFrame(columns=["Rentabilidade Acumulada", "Rentabilidade Anualizada", "Volatilidade Anualizada", "Sharpe"],
                                    index=["Portfólio Modelo", benchmark])

    portf_vs_bench_1.iloc[0, 0] = portfolio_acc.iloc[-1, 0]
    portf_vs_bench_1.iloc[1, 0] = portfolio_acc.iloc[-1, 1]
    portf_vs_bench_1.iloc[0, 1] = (1 + portfolio_acc.iloc[-1, 0])**(252 / portfolio_acc.shape[0]) - 1
    portf_vs_bench_1.iloc[1, 1] = (1 + portfolio_acc.iloc[-1, 1])**(252 / portfolio_acc.shape[0]) - 1
    portf_vs_bench_1.iloc[0, 2] = np.sqrt(252) * np.std(portfolio_return)
    portf_vs_bench_1.iloc[1, 2] = np.sqrt(252) * np.std(benchmark_Returns[benchmark])
    portf_vs_bench_1.iloc[0, 3] = (portf_vs_bench_1.iloc[0, 1] - (((benchmark_Returns['CDI'] + 1).to_numpy().prod())**(252 / benchmark_Returns.shape[0]) - 1)) / portf_vs_bench_1.iloc[0, 2]
    portf_vs_bench_1.iloc[1, 3] = (portf_vs_bench_1.iloc[1, 1] - (((benchmark_Returns['CDI'] + 1).to_numpy().prod())**(252 / benchmark_Returns.shape[0]) - 1)) / portf_vs_bench_1.iloc[1, 2]

    # Portfolio vs. Benchmark: Returns
    portf_vs_bench_2 = pd.DataFrame(columns=["Mes", "Ano", "6 meses", "12 meses", "2 anos", "Saldo Inicial (" + date_first.strftime("%d/%m/%Y") + ")", "Saldo Final (" + portfolio_acc.index[-1].strftime("%d/%m/%Y") + ")"],
                                    index=["Portfólio Modelo", benchmark])

    date_MTD = dt.date(portfolio_return.index[-1].year, portfolio_return.index[-1].month, 1)
    date_YTD = dt.date(portfolio_return.index[-1].year, 1, 1)
    date_6M = portfolio_return.index[-1] - relativedelta(months=6)
    date_12M = portfolio_return.index[-1] - relativedelta(months=12)
    date_24M = portfolio_return.index[-1] - relativedelta(months=24)

    portf_vs_bench_2.iloc[0, 0] = (1 + portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_MTD)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1, 0] = (1 + benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_MTD)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0, 1] = (1 + portfolio_return[(pd.to_datetime(portfolio_return.index).date > date_YTD)]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1, 1] = (1 + benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date > date_YTD)][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0, 2] = (1 + portfolio_return[(portfolio_return.index >= pd.Timestamp(date_6M))]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1, 2] = (1 + benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_6M))][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0, 3] = (1 + portfolio_return[(portfolio_return.index >= pd.Timestamp(date_12M))]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1, 3] = (1 + benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_12M))][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0, 4] = (1 + portfolio_return[(portfolio_return.index >= pd.Timestamp(date_24M))]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[1, 4] = (1 + benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_24M))][benchmark]).to_numpy().prod() - 1
    portf_vs_bench_2.iloc[0, 5] = amount
    portf_vs_bench_2.iloc[1, 5] = amount
    portf_vs_bench_2.iloc[0, 6] = amount * (1 + portfolio_acc.iloc[-1, 0])
    portf_vs_bench_2.iloc[1, 6] = amount * (1 + portfolio_acc.iloc[-1, 1])

    # Portfolio vs. Benchmark: Monthly results
    portf_vs_bench_aux = portfolio_return.copy()
    portf_vs_bench_aux.name = "Portfólio Modelo"
    portf_vs_bench_aux = pd.concat([portf_vs_bench_aux, benchmark_Returns[benchmark]], axis=1)

    portf_vs_bench_aux.index = pd.to_datetime(portf_vs_bench_aux.index)
    portf_vs_bench_3 = portf_vs_bench_aux.groupby(pd.Grouper(freq='M')).apply(lambda x: (1 + x).prod() - 1)
    portf_vs_bench_3.index = portf_vs_bench_3.index.strftime("%Y-%m")

    portf_vs_bench_3['%' + benchmark] = portf_vs_bench_3.apply(lambda row: '-' if row[benchmark] < 0 else row["Portfólio Modelo"] / row[benchmark], axis=1)

    portf_vs_bench_3.index = [sub.replace('-01', 'Jan.').replace('-02', 'Fev.').replace('-03', 'Mar.').replace('-04', 'Abr.').replace('-05', 'Mai.')
                              .replace('-06', 'Jun.').replace('-07', 'Jul.').replace('-08', 'Ago.').replace('-09', 'Set.').replace('-10', 'Out.')
                              .replace('-11', 'Nov.').replace('-12', 'Dez.') for sub in list(portf_vs_bench_3.index)]

    # Portfolio vs. Benchmark: Statistics
    portf_vs_bench_4 = pd.DataFrame(columns=["Meses\nPositivos", "Meses\nNegativos", "Maior Retorno\nMensal", "Menor Retorno\nMensal", "Acima do CDI\n(meses)", "Abaixo do CDI\n(meses)"],
                                    index=["Portfólio Modelo", benchmark])

    portf_vs_bench_4.iloc[0, 0] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] >= 0)]["Portfólio Modelo"].count()
    portf_vs_bench_4.iloc[1, 0] = portf_vs_bench_3[(portf_vs_bench_3[benchmark] >= 0)][benchmark].count()
    portf_vs_bench_4.iloc[0, 1] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] < 0)]["Portfólio Modelo"].count()
    portf_vs_bench_4.iloc[1, 1] = portf_vs_bench_3[(portf_vs_bench_3[benchmark] < 0)][benchmark].count()
    portf_vs_bench_4.iloc[0, 2] = portf_vs_bench_3["Portfólio Modelo"].max()
    portf_vs_bench_4.iloc[1, 2] = portf_vs_bench_3[benchmark].max()
    portf_vs_bench_4.iloc[0, 3] = portf_vs_bench_3["Portfólio Modelo"].min()
    portf_vs_bench_4.iloc[1, 3] = portf_vs_bench_3[benchmark].min()

    if benchmark == "CDI":
        portf_vs_bench_4.iloc[0, 4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] >= portf_vs_bench_3[benchmark])]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[0, 5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] < portf_vs_bench_3[benchmark])]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1, 4] = 0
        portf_vs_bench_4.iloc[1, 5] = 0
    else:
        retorno_CDI_M_aux = benchmark_Returns["CDI"].copy()
        retorno_CDI_M_aux.name = "retorno_CDI_M"
        retorno_CDI_M_aux.index = pd.to_datetime(portf_vs_bench_aux.index)
        retorno_CDI_M = retorno_CDI_M_aux.groupby(pd.Grouper(freq='M')).apply(lambda x: (1 + x).prod() - 1)
        retorno_CDI_M.index = retorno_CDI_M.index.strftime("%Y-%m")

        retorno_CDI_M.index = [sub.replace('-01', 'Jan.').replace('-02', 'Fev.').replace('-03', 'Mar.').replace('-04', 'Abr.').replace('-05', 'Mai.')
                               .replace('-06', 'Jun.').replace('-07', 'Jul.').replace('-08', 'Ago.').replace('-09', 'Set.').replace('-10', 'Out.')
                               .replace('-11', 'Nov.').replace('-12', 'Dez.') for sub in list(retorno_CDI_M.index)]

        portf_vs_bench_4.iloc[0, 4] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] >= retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[0, 5] = portf_vs_bench_3[(portf_vs_bench_3["Portfólio Modelo"] < retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1, 4] = portf_vs_bench_3[(portf_vs_bench_3[benchmark] >= retorno_CDI_M)]["Portfólio Modelo"].count()
        portf_vs_bench_4.iloc[1, 5] = portf_vs_bench_3[(portf_vs_bench_3[benchmark] < retorno_CDI_M)]["Portfólio Modelo"].count()

    # Portfolio vs. Benchmark: Volatility
    portf_vs_bench_5 = pd.DataFrame(columns=["Mes", "Ano", "6 meses", "12 meses", "2 anos"],
                                    index=["Portfólio Modelo", benchmark])

    portf_vs_bench_5.iloc[0, 0] = np.sqrt(252) * np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_MTD)])
    portf_vs_bench_5.iloc[1, 0] = np.sqrt(252) * np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_MTD)][benchmark])
    portf_vs_bench_5.iloc[0, 1] = np.sqrt(252) * np.std(portfolio_return[(pd.to_datetime(portfolio_return.index).date >= date_YTD)])
    portf_vs_bench_5.iloc[1, 1] = np.sqrt(252) * np.std(benchmark_Returns[(pd.to_datetime(benchmark_Returns.index).date >= date_YTD)][benchmark])
    portf_vs_bench_5.iloc[0, 2] = np.sqrt(252) * np.std(portfolio_return[(portfolio_return.index >= pd.Timestamp(date_6M))])
    portf_vs_bench_5.iloc[1, 2] = np.sqrt(252) * np.std(benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_6M))][benchmark])
    portf_vs_bench_5.iloc[0, 3] = np.sqrt(252) * np.std(portfolio_return[(portfolio_return.index >= pd.Timestamp(date_12M))])
    portf_vs_bench_5.iloc[1, 3] = np.sqrt(252) * np.std(benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_12M))][benchmark])
    portf_vs_bench_5.iloc[0, 4] = np.sqrt(252) * np.std(portfolio_return[(portfolio_return.index >= pd.Timestamp(date_24M))])
    portf_vs_bench_5.iloc[1, 4] = np.sqrt(252) * np.std(benchmark_Returns[(benchmark_Returns.index >= pd.Timestamp(date_24M))][benchmark])

    volatility = pd.concat([portfolio_return, benchmark_Returns[benchmark]], axis=1)
    volatility.columns = ["Portfólio Modelo", benchmark]

    for i in range(volatility.shape[0] - 21):  # Rolling 21-days volatility
        volatility.iloc[i + 21, 0] = np.sqrt(252) * np.std(portfolio_return[i:i + 21])
        volatility.iloc[i + 21, 1] = np.sqrt(252) * np.std(benchmark_Returns[benchmark][i:i + 21])

    volatility = volatility.iloc[21:, :]

    # Portfolio vs. Benchmark: Drawdown
    portf_vs_bench_6 = pd.DataFrame(columns=["Mes", "Ano", "6 meses", "12 meses", "2 anos", "Drawdown Máximo", "Data", "Tempo de Recuperação"],
                                    index=["Portfólio Modelo", benchmark])

    drawdown = pd.DataFrame(np.zeros((portfolio_return.shape[0], 2)), columns=["Portfólio Modelo", benchmark], index=portfolio_return.index)

    drawdown_acc_returns = portfolio_acc.copy()

    drawdown_MTD = drawdown[(pd.to_datetime(drawdown.index).date >= date_MTD)]
    drawdown_YTD = drawdown[(pd.to_datetime(drawdown.index).date >= date_YTD)]
    drawdown_6M = drawdown[(drawdown.index >= pd.Timestamp(date_6M))]
    drawdown_12M = drawdown[(drawdown.index >= pd.Timestamp(date_12M))]
    drawdown_24M = drawdown[(drawdown.index >= pd.Timestamp(date_24M))]

    # Calculate drawdown series (entire period)
    for i in range(drawdown.shape[0] - 1):
        drawdown.iloc[i + 1, :] = (1 + drawdown.iloc[i, :]) * (1 + drawdown_acc_returns.iloc[i + 1, :]) / (1 + drawdown_acc_returns.iloc[i, :]) - 1
        drawdown.iloc[i + 1, :].values[drawdown.iloc[i + 1, :].values > 0] = 0

    # Calculate drawdown series (MTD)
    for i in range(drawdown_MTD.shape[0] - 1):
        drawdown_MTD.iloc[i + 1, :] = (1 + drawdown_MTD.iloc[i, :]) * (1 + drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_MTD)].iloc[i + 1, :]) / (1 + drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_MTD)].iloc[i, :]) - 1
        drawdown_MTD.iloc[i + 1, :].values[drawdown_MTD.iloc[i + 1, :].values > 0] = 0

    # Calculate drawdown series (YTD)
    for i in range(drawdown_YTD.shape[0] - 1):
        drawdown_YTD.iloc[i + 1, :] = (1 + drawdown_YTD.iloc[i, :]) * (1 + drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_YTD)].iloc[i + 1, :]) / (1 + drawdown_acc_returns[(pd.to_datetime(drawdown_acc_returns.index).date >= date_YTD)].iloc[i, :]) - 1
        drawdown_YTD.iloc[i + 1, :].values[drawdown_YTD.iloc[i + 1, :].values > 0] = 0

    # Calculate drawdown series (6M)
    for i in range(drawdown_6M.shape[0] - 1):
        drawdown_6M.iloc[i + 1, :] = (1 + drawdown_6M.iloc[i, :]) * (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_6M))].iloc[i + 1, :]) / (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_6M))].iloc[i, :]) - 1
        drawdown_6M.iloc[i + 1, :].values[drawdown_6M.iloc[i + 1, :].values > 0] = 0

    # Calculate drawdown series (12M)
    for i in range(drawdown_12M.shape[0] - 1):
        drawdown_12M.iloc[i + 1, :] = (1 + drawdown_12M.iloc[i, :]) * (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_12M))].iloc[i + 1, :]) / (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_12M))].iloc[i, :]) - 1
        drawdown_12M.iloc[i + 1, :].values[drawdown_12M.iloc[i + 1, :].values > 0] = 0
    # Calculate drawdown series (24M)
    for i in range(drawdown_24M.shape[0] - 1):
        drawdown_24M.iloc[i + 1, :] = (1 + drawdown_24M.iloc[i, :]) * (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_24M))].iloc[i + 1, :]) / (1 + drawdown_acc_returns[(drawdown_acc_returns.index >= pd.Timestamp(date_24M))].iloc[i, :]) - 1
        drawdown_24M.iloc[i + 1, :].values[drawdown_24M.iloc[i + 1, :].values > 0] = 0

    portf_vs_bench_6.iloc[0, 0] = min(drawdown_MTD.iloc[:, 0])
    portf_vs_bench_6.iloc[1, 0] = min(drawdown_MTD.iloc[:, 1])
    portf_vs_bench_6.iloc[0, 1] = min(drawdown_YTD.iloc[:, 0])
    portf_vs_bench_6.iloc[1, 1] = min(drawdown_YTD.iloc[:, 1])
    portf_vs_bench_6.iloc[0, 2] = min(drawdown_6M.iloc[:, 0])
    portf_vs_bench_6.iloc[1, 2] = min(drawdown_6M.iloc[:, 1])
    portf_vs_bench_6.iloc[0, 3] = min(drawdown_12M.iloc[:, 0])
    portf_vs_bench_6.iloc[1, 3] = min(drawdown_12M.iloc[:, 1])
    portf_vs_bench_6.iloc[0, 4] = min(drawdown_24M.iloc[:, 0])
    portf_vs_bench_6.iloc[1, 4] = min(drawdown_24M.iloc[:, 1])

    if min(drawdown.iloc[:, 0]) == 0:
        portf_vs_bench_6.iloc[0, 5] = 0
        portf_vs_bench_6.iloc[0, 6] = "-"
        portf_vs_bench_6.iloc[0, 7] = "-"
    else:
        portf_vs_bench_6.iloc[0, 5] = min(drawdown.iloc[:, 0])
        portf_vs_bench_6.iloc[0, 6] = list(drawdown[(drawdown.iloc[:, 0] == min(drawdown.iloc[:, 0]))].index.strftime('%d/%m/%Y'))[0]
        aux_time_recover1 = drawdown.loc[:drawdown.iloc[:, 0].idxmin()].iloc[:, 0]
        aux_time_recover1 = aux_time_recover1[(aux_time_recover1.index > max(aux_time_recover1[(aux_time_recover1 == 0)].index))]
        aux_time_recover2 = drawdown.loc[drawdown.iloc[:, 0].idxmin():].iloc[:, 0]

        if len(aux_time_recover2[(aux_time_recover2 == 0)].index == 0) > 0:
            aux_time_recover2 = aux_time_recover2[(aux_time_recover2.index <= min(aux_time_recover2[(aux_time_recover2 == 0)].index))]
            portf_vs_bench_6.iloc[0, 7] = len(aux_time_recover1) + len(aux_time_recover2) - 1
        else:  # has not yet recovered
            time_recover = len(aux_time_recover2)
            portf_vs_bench_6.iloc[0, 7] = str(time_recover) + '(+...) '

    if min(drawdown.iloc[:, 1]) == 0:
        portf_vs_bench_6.iloc[1, 5] = 0
        portf_vs_bench_6.iloc[1, 6] = "-"
        portf_vs_bench_6.iloc[1, 7] = "-"
    else:
        portf_vs_bench_6.iloc[1, 5] = min(drawdown.iloc[:, 1])
        portf_vs_bench_6.iloc[1, 6] = list(drawdown[(drawdown.iloc[:, 1] == min(drawdown.iloc[:, 1]))].index.strftime('%d/%m/%Y'))[0]
        aux_time_recover1 = drawdown.loc[:drawdown.iloc[:, 1].idxmin()].iloc[:, 1]
        aux_time_recover1 = aux_time_recover1[(aux_time_recover1.index > max(aux_time_recover1[(aux_time_recover1 == 0)].index))]
        aux_time_recover2 = drawdown.loc[drawdown.iloc[:, 1].idxmin():].iloc[:, 1]

        if len(aux_time_recover2[(aux_time_recover2 == 0)].index == 0) > 0:
            aux_time_recover2 = aux_time_recover2[(aux_time_recover2.index <= min(aux_time_recover2[(aux_time_recover2 == 0)].index))]
            portf_vs_bench_6.iloc[1, 7] = len(aux_time_recover1) + len(aux_time_recover2) - 1
        else:  # has not yet recovered
            time_recover = len(aux_time_recover2)
            portf_vs_bench_6.iloc[1, 7] = str(time_recover) + '(+...) '

    correlation = calculate_correlation_matrix(assets_returns, benchmark_Returns, benchmark, portfolio, cnpj_list)
    
    return class_perf_attr, strategy_perf_attr, portf_vs_bench_1, portf_vs_bench_2, portf_vs_bench_3, portf_vs_bench_4, portf_vs_bench_5, portf_vs_bench_6, correlation

def print_simulation_results(portfolio_acc, benchmark, benchmark_Returns, portfolio_return, date_first, date_last):
    print("\nResultados da Simulação para o Período Completo (de {0} a {1})\n".format(date_first.strftime('%Y-%m-%d'), date_last.strftime('%Y-%m-%d')))
    
    # Rentabilidade Acumulada
    accumulated_return_portfolio = portfolio_acc.iloc[-1, 0]
    accumulated_return_benchmark = portfolio_acc.iloc[-1, 1]
    print("#### Rentabilidade Acumulada")
    print(f"- Portfólio Modelo: {accumulated_return_portfolio:.2%}")
    print(f"- {benchmark}: {accumulated_return_benchmark:.2%}")
    print()
    
    # Rentabilidade Anualizada
    annualized_return_portfolio = (1 + accumulated_return_portfolio)**(252 / portfolio_acc.shape[0]) - 1
    annualized_return_benchmark = (1 + accumulated_return_benchmark)**(252 / portfolio_acc.shape[0]) - 1
    print("#### Rentabilidade Anualizada")
    print(f"- Portfólio Modelo: {annualized_return_portfolio:.2%}")
    print(f"- {benchmark}: {annualized_return_benchmark:.2%}")
    print()
    
    # Volatilidade Anualizada
    annualized_volatility_portfolio = np.sqrt(252) * np.std(portfolio_return)
    annualized_volatility_benchmark = np.sqrt(252) * np.std(benchmark_Returns[benchmark])
    print("#### Volatilidade Anualizada")
    print(f"- Portfólio Modelo: {annualized_volatility_portfolio:.2%}")
    print(f"- {benchmark}: {annualized_volatility_benchmark:.2%}")
    print()
    
    # Índice de Sharpe
    risk_free_rate = (((benchmark_Returns['CDI'] + 1).to_numpy().prod())**(252 / benchmark_Returns.shape[0]) - 1)
    sharpe_ratio_portfolio = (annualized_return_portfolio - risk_free_rate) / annualized_volatility_portfolio
    sharpe_ratio_benchmark = 0  # CDI is considered risk-free in this context
    print("#### Índice de Sharpe")
    print(f"- Portfólio Modelo: {sharpe_ratio_portfolio:.2f}")
    print(f"- {benchmark}: {sharpe_ratio_benchmark:.2f}")

def main_code():
    portfolio, date_first, date_last, benchmark, amount, taxa_gestao, flag_MFO, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ = read_selected_funds()
    if portfolio is None:
        print("Failed to read portfolio data.")
        return

    print("Colunas disponíveis após leitura dos dados:", portfolio.columns)
    
    portfolio, fund_prices, fixedIncome_prices, stock_prices, benchmark_prices = import_fund_prices(
        portfolio, date_first, date_last, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ
    )

    portfolio = manipulate_portfolio_categorical_columns(portfolio)

    # Verificar valores dos fundos durante uma semana
    print(f"Verificando valores dos fundos de {date_first} a {date_last}:")
    print(fund_prices.loc[date_first:date_last])

    # Calculation of weights based on 'PL_NOVO'
    portfolio['%PL_NOVO'] = portfolio['PL_NOVO'] / portfolio['PL_NOVO'].sum()
    print("Colunas após adicionar '%PL_NOVO':", portfolio.columns)

    fund_Returns, fixedIncome_Returns, stock_Returns, benchmark_Returns = get_assets_and_benchmarks_daily_returns(
        portfolio, date_first, date_last, flag_fixedIncome_MtM, cnpj_list, fixedIncome_list, stocks_list, benchmark_list, dict_CNPJ, fund_prices, fixedIncome_prices, stock_prices, benchmark_prices
    )

    print("\nFund Returns:")
    print(fund_Returns.describe())  # Descriptive Statistics for Funds Returns
    print(benchmark_Returns.describe())  # Descriptive Statistics for Benchmark Returns

    # Initializing assets_returns correctly
    assets_returns = pd.DataFrame(index=benchmark_Returns.index)

    assets_returns, portfolio_return, portfolio_acc, strategy_weights, strategy_returns, strategy_attr, strategy_acc, class_weights, class_returns, class_attr, class_acc = calculate_portfolio_returns(
        portfolio, assets_returns, fund_Returns, stock_Returns, fixedIncome_Returns, benchmark_Returns, cnpj_list, stocks_list, fixedIncome_list, flag_MFO, taxa_gestao, benchmark, flag_fixedIncome_MtM
    )

    print("\nPortfolio Returns:")
    print(portfolio_return.describe())  # Descriptive Statistics for Portfolio Returns
    print("\nStrategy Weights:")
    print(strategy_weights)
    print("\nClass Weights:")
    print(class_weights)

    if portfolio_return.empty:
        print("Erro: O DataFrame 'portfolio_return' está vazio.")
        return
    
    if class_attr.empty:
        print("Erro: O DataFrame 'class_attr' está vazio.")
        return

    print("Cálculo das métricas de performance será feito na próxima etapa.")

    # Section 7: Performance Metrics calculations
    class_perf_attr, strategy_perf_attr, portf_vs_bench_1, portf_vs_bench_2, portf_vs_bench_3, portf_vs_bench_4, portf_vs_bench_5, portf_vs_bench_6, correlation = calculate_performance_metrics(
        class_attr, strategy_attr, flag_MFO, taxa_gestao, portfolio_acc, portfolio_return, benchmark, date_first, amount, benchmark_Returns, assets_returns, portfolio, cnpj_list
    )

    print_simulation_results(portfolio_acc, benchmark, benchmark_Returns, portfolio_return, date_first, date_last)

    # Performance metrics debugging
    print("\nCorrelation Matrix:\n", correlation)

if __name__ == "__main__":
    main_code()
