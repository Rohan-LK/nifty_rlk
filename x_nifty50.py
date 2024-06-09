import pandas as pd

df_tickers = pd.read_csv('tickers/ind_nifty50list.csv')
tickers_lst = list(df_tickers['Symbol'])
# print(tickers_lst)

df_calc = pd.read_excel('results/multiple_dfs_178.xlsx', sheet_name='Summary')
df_calc['Tickers'] = df_calc['Tickers'].apply(lambda x: x[:-2])
# df_calc.head(2)

df_new = df_calc.loc[df_calc['Tickers'].isin(tickers_lst)]

df_new.reset_index(drop=True, inplace=True)
df_new.drop(columns="Unnamed: 0", inplace=True)

df_new.to_excel('results/summary_nifty50.xlsx', index=False)
