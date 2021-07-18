# current version=1.06 (13-7-2021)

import pandas as pd
# import time

# targetTime = time.ctime()
# if (targetTime >= "Sat Jan 01 00:00:00 2022"):
# 	print('Your application lease has expired, please contact developer')
# 	quit()

excel_file = input('Enter the name and extension of the raw data file: ')
if '.csv' not in excel_file:
    excel_file = str(f'{excel_file}.csv')
else:
    pass
source_file = 'rowa_database.xlsx'

df_file = pd.read_csv(excel_file, sep=';')
df_source = pd.read_excel(source_file)

for column in df_file.columns:
    if column == 'Serial number,':
        df_file.rename(columns={'Serial number,': 'Serial number'}, inplace=True)

df_file.drop(['Id', 'Source', 'Input date', 'New delivery',
              'Serial number'], axis=1, inplace=True)

df_file.rename(columns={'Barcode': 'NDC'}, inplace=True)
df_file.rename(columns={'External code': 'PHX'}, inplace=True)
df_file.rename(columns={'Article name': 'Description'}, inplace=True)
df_file = df_file.reindex(
    columns=['PHX', 'Description', 'Expiration date', 'Batch number', 'NDC'])

df_output = pd.merge(df_file, df_source[['NDC', 'UOM']], on='NDC', how='left')
df_output = pd.merge(
    df_output, df_source[['NDC', 'MaxSubQty']], on='NDC', how='left')
df_output.rename(columns={'MaxSubQty': 'QOH'}, inplace=True)
df_output.drop(['NDC'], axis=1, inplace=True)
df_output['UOM'].replace({'cap': 'ea', 'tab': 'ea'}, inplace=True)
df_output.loc[df_output['UOM'] == 'mL', 'QOH'] = 1

df_final = df_output.groupby(['PHX', 'Description', 'Expiration date',
                              'Batch number', 'UOM'], as_index=False).agg({'QOH': 'sum'})

df_final.to_excel(f'{excel_file.rstrip(".csv")}_final.xlsx', index=False)
