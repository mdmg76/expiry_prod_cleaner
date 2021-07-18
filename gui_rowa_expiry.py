from tkinter import *
import pandas as pd

root = Tk()

title_1 = Label(root, text='Rowa Expiry File Formatter')
title_1.grid(row=0, column=1)


ex_file = Entry(root, width=30, borderwidth=2)
ex_file.grid(row=1, column=2)
ex_file_label = Label(root, text='Enter File Name:')
ex_file_label.grid(row=1, column=0)
# act_wt_in = Entry(root, width=10, borderwidth=2)
# act_wt_in.grid(row=2, column=1)
# act_wt_in_label = Label(root, text='Enter Weight in (kg):')
# act_wt_in_label.grid(row=2, column=0)
# gender_in = Entry(root, width=10, borderwidth=2)
# gender_in.grid(row=3, column=1)
# gender_in_label = Label(root, text='Enter Gender:')
# gender_in_label.grid(row=3, column=0)


def start():
    excel_file = ex_file.get()
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
    done_label = Label(
            root, text='Done\nFile Created')

    done_label.grid(row=4, column=0)
    # ibw_label.grid(row=7, column=0)
    # adw_label.grid(row=8, column=0)
    # lbw_label.grid(row=9, column=0)

def close():
    return quit()


calculate1 = Button(root, text='Calculate', command=start)
calculate1.grid(row=2, column=0)
close1 = Button(root, text='Exit', command=close)
close1.grid(row=2, column=2)

root.mainloop()
