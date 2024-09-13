# current version=1.09 (28-7-2021)
import sys
from tkinter import *
from tkinter import filedialog as fd
# from tkinter.font import BOLD
import pandas as pd
from PIL import ImageTk, Image

root = Tk()
root.title('Expiry File Formatter')
# root.iconbitmap('aidIcon.ico')
canvas = Canvas(root, width=183, height=137)
canvas.grid(row=1, column=1, padx=0, pady=0)
img = ImageTk.PhotoImage(Image.open("aid.jpg"))
canvas.create_image(20, 20, anchor=NW, image=img)
# root.geometry("500x100")

title_1 = Label(
    root, text='Press "Select File"\nButtonBelow to Start',)
title_1.grid(row=0, column=1, padx=0, pady=10)


def start():
    excel_file = fd.askopenfilename(filetypes=(("Comma-Separated Values files",
                                                "*.csv*"),
                                                ("CSV Document",
                                                "*.csv"),
                                               ("all files",
                                                "*.*")))
    source_file = 'database.xlsx' #edit to match file name or rename file

    df_file = pd.read_csv(excel_file, sep=';')
    df_source = pd.read_excel(source_file)

    for column in df_file.columns:
        if column == 'Serial number,':
            df_file.rename(
                columns={'Serial number,': 'Serial number'}, inplace=True)

    df_file.drop(['Id', 'Source', 'Input date', 'New delivery',
                  'Serial number'], axis=1, inplace=True)

    df_file.rename(columns={'Barcode': 'NDC'}, inplace=True)
    df_file.rename(columns={'External code': 'PHX'}, inplace=True)
    df_file.rename(columns={'Article name': 'Description'}, inplace=True)
    df_file = df_file.reindex(
        columns=['PHX', 'Description', 'Expiration date', 'Batch number', 'NDC'])

    df_output = pd.merge(
        df_file, df_source[['NDC', 'UOM']], on='NDC', how='left')
    df_output = pd.merge(
        df_output, df_source[['NDC', 'MaxSubQty']], on='NDC', how='left')
    df_output.rename(columns={'MaxSubQty': 'QOH'}, inplace=True)
    df_output.drop(['NDC'], axis=1, inplace=True)
    df_output['UOM'].replace({'cap': 'ea', 'tab': 'ea'}, inplace=True)
    df_output.loc[df_output['UOM'] == 'mL', 'QOH'] = 1

    df_final = df_output.groupby(['PHX', 'Description', 'Expiration date',
                                  'Batch number', 'UOM'], as_index=False).agg({'QOH': 'sum'})

    df_final.to_excel(f'{excel_file.rstrip(".csv")}_output.xlsx', index=False)
    done_label = Label(
        root, text=f'Done!\n\n{excel_file.rstrip(".csv")}_output.xlsx\n\nFile Created\n\n')

    done_label.grid(row=4, column=1)


def close():
    return sys.exit()


calculate1 = Button(root, text='Select File', command=start,
                    width=10, borderwidth=2)
calculate1.grid(row=3, column=0, padx=20, pady=20)
close1 = Button(root, text='Exit', command=close,
                width=10, borderwidth=2)
close1.grid(row=3, column=2, padx=30, pady=20)

root.mainloop()
