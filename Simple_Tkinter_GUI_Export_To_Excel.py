################
################
# This script is a simple GUI that allows the user to input a list of order numbers separated by a comma.
# Features: Label, EntryBox, tkinter, pandas, mysql.connector, read from excel, download to directory
################
################

import os
import sys
import time as t
import pandas as pd
import mysql.connector
from tkinter import *
from pathlib import Path



########################################################################################################################
# run queries and return result in df
def run_query(query):
    conn = mysql.connector.connect(host=host,
                                   database=db,
                                   user=username,
                                   password=password)

    result = pd.read_sql(query, con=conn)
    return result


########################################################################################################################
def create_excel_file(po_df):
    timestamp = t.strftime("%Y-%m-%d_%H%M%S")
    file_path = str(Path.home() / "Downloads")
    file_name = 'Order_Export_' + timestamp + '.xlsx'

    os.chdir(file_path)
    print('Writing data to excel file...')
    downloads_path = str(Path.home() / "Downloads")
    os.chdir(downloads_path)
    # Creates file with filepath and name
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

    po_df.to_excel(writer, sheet_name='Data', index=False)

    # Convert to table
    (max_row, max_col) = po_df.shape

    print('convert to excel table')
    # get list of column headers to us in add_table
    column_settings = []
    for header in po_df.columns:
        column_settings.append({'header': header})

    wb = writer.book
    ws = wb.get_worksheet_by_name('Data')

    # add table
    # make the columns and table name the same as the original.
    ws.add_table(0, 0, max_row, max_col - 1,
                 {'columns': column_settings, 'style': 'Table Style Medium 2', 'name': 'Table_Query_from_HQ'})

    # Make the columns wider for easier viewing.
    ws.set_column(0, max_col - 1, 15)
    ws.set_column(2, 2, 20)
    ws.set_column(3, 3, 50)

    writer.close()

    return file_name


########################################################################################################################
def get_data(order_list):
    global order_data, message

    try:
        # # run query with po list in where statement, output to excel, clear any messages that may be present, print error message in window, then clear entry box
        # query = f'''
        #             SELECT *
        #             FROM z
        #             WHERE TRUE
        #                 AND order_id IN ({order_list})
        #             ORDER BY order_id;
        #                  '''
        #
        # order_data = run_query(query)

        # dummy data
        order_data = pd.DataFrame()
        order_data['order_id'] = [1, 2, 3]
        order_data['customer'] = ['John', 'Jane', 'Doe']
        order_data['amount'] = [100, 200, 300]

        # dump into excel and download
        file_name = create_excel_file(order_data)

        txt = f'Please check your downloads \nfor: {file_name}'
        message.destroy()
        message = Label(root, text=txt, font=("Arial", 10))
        message.grid(row=5, column=0, padx=10, pady=20)
        order_entry.delete(0, END)

    except Exception as e:
        # if error, delete any possible messages that already appear, print error message in window, then clear entry box.
        txt = 'Download Error, make sure Order IDs are separated by a comma.\nContact Tori if you have further issues.'
        message.destroy()
        message = Label(root, text=txt, font=("Arial", 10), foreground='Red')
        message.grid(row=5, column=0, padx=10, pady=20)
        order_entry.delete(0, END)


########################################################################################################################
if __name__ == '__main__':
    # main application widget
    root = Tk()
    root.title('Orders_to_Excel.xlsx')
    # root.geometry("380x220")

    # global labels
    message = Label(root)

    # Create PO Label
    order_txt = Label(root, text='Enter Order Number:', font=("Arial", 14), pady=3)
    order_txt.grid(row=0, column=0)
    format_txt = Label(root, text='Please list out Order Numbers separated by a comma', font=("Arial", 9))
    format_txt.grid(row=1, column=0)

    # Create PO TextBox
    orderlist = StringVar()
    order_entry = Entry(root, font=("Arial", 13), textvariable=orderlist, width=38)
    order_entry.grid(row=2, column=0, padx=10, pady=20)

    # Create Button to Generate Data File
    export_button = Button(root, text="Generate File", font=("Arial", 11), command=lambda: get_data(orderlist.get()))
    export_button.grid(row=3, column=0, pady=10)

    root.mainloop()
