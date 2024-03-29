import pandas as pd
import numpy as np
import math

from datetime import date
from datetime import datetime
from datetime import timedelta

from os.path import exists as file_exists
import os
from openpyxl import load_workbook

file_folder = 'Deliveries'

in_file = {
    # 'file': "Open_orders_test_12.05.2022.xlsx",
    'file': "Open_orders_test_18.08.2023.xlsx",
    'sheet_name': "Open_orders_A101",
    'stock_sheet': "MBEW_A101",
    # 'stock_sheet': "A101_Stock",
    # 'SO_exceptions': "on_hold", # list of orders that are not urgent
    # 'SO_comments': 'SO_comments', # comments already given to promote SO
    # 'confirmations_SheetName': "Confirmations", # tracking of confirmations
    'PL40_Prio_List': 'Prio_List_PL40'
}

out_file = {
    'file': "Open_orders_test_py_out.xlsx"
}

confirmation_file = {
    'file': "Delivery_confirmations_py_out.xlsx"
}


def read_in_file(file:str, folder=file_folder, in_SheetName='', header_row=2):
    in_file = os.path.join(folder, file)
    print(in_file)

    try:
        wb = pd.read_excel(in_file, sheet_name=None, header=None)
        print('1')
        sheets = list(wb.keys())
        if in_SheetName == '':
            print("Loading last sheet in file")
#             df = pd.read_excel(in_file, sheet_name = -1, header = header_row)
            df = wb.get(sheets[-1])
        else:
            #             df = pd.read_excel(in_file, in_SheetName, header = header_row)
            df = wb.get(in_SheetName)

        del (wb)

    except TypeError:
        wb = load_workbook(filename=in_file, data_only=True)
        print('2')

        if in_SheetName == '':
            print("Loading last sheet in file")
#             wb = load_workbook(filename = in_file, data_only = True)
            df = pd.DataFrame(wb[wb.sheetnames[-1]].values)

        else:
            #             wb = load_workbook(filename = in_file, data_only = True)
            df = pd.DataFrame(wb[in_SheetName].values)

        sheets = wb.sheetnames
        wb.close()

    except FileNotFoundError:
        df = pd.DataFrame()
        print(f'File {in_file} not found')

    # except Exception as Err:
    #     print("oops!")
    #     print(Err)

    column_names = rename_columns(df.iloc[header_row].values)
    df.columns = column_names
    df.drop(index=list(np.arange(header_row+1)), inplace=True)
    df.reset_index(drop=True, inplace=True)

    return df, sheets


def rename_columns(column_names):
    for item in enumerate(column_names):
        try:
            column_names[item[0]] = '_'.join(item[1].split())
        except:
            column_names[item[0]] = item[1]
    return column_names

# print(file_exists(os.path.join(file_folder, in_file['file'])))


df, sheets = read_in_file(file=in_file['file'], folder=file_folder, in_SheetName=in_file['sheet_name'], header_row=2)
df_a101, _ = read_in_file(file=in_file['file'], folder=file_folder, in_SheetName=in_file['stock_sheet'], header_row=1)
print(sheets)

df = df.convert_dtypes()
# drop rows with all N/A values
df = df.dropna(thresh=len(df.columns))
# print(df.head())

df_a101 = df_a101.convert_dtypes()
# drop rows with all N/A values
df_a101 = df_a101.dropna(thresh=len(df_a101.columns))
df_a101 = df_a101[['Material', 'Total_Stock']]
df_a101.set_index('Material', drop=True, inplace=True)
# print(df_a101.head())

# filtration part
df = df.drop(['Product_hierarchy', 'Product_hierarchy_text',
             'Plant', 'Order_Value'], axis=1)
df.insert(0, 'item_index', df['Sales_document'] +
          '_' + df['Sales_Document_Item'])
df.insert(1, 'SO_Mat_index', df['Sales_document'] + '_' + df['Material'])
df.drop(df[df['SD_Document_Category'] != 'Order'].index, inplace=True)

df = df.join(df_a101, on='Material', how='left')
# print(df.head())

# check changes in confirmation dates
# drop lines without confirmation date
df_cur_conf = df[df['Confirmed_Date'] != '#']

# drop already partially delivered lines
df_part_delivered = df_cur_conf[(df_cur_conf['Delivery_Quantity'] > 0) &
                                (df_cur_conf['Confirmed_Date'] <= df_cur_conf['Created_on'].max())]
df_cur_conf = df_cur_conf.drop(df_part_delivered.index)
df_cur_conf = df_cur_conf.sort_values(
    by=['item_index', 'Confirmed_Date'], ascending=[True, True])
cur_conf_col = 'Confirmation_' + df_cur_conf['Created_on'].max()


def adj_requested_date(date_str):
    '''returns next thursday from given date'''
    parts = date_str if isinstance(
        date_str, date) else datetime.strptime(date_str, '%Y.%m.%d')
    delta = 4-parts.isoweekday() if parts.isoweekday() <= 4 else 7+4-parts.isoweekday()

#     if past:
#         delta -=7

    parts = parts+timedelta(days=delta)
    parts = parts.strftime('%Y.%m.%d')

    return parts


def test(base_date, comparision_date):
    '''
    take:
    base_date - requested date
    comparision_date - confirmed date
    returns:
    is_overdue - 
    adj_base_date - base_date changed to nearest thursday in future
    '''
    base_date = pd.to_datetime(
        base_date, yearfirst=True, errors='coerce').dt.date
    comparision_date = pd.to_datetime(
        comparision_date, yearfirst=True, errors='coerce').dt.date

    adj_base_date = base_date.apply(lambda x:
                                    4-x.isoweekday() if x.isoweekday() <= 4 else 7+4-x.isoweekday())
#     adj_base_date = base_date + pd.Timedelta(adj_base_date, unit='d')
    adj_base_date = base_date + \
        adj_base_date.apply(lambda x: pd.Timedelta(x, unit='d'))
#     delta = 4-base_date.isoweekday() if base_date.isoweekday()<=4 else 7+4-base_date.isoweekday()
#     adj_base_date = base_date+timedelta(days=delta)

    is_overdue = comparision_date > adj_base_date

#     adj_base_date = adj_base_date.dt.strftime('%Y.%m.%d')
#     adj_base_date = adj_base_date.strftime('%Y.%m.%d')

    return is_overdue, adj_base_date


df_cur_conf[cur_conf_col] = df_cur_conf['Confirmed_Date'].astype('str') + \
    "_" + df_cur_conf['Confirmed_Quantity'].astype('str') + "/" \
    + (df_cur_conf['Order_Quantity'] -
       df_cur_conf['Delivery_Quantity']).astype('str')

# already confirmed but not taken by customers
df_forgoten = df_cur_conf[(df_cur_conf['Delivery_Quantity'] == 0) &
                          (df_cur_conf['Confirmed_Date'] < df_cur_conf['Created_on'].max())]

df_cur_conf = df_cur_conf.drop(columns=['Sold-To_Party',
                                        'Sales_document', 'SD_Document_Category', 'Customer_Reference',
                                        'Sales_Document_Item', 'Material', 'Material_text', 'Standard_Lead_Time', 'MRP_Controller',
                                        'Created_on', 'Order_Quantity', 'Delivery_Quantity'])


print(
    f'There are {df_part_delivered.shape[0]} lines partially delivered since last session')
# print('Here is a list of current confirmations')
# print(df_cur_conf.head())

if df_forgoten.shape[0]:
    print(f'There are {df_forgoten.shape[0]} lines in "forgotten" orders')
    print(df_forgoten.head())


def columns_to_drop(df, start_index=0, pattern='Confirmation_'):
    columns = df.columns

    drop_list = df.columns[start_index:]
    for stop_index, value in enumerate(columns):
        if pattern in value:
            drop_list = df.columns[start_index:stop_index]
            break

    return drop_list


# df_conf, sheets = read_in_file(file=confirmation_file['file'], folder=file_folder, header_row=0)
# # print (sheets)

# if df_conf.size == 0:
#     file = os.path.join(file_folder, confirmation_file['file'])

#     with pd.ExcelWriter(file, mode='w') as writer:
#         print(f"Creating new {confirmation_file['file']} file\n")
#         df_cur_conf.to_excel(writer, sheet_name=cur_conf_col, index=False)
#         writer.save()
# else:
#     df_conf.set_index('item_index', drop=True, inplace=True)
# #     df_conf.pop('changed')

#     df_conf = df_conf.drop(columns=columns_to_drop(
#         df_conf, start_index=1, pattern='Confirmation_'))

# # check if column with current date already exists and replace it
# if cur_conf_col in df_conf.columns:
#     print(
#         f'\ncolumn {cur_conf_col} exists in file {confirmation_file}. Replacing it with new data')
#     df_conf.pop(cur_conf_col)

# print(df_conf.head())
