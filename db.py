import math
import shutil

import pandas as pd
import pyodbc

from robot_class import get_information_class
from utils.common import establish_range


def clean_class_file(df):
    df = pd.DataFrame(df[0])
    df.drop([0], inplace=True)
    df = df.loc[:,[0, 6]]
    df.rename(columns={0:'PRODUCT_NO', 6:'HTS Class'}, inplace=True)
    df.dropna(subset=['PRODUCT_NO'], inplace=True)
    df = df.convert_dtypes()
    df['PRODUCT_NO'] = df['PRODUCT_NO'].apply(lambda x : x[:x.rfind('#')] if x.rfind('#') != -1 else x)
    df['PRODUCT_NO'] = df['PRODUCT_NO'].apply(lambda x : x[:x.rfind(' ')] if x.rfind(' ') != -1 else x)
    return df


def call_class_robot(df_gts, country):
    list_part_number = list(df_gts['PRODUCT_NO'].drop_duplicates())
    l = len(list_part_number)
    g = math.ceil(l/400)
    df_class_consolidate = pd.DataFrame()
    for i in range(g):
        initial = i * 400
        final = ((i + 1) * 400) -1
        final = l if (final > l) else final
        list_pn = list_part_number[initial:final]
        get_information_class(list_pn, country)
        shutil.move('C:\\Users\\PlazasAr\\Downloads\\SYN_CLS_PART_LOOKUP_MULTI.part_lookup_multi_export', f'./input/look_up{i}.xls')
        df_SYN_CLS = pd.read_html(f'./input/look_up{i}.xls')
        df_SYN_CLS = clean_class_file(df_SYN_CLS)
        df_class_consolidate = pd.concat([df_class_consolidate, df_SYN_CLS])
    return df_class_consolidate


def join_data(x):
    return ', '.join(x)


def run():
    COUNTRY = ['AR', 'CL', 'CO', 'MX', 'PE']
    
    first_day, today = establish_range()

    df_ww_cusres = pd.read_excel('./input/WW_CUSRES.xlsx', skiprows=1, sheet_name='Sheet1')
    df_ww_cusres.columns = ['Vendor Code', 'Import Export Flag', 'IOR # (for imports only)', 'EOR # (for exports only)', 'Country of Departure', 'Country of Destinatn', 'Place of Importation', 'BW Record Population Date', 'Entry Number', 'Entry Date', 'Import Declaration Time', 'VAT ID of Importer', 'VAT ID of Exporter', 'End Usr Name Add 1', 'End Usr Name Add 2', 'End Usr Name Add 3', 'End Usr Name Add 4', 'Division', 'IC Invoice Number', 'Invoice Line Item Number', 'HP Product Number', 'Z description', "Product's Country of Origin (CIC) in ISO Code", 'Sales Order - CusRes', 'Lagacy Order Number', 'PO number (IC)', 'TT5C Shipref', 'Bill of Lading Nr.', 'House Airway Bill', 'Master Airway Bill Number', 'Incoterms', 'Entry Value Currency', 'HTS assigned', 'Total statistical vl', 'Invoice total value', 'Statistical value cu', 'Qty in Doc. UoM', 'Exchange Rate (KF)', 'Duty Charge Amount', 'Import VAT Tax amount', 'Import VAT Tax Rate', 'Item Total Amount', 'Total Logistics Charge', 'Unit Price', 'USD converted Import VAT Tax Amount', 'Harbor Maintence Fee (Amount)', 'Duty Rate']
    
    df_ww_cusres = df_ww_cusres[['IC Invoice Number','Country of Destinatn','HP Product Number', 'End Usr Name Add 1', 'Entry Date', "Product's Country of Origin (CIC) in ISO Code",'Incoterms','Entry Value Currency','HTS assigned','Total statistical vl', 'Invoice total value','Statistical value cu','Qty in Doc. UoM','Exchange Rate (KF)','Duty Charge Amount','Import VAT Tax amount','Import VAT Tax Rate','Item Total Amount','Total Logistics Charge','Unit Price','USD converted Import VAT Tax Amount','Harbor Maintence Fee (Amount)','Duty Rate', 'Entry Number', 'House Airway Bill', 'TT5C Shipref']]

    df_ww_cusres['IC Invoice Number'] = df_ww_cusres['IC Invoice Number'].fillna(method="ffill")
    df_ww_cusres['Country of Destinatn'] = df_ww_cusres['Country of Destinatn'].fillna(method="ffill")
    df_ww_cusres['Entry Number'] = df_ww_cusres['Entry Number'].fillna(method="ffill")
    df_ww_cusres['Entry Date'] = df_ww_cusres['Entry Date'].fillna(method="ffill")
    df_ww_cusres['End Usr Name Add 1'] = df_ww_cusres['End Usr Name Add 1'].fillna(method="ffill")
    df_ww_cusres['Entry Date'] = pd.to_datetime(df_ww_cusres['Entry Date'], format='%d.%m.%Y')

    df_equivalent = pd.read_excel('input/TABLA DE EQUIVALENCIAS.xlsx')
    df_ww_cusres = df_ww_cusres.merge(df_equivalent, on="Product's Country of Origin (CIC) in ISO Code", how ='left')
    df_ww_cusres["Product's Country of Origin (CIC) in ISO Code"] = df_ww_cusres["COO GTS"]
    df_ww_cusres = df_ww_cusres.drop(['COO GTS'], axis=1)

    mask = (df_ww_cusres['Entry Date'] >= first_day) & (df_ww_cusres['Entry Date'] <= today)
    df_ww_cusres_dq = df_ww_cusres.loc[mask]
    print(today, first_day)
    f = list(df_ww_cusres_dq['IC Invoice Number'].drop_duplicates())


    server = 'DERUEDAL25,1433' 
    database = 'GTS_INV_MCA' 
    username = 'sa' 
    password = 'Hpadmin123*' 
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    query = "SELECT * FROM GTS_Report;"

    df_gts = pd.read_sql(query, cnxn)
    df_gts =df_gts.convert_dtypes()
    df_gts['LINE_QTY'] = df_gts['LINE_QTY'].str.strip()
    df_gts['LINE_QTY'] = df_gts['LINE_QTY'].apply(lambda x : x[:x.rfind('.')] if x.rfind('.') != -1 else x)

    df_gts_qty = df_gts[df_gts['SHIP_TO_COUNTRY'].isin(COUNTRY)] 
    df_gts_qty['PRODUCT_NO'] = df_gts_qty['PRODUCT_NO'].apply(lambda x : x[:x.rfind('#')] if x.rfind('#') != -1 else x)
    df_gts_qty['PRODUCT_NO'] = df_gts_qty['PRODUCT_NO'].apply(lambda x : x[:x.rfind(' ')] if x.rfind(' ') != -1 else x)
    df_gts_dq = df_gts_qty[df_gts_qty['INVOICE_NUMBER'].isin(f)]
    
    df_gts_dq_class = pd.DataFrame()
    for c in COUNTRY:
        print(c)
        try:
            df_gts_country = df_gts_dq[df_gts_dq['SHIP_TO_COUNTRY'] == c]
            print(df_gts_country.shape) 
            df_class_consolidate = call_class_robot(df_gts_country, c)
            df_gts_country = df_gts_country.merge(df_class_consolidate, on='PRODUCT_NO', how='left')
            df_gts_dq_class = pd.concat([df_gts_dq_class, df_gts_country])
        except:
            print('There is no data')

    print(df_gts_dq.shape, df_gts_dq_class.shape)
    
    # create inputs files
       
    with pd.ExcelWriter('./input/GTS_qty.xlsx') as writer:
        df_gts_qty.to_excel(writer, sheet_name='GTS', index=False)
    with pd.ExcelWriter('./input/GTS_dq.xlsx') as writer:
        df_gts_dq_class.to_excel(writer, sheet_name='GTS', index=False)
    with pd.ExcelWriter('./input/cusres_qty.xlsx') as writer:
        df_ww_cusres.to_excel(writer, sheet_name='cusres', index=False)
    with pd.ExcelWriter('./input/cusres_dq.xlsx') as writer:
        df_ww_cusres_dq.to_excel(writer, sheet_name='cusres', index=False)

if __name__ == '__main__':
    run()
