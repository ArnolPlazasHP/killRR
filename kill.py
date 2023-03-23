import pandas as pd 
import numpy as np
import pyodbc

from datetime import date, datetime
import datetime as dt

from bw import bw


COUNTRIES = {'Peru': ['PE', 'PER'], 'Colombia': ['CO'], 'Mexico': ['MX'], 'Argentina': ['AR']}

def get_cusres(country, df_ww_cusres):
    df_ww_cusres = df_ww_cusres[df_ww_cusres['Country of Destinatn'].isin([country])]
    df_ww_cusres = df_ww_cusres.convert_dtypes()
    df_ww_cusres.fillna('-', inplace=True)
    df_ww_cusres['HP Product Number'] = df_ww_cusres['HP Product Number'].apply(lambda x : x[:x.rfind(' ')] if x.rfind(' ') != -1 else x)
    f = list(df_ww_cusres['IC Invoice Number'].drop_duplicates())
    df_ww_cusres['key'] = df_ww_cusres['IC Invoice Number'] + df_ww_cusres['HP Product Number']
    return df_ww_cusres, f

def get_gts(country, df_gts, f):
    df_gts = df_gts[df_gts['SHIP_TO_COUNTRY'].isin([country])]
    df_gts =df_gts.convert_dtypes()
    df_gts['PRODUCT_NO'] = df_gts['PRODUCT_NO'].apply(lambda x : x[:x.rfind('#')] if x.rfind('#') != -1 else x)
    df_gts = df_gts[df_gts['INVOICE_NUMBER'].isin(f)]
    df_gts['key'] = df_gts['INVOICE_NUMBER'] + df_gts['PRODUCT_NO']
    df_gts = df_gts[df_gts['INVOICE_DT'].str.contains('202107')]
    return df_gts

def get_db_additional(country):
    df_equivalent = pd.read_excel('input/TABLA DE EQUIVALENCIAS.xlsx')
    df_duty = pd.read_excel('input/Duty Tariff MCA.xlsx', sheet_name=country)
    df_duty['HTS'] = df_duty['HTS'].astype('string')
    df_duty.dropna(inplace=True)
    if country == 'PE':
        df_duty['HTS'] = df_duty['HTS'].apply(lambda x : x[:8])
    

    return df_equivalent, df_duty


def compare(df_ww_cusres, df_gts, country, df_equivalent, df_duty):
    df_final = df_ww_cusres.merge(df_gts, on= 'key', how='right')
    df_final = df_final[['SHIP_TO_COUNTRY', 'INVOICE_DT', 'INVOICE_NUMBER', 'IC Invoice Number', 'PRODUCT_NO' ,'HP Product Number', 'COO', "Product's Country of Origin (CIC) in ISO Code", 'TERMS_OF_DELIVERY', 'Incoterms', 'CURRENCY', 'Entry Value Currency', 'HTS', 'HTS assigned', 'LINE_AMOUNT', 'Total statistical vl', 'TOTAL_INV_AMT', 'Invoice total value', 'LINE_QTY', 'Qty in Doc. UoM', 'Duty Charge Amount','Import VAT Tax amount', 'Import VAT Tax Rate', 'Item Total Amount', 'UNIT_PRICE', 'Unit Price', 'USD converted Import VAT Tax Amount', 'Duty Rate', 'Harbor Maintence Fee (Amount)', 'Entry Number', 'Exchange Rate (KF)']]
    
    # tranform Product's Country of Origin (CIC) in ISO Code
    df_final = df_final.merge(df_equivalent, on="Product's Country of Origin (CIC) in ISO Code", how ='left')
    df_final["Product's Country of Origin (CIC) in ISO Code"] = df_final["COO GTS"]
    df_final = df_final.drop(['COO GTS'], axis=1)

    # CURRENCY
    df_final = df_final.astype('string')
    df_final.fillna('-', inplace=True)
    df_final['CURRENCY'] = df_final['CURRENCY'].apply(lambda x: x.strip())
    df_final['CURRENCY'] = df_final['CURRENCY'].apply(lambda x: 'USD' if x == 'US Dollar' else x)

    # HTS
    if country == 'PE':
        df_final['HTS assigned'] = df_final['HTS assigned'].apply(lambda x : x[:8])
        df_final['HTS'] = df_final['HTS'].apply(lambda x : x[:8])
    elif country == 'AR':
        df_final['HTS assigned'] = df_final['HTS assigned'].str.replace('.', '')

    df_final['HTS'] = df_final['HTS'].astype('string')
    df_final = df_final.merge(df_duty, on='HTS', how='left')


    first_column = df_final.pop('Entry Number')
    df_final.insert(0, 'Entry Number', first_column)
    df_final.loc[:,'LINE_AMOUNT':] = df_final.loc[:,'LINE_AMOUNT':].apply(pd.to_numeric, errors='coerce')
    df_final.loc[:, :'HTS assigned'] = df_final.loc[:, :'HTS assigned'].astype('string')
    df_final['Duty Rate'] = df_final['Duty Rate'].apply(lambda x : x/100)
    df_final['Import VAT Tax Rate'] = df_final['Import VAT Tax Rate'].apply(lambda x : x/100)
    df_final['VALIDATION DUTY'] = (df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * df_final['DUTY']
    if country == 'CO':
        df_final['VALIDATION VAT'] = ((df_final['VALIDATION DUTY'] + df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * df_final['VAT']) * df_final['Exchange Rate (KF)']
    else:
        df_final['VALIDATION VAT'] = (df_final['VALIDATION DUTY'] + df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * df_final['VAT']
    df_final.loc[:,'LINE_AMOUNT':] = df_final.loc[:,'LINE_AMOUNT':].apply(lambda x: round(x, 2))

    df_final.fillna('-', inplace=True)
    df_final['COO'] = df_final['COO'].replace('', '-')
    df_final['HTS'] = df_final['HTS'].replace('', '-')
    
    df_final['COMPARE_PRODUCT_NUMBER'] = df_final.apply(lambda x: compare_columns(x['PRODUCT_NO'], x['HP Product Number']), axis=1)
    df_final['COMPARE_COUNTRY_ORIGIN'] = df_final.apply(lambda x: compare_columns(x['COO'], x["Product's Country of Origin (CIC) in ISO Code"]), axis=1)
    df_final['COMPARE_INCOTERMS'] = df_final.apply(lambda x: compare_columns(x['TERMS_OF_DELIVERY'], x['Incoterms']), axis=1)
    df_final['COMPARE_CURRENCY'] = df_final.apply(lambda x: compare_columns(x['CURRENCY'], x['Entry Value Currency']), axis=1)
    df_final['COMPARE_HTS'] = df_final.apply(lambda x: compare_columns(x['HTS'], x['HTS assigned']), axis=1)
    df_final['COMPARE_DUTY_RATE'] = df_final.apply(lambda x: compare_columns(x['Duty Rate'], x['VALIDATION DUTY']), axis=1)
    df_final['COMPARE_DUTY'] = df_final.apply(lambda x: compare_columns(x['Duty Charge Amount'], x['DUTY']), axis=1)
    df_final['COMPARE_LINE_AMOUNT'] = df_final.apply(lambda x: compare_columns(x['Item Total Amount'], x['LINE_AMOUNT']), axis=1)
    df_final['COMPARE_LINE_QTY'] = df_final.apply(lambda x: compare_columns(x['Qty in Doc. UoM'], x['LINE_QTY']), axis=1)
    df_final['COMPARE_VAT_RATE'] = df_final.apply(lambda x: compare_columns(x['Import VAT Tax Rate'], x['VAT']), axis=1)
    df_final['COMPARE_VAT'] = df_final.apply(lambda x: compare_columns(x['Import VAT Tax amount'], x['VALIDATION VAT']), axis=1)
    df_final['COMPARE_UNIT_PRICE'] = df_final.apply(lambda x: compare_columns(x['Unit Price'], x['UNIT_PRICE']), axis=1)
    df_final['COMPARE_TOTAL_INV_AMT'] = df_final.apply(lambda x: compare_columns(x['Invoice total value'], x['TOTAL_INV_AMT']), axis=1)

    df_final.rename(columns={'Entry Number': 'ENTRY_NUMBER', 'INVOICE_NUMBER': 'INVOICE_NUMBER_GTS', 'IC Invoice Number': 'INVOICE_NUMBER_CUSRES', 'PRODUCT_NO': 'PRODUCT_NUMBER_GTS', 'HP Product Number': 'PRODUCT_NUMBER_CUSRES', 'COO': 'COUNTRY_ORIGIN_GTS', "Product's Country of Origin (CIC) in ISO Code": 'COUNTRY_ORIGIN_CUSRES', 'TERMS_OF_DELIVERY': 'INCOTERMS_GTS', 'Incoterms': 'INCOTERMS_CUSRES', 'CURRENCY': 'CURRENCY_GTS', 'Entry Value Currency': 'CURRENCY_CUSRES', 'HTS': 'HTS_GTS', 'HTS assigned': 'HTS_CUSRES', 'VALIDATION DUTY': 'DUTY_GTS', 'Duty Charge Amount': 'DUTY_CUSRES', 'VALIDATION VAT': 'VAT_GTS', 'Import VAT Tax amount': 'VAT_CUSRES', 'LINE_AMOUNT': 'LINE_AMOUNT_GTS', 'Item Total Amount': 'LINE_AMOUNT_CUSRES', 'LINE_QTY': 'LINE_QTY_GTS', 'Qty in Doc. UoM': 'LINE_QTY_CUSRES', 'VAT': 'VAT_RATE_GTS', 'Import VAT Tax Rate': 'VAT_RATE_CUSRES', 'UNIT_PRICE': 'UNIT_PRICE_GTS', 'Unit Price': 'UNIT_PRICE_CUSRES', 'DUTY': 'DUTY_RATE_GTS', 'Duty Rate': 'DUTY_RATE_CUSRES', 'TOTAL_INV_AMT': 'TOTAL_INV_AMT_GTS', 'Invoice total value': 'TOTAL_INV_AMT_CUSRES', 'Total statistical vl': 'TOTAL_STATISTICAL_CUSRES', 'Harbor Maintence Fee (Amount)': 'HARBOR MAINTENCE FEE (AMOUNT)'}, inplace=True)

    df_final = df_final[['SHIP_TO_COUNTRY', 'ENTRY_NUMBER', 'INVOICE_DT','INVOICE_NUMBER_GTS','INVOICE_NUMBER_CUSRES','PRODUCT_NUMBER_GTS','PRODUCT_NUMBER_CUSRES','COMPARE_PRODUCT_NUMBER','COUNTRY_ORIGIN_GTS','COUNTRY_ORIGIN_CUSRES','COMPARE_COUNTRY_ORIGIN','INCOTERMS_GTS','INCOTERMS_CUSRES','COMPARE_INCOTERMS','CURRENCY_GTS','CURRENCY_CUSRES','COMPARE_CURRENCY','HTS_GTS','HTS_CUSRES','COMPARE_HTS','LINE_AMOUNT_GTS','LINE_AMOUNT_CUSRES','COMPARE_LINE_AMOUNT','TOTAL_INV_AMT_GTS','TOTAL_INV_AMT_CUSRES','COMPARE_TOTAL_INV_AMT','TOTAL_STATISTICAL_CUSRES','DUTY_RATE_GTS','DUTY_RATE_CUSRES','COMPARE_DUTY_RATE', 'DUTY_GTS','DUTY_CUSRES','COMPARE_DUTY', 'VAT_RATE_GTS', 'VAT_RATE_CUSRES', 'COMPARE_VAT_RATE', 'VAT_GTS', 'VAT_CUSRES', 'COMPARE_VAT','LINE_QTY_GTS','LINE_QTY_CUSRES','COMPARE_LINE_QTY','UNIT_PRICE_GTS','UNIT_PRICE_CUSRES','COMPARE_UNIT_PRICE','HARBOR MAINTENCE FEE (AMOUNT)'
]]
    return df_final


def compare_columns(c1, c2):
    if c1 == '-' or c2 == '-':
        return 'Blanks'
    elif c1 == c2:
        return 'Equal'
    else:
        return 'Different'


def run():
    server = 'DERUEDAL6\MSSQL11INC' 
    database = 'GTS_INV_MCA' 
    username = 'sa' 
    password = 'Hpadmin123*' 
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    query = "SELECT * FROM GTS_Report;"
    df_gts = pd.read_sql(query, cnxn)

    yesterday = (date.today() - dt.timedelta(days=1)).strftime('%d.%m.%Y')
    thirty_days_before = (date.today() - dt.timedelta(days=30)).strftime('%d.%m.%Y')
    #bw(thirty_days_before, yesterday)
    df_ww_cusres = pd.read_excel(r'./input/WW_CUSRES.xlsx', skiprows=1, sheet_name='Sheet1')
    df_ww_cusres.columns = ['Vendor Code','IOR # (for imports only)','EOR # (for exports only)','Country of Departure','Country of Destinatn','Place of Importation','BW Record Population Date','Entry Number','Entry Date','Import Declaration Time','VAT ID of Importer','VAT ID of Exporter','End Usr Name Add 1','End Usr Name Add 2','End Usr Name Add 3','End Usr Name Add 4','Division','IC Invoice Number','Invoice Line Item Number','HP Product Number',"Product's Country of Origin (CIC) in ISO Code",'Sales Order - CusRes','Lagacy Order Number','PO number (IC)','TT5C Shipref','Bill of Lading Nr.','House Airway Bill','Master Airway Bill Number','Incoterms','Entry Value Currency','HTS assigned','Total statistical vl','Invoice total value','Statistical value cu','Qty in Doc. UoM','Exchange Rate (KF)','Duty Charge Amount','Import VAT Tax amount','Import VAT Tax Rate','Item Total Amount','Total Logistics Charge','Unit Price','USD converted Import VAT Tax Amount','Harbor Maintence Fee (Amount)','Duty Rate']
    df_ww_cusres = df_ww_cusres[['IC Invoice Number','Country of Destinatn','HP Product Number',"Product's Country of Origin (CIC) in ISO Code",'Incoterms','Entry Value Currency','HTS assigned','Total statistical vl', 'Invoice total value','Statistical value cu','Qty in Doc. UoM','Exchange Rate (KF)','Duty Charge Amount','Import VAT Tax amount','Import VAT Tax Rate','Item Total Amount','Total Logistics Charge','Unit Price','USD converted Import VAT Tax Amount','Harbor Maintence Fee (Amount)','Duty Rate', 'Entry Number']]
    df_ww_cusres['IC Invoice Number'] = df_ww_cusres['IC Invoice Number'].fillna(method="ffill")
    df_ww_cusres['Country of Destinatn'] = df_ww_cusres['Country of Destinatn'].fillna(method="ffill")
    
    for x, y in COUNTRIES.items():
        if x == 'Peru':
            df_ww_cusres, f = get_cusres(y[1], df_ww_cusres)
        else:
            df_ww_cusres, f = get_cusres(y[0], df_ww_cusres)
        df_gts = get_gts(y[0], df_gts, f)
        df_equivalent, df_duty = get_db_additional(y[0])
        df_final = compare(df_ww_cusres, df_gts, y[0], df_equivalent, df_duty)
        with pd.ExcelWriter(f'output/Import Declaration Reconciliation Report {x}.xlsx',index=False) as writer:
            df_final.to_excel(writer, index=False)



if __name__ == '__main__':
    run()

