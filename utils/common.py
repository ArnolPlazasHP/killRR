import pandas as pd
import json
from datetime import datetime


def get_path(source, type, name):
    # Opening JSON file
    DB_FILE = open('./data.json')
    # returns JSON object as a dictionary
    DATA = json.load(DB_FILE)   
    DB_FILE.close()
    for db, value in DATA[source][type].items():
        if db == name:
            settings = value
            break
        else:
            raise KeyError('DB not exist')
    return settings



def establish_range(country='all'):
    PATH = get_path('INPUT', 'FILES', 'CUSRES')
    df_ww_cusres = pd.read_excel(PATH, skiprows=1, sheet_name='Sheet1')
    df_ww_cusres.columns = ['Vendor Code', 'Import Export Flag', 'IOR # (for imports only)', 'EOR # (for exports only)', 'Country of Departure', 'Country of Destinatn', 'Place of Importation', 'BW Record Population Date', 'Entry Number', 'Entry Date', 'Import Declaration Time', 'VAT ID of Importer', 'VAT ID of Exporter', 'End Usr Name Add 1', 'End Usr Name Add 2', 'End Usr Name Add 3', 'End Usr Name Add 4', 'Division', 'IC Invoice Number', 'Invoice Line Item Number', 'HP Product Number', 'Z description', "Product's Country of Origin (CIC) in ISO Code", 'Sales Order - CusRes', 'Lagacy Order Number', 'PO number (IC)', 'TT5C Shipref', 'Bill of Lading Nr.', 'House Airway Bill', 'Master Airway Bill Number', 'Incoterms', 'Entry Value Currency', 'HTS assigned', 'Total statistical vl', 'Invoice total value', 'Statistical value cu', 'Qty in Doc. UoM', 'Exchange Rate (KF)', 'Duty Charge Amount', 'Import VAT Tax amount', 'Import VAT Tax Rate', 'Item Total Amount', 'Total Logistics Charge', 'Unit Price', 'USD converted Import VAT Tax Amount', 'Harbor Maintence Fee (Amount)', 'Duty Rate']
    
    df_ww_cusres = df_ww_cusres[['Country of Destinatn', 'Entry Date']]
    df_ww_cusres = df_ww_cusres.fillna(method="ffill")

    if country != 'all':
        df_ww_cusres = df_ww_cusres[df_ww_cusres['Entry Date'] == country]
    
    df_ww_cusres['Entry Date'] = pd.to_datetime(df_ww_cusres['Entry Date'], format='%d.%m.%Y')

    df_ww_cusres['month'] = df_ww_cusres['Entry Date'].dt.month
    list_month = list(df_ww_cusres['month'].value_counts().reset_index()['index'])
    if list_month == [12, 1]:
        max_month = 1
    else:
        max_month = df_ww_cusres['month'].max()

    df_ww_cusres = df_ww_cusres[df_ww_cusres['month'] == max_month]
    first_day = df_ww_cusres['Entry Date'].min()
    today = df_ww_cusres['Entry Date'].max()

    
    return first_day, today

# if __name__ == '__main__':
#     today, first_day = establish_range()
#     print(today, first_day)
#     print(type(today), type(first_day))
