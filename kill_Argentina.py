import pandas as pd 
import win32com.client as client
import pathlib

from utils.common import establish_range

def get_cusres(country, df_ww_cusres):
    df_ww_cusres = df_ww_cusres[df_ww_cusres['Country of Destinatn'].isin([country])]
    df_ww_cusres = df_ww_cusres.convert_dtypes()
    df_ww_cusres['HP Product Number'].fillna('Blank', inplace=True)
    df_ww_cusres['HP Product Number'] = df_ww_cusres['HP Product Number'].apply(lambda x : x[:x.rfind('#')] if x.rfind('#') != -1 else x)
    df_ww_cusres['HP Product Number'] = df_ww_cusres['HP Product Number'].apply(lambda x : x[:x.rfind(' ')] if x.rfind(' ') != -1 else x)
    df_ww_cusres['IC Invoice Number'] = df_ww_cusres['IC Invoice Number'].astype('string')
    df_ww_cusres['key'] = df_ww_cusres['IC Invoice Number'] + df_ww_cusres['HP Product Number']
    return df_ww_cusres

def get_gts(country, df_gts):
    df_gts = df_gts[df_gts['SHIP_TO_COUNTRY'].isin([country])]
    df_gts =df_gts.convert_dtypes()
    df_gts['INVOICE_NUMBER'] = df_gts['INVOICE_NUMBER'].astype('string')
    df_gts['key'] = df_gts['INVOICE_NUMBER'] + df_gts['PRODUCT_NO']
    return df_gts

def get_db_additional(country):
    df_duty = pd.read_excel('../../PBI/Kill Random Review/Duty Tariff MCA.xlsx', sheet_name=country)
    df_duty = df_duty.convert_dtypes()
    df_duty['HTS'] = df_duty['HTS'].astype('string')
    df_duty.dropna(inplace=True)

    return df_duty


def compare(df_final, df_duty):

    # CURRENCY
    df_final = df_final.astype('string')
    df_final['CURRENCY'].fillna('Blank', inplace=True)
    df_final['CURRENCY'] = df_final['CURRENCY'].apply(lambda x: x.strip())
    df_final['CURRENCY'] = df_final['CURRENCY'].apply(lambda x: 'USD' if x == 'US Dollar' else x)

    # HTS
    df_final['HTS assigned'] = df_final['HTS assigned'].str.replace('.', '')

    df_final['HTS'] = df_final['HTS'].astype('string')
    df_final = df_final.merge(df_duty, on='HTS', how='left')
    duty_84714900200J = list(df_duty[df_duty['HTS'] == '84714900200J']['DUTY'])[0]
    vat_84714900200J = list(df_duty[df_duty['HTS'] == '84714900200J']['VAT'])[0]

    first_column = df_final.pop('Entry Number')
    df_final.insert(0, 'Entry Number', first_column)
    df_final.loc[:,'LINE_AMOUNT':] = df_final.loc[:,'LINE_AMOUNT':].apply(pd.to_numeric, errors='coerce')
    df_final.loc[:, :'HTS assigned'] = df_final.loc[:, :'HTS assigned'].astype('string')
    df_final['Duty Rate'] = df_final['Duty Rate'].apply(lambda x : x/100)
    df_final['Import VAT Tax Rate'] = df_final['Import VAT Tax Rate'].apply(lambda x : x/100)
    df_final['VALIDATION DUTY'] = (df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * df_final['DUTY']
    df_final.loc[(df_final['HTS'] == '85285220000F') & (df_final['HTS assigned'] == '84714900200J'), 'VALIDATION DUTY'] = (df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * duty_84714900200J
    df_final['VALIDATION VAT'] = (df_final['VALIDATION DUTY'] + df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * df_final['VAT']
    df_final.loc[(df_final['HTS'] == '85285220000F') & (df_final['HTS assigned'] == '84714900200J'), 'VALIDATION VAT'] = (df_final['VALIDATION DUTY'] + df_final['LINE_AMOUNT'] + df_final['Harbor Maintence Fee (Amount)']) * vat_84714900200J

    df_final.loc[:, ['LINE_AMOUNT', 'Total statistical vl', 'TOTAL_INV_AMT',
       'Invoice total value', 'LINE_QTY', 'Qty in Doc. UoM',
       'Duty Charge Amount', 'Import VAT Tax amount',
       'Item Total Amount', 'UNIT_PRICE', 'Unit Price',
       'USD converted Import VAT Tax Amount',
       'Harbor Maintence Fee (Amount)', 'Exchange Rate (KF)',
       'VALIDATION DUTY', 'VALIDATION VAT']] = df_final.loc[:,
                                                   ['LINE_AMOUNT', 'Total statistical vl', 'TOTAL_INV_AMT',
       'Invoice total value', 'LINE_QTY', 'Qty in Doc. UoM',
       'Duty Charge Amount', 'Import VAT Tax amount',
       'Item Total Amount', 'UNIT_PRICE', 'Unit Price',
       'USD converted Import VAT Tax Amount',
       'Harbor Maintence Fee (Amount)', 'Exchange Rate (KF)',
       'VALIDATION DUTY', 'VALIDATION VAT']].apply(lambda x: round(x, 1))
    
    df_final = df_final.astype('string')
    df_final.fillna('Blank', inplace=True)
    df_final['LINE_QTY'] = df_final['LINE_QTY'].apply(lambda x : x[:x.rfind('.')] if x.rfind('.') != -1 else x)
    df_final['Qty in Doc. UoM'] = df_final['Qty in Doc. UoM'].apply(lambda x : x[:x.rfind('.')] if x.rfind('.') != -1 else x)
    df_final['COO'] = df_final['COO'].replace('', 'Blank')
    df_final['HTS'] = df_final['HTS'].replace('', 'Blank')
    
    df_final['COMPARE_PRODUCT_NUMBER'] = df_final.apply(lambda x: compare_columns(x['PRODUCT_NO'], x['HP Product Number']), axis=1)
    df_final['COMPARE_COUNTRY_ORIGIN'] = df_final.apply(lambda x: compare_columns(x['COO'], x["Product's Country of Origin (CIC) in ISO Code"]), axis=1)
    df_final['COMPARE_INCOTERMS'] = df_final.apply(lambda x: compare_columns(x['TERMS_OF_DELIVERY'], x['Incoterms']), axis=1)
    df_final['COMPARE_CURRENCY'] = df_final.apply(lambda x: compare_columns(x['CURRENCY'], x['Entry Value Currency']), axis=1)
    df_final['COMPARE_HTS'] = df_final.apply(lambda x: compare_columns(x['HTS'], x['HTS assigned']), axis=1)
    df_final['COMPARE_DUTY_RATE'] = df_final.apply(lambda x: compare_columns(x['DUTY'], x['Duty Rate']), axis=1) 
    df_final['COMPARE_DUTY'] = df_final.apply(lambda x: compare_columns2(x['VALIDATION DUTY'], x['Duty Charge Amount']) if x['VALIDATION DUTY'] != 'Blank' and x['Duty Charge Amount'] != 'Blank' else compare_columns(x['VALIDATION DUTY'], x['Duty Charge Amount']), axis=1)
    df_final['COMPARE_LINE_AMOUNT'] = df_final.apply(lambda x: compare_columns2(x['LINE_AMOUNT'], x['Item Total Amount']), axis=1)
    df_final['COMPARE_LINE_QTY'] = df_final.apply(lambda x: compare_columns(x['LINE_QTY'], x['Qty in Doc. UoM']), axis=1)
    df_final['COMPARE_VAT_RATE'] = df_final.apply(lambda x: compare_columns(x['VAT'], x['Import VAT Tax Rate']), axis=1)
    df_final['COMPARE_VAT'] = df_final.apply(lambda x: compare_columns2(x['VALIDATION VAT'], x['Import VAT Tax amount']) if x['VALIDATION VAT'] != 'Blank' and x['Import VAT Tax amount'] != 'Blank' else compare_columns(x['VALIDATION VAT'], x['Import VAT Tax amount']) , axis=1)
    df_final['COMPARE_UNIT_PRICE'] = df_final.apply(lambda x: compare_columns2(x['UNIT_PRICE'], x['Unit Price']), axis=1)
    df_final['COMPARE_TOTAL_INV_AMT'] = df_final.apply(lambda x: compare_columns2(x['TOTAL_INV_AMT'], x['Invoice total value']), axis=1)
    df_final['COMPARE_HTS_CLASS'] = df_final.apply(lambda x: compare_columns4(x['HTS'], x['HTS Class']), axis=1) 

    df_final.rename(columns={'Entry Number': 'ENTRY_NUMBER', 'INVOICE_NUMBER': 'INVOICE_NUMBER_GTS', 'IC Invoice Number': 'INVOICE_NUMBER_CUSRES', 'PRODUCT_NO': 'PRODUCT_NUMBER_GTS', 'HP Product Number': 'PRODUCT_NUMBER_CUSRES', 'End Usr Name Add 1': 'END_USER', 'Entry Date': 'ENTRY_DATE', 'COO': 'COUNTRY_ORIGIN_GTS', "Product's Country of Origin (CIC) in ISO Code": 'COUNTRY_ORIGIN_CUSRES', 'TERMS_OF_DELIVERY': 'INCOTERMS_GTS', 'Incoterms': 'INCOTERMS_CUSRES', 'CURRENCY': 'CURRENCY_GTS', 'Entry Value Currency': 'CURRENCY_CUSRES', 'HTS': 'HTS_GTS', 'HTS assigned': 'HTS_CUSRES', 'VALIDATION DUTY': 'DUTY_GTS', 'Duty Charge Amount': 'DUTY_CUSRES', 'VALIDATION VAT': 'VAT_GTS', 'Import VAT Tax amount': 'VAT_CUSRES', 'LINE_AMOUNT': 'LINE_AMOUNT_GTS', 'Item Total Amount': 'LINE_AMOUNT_CUSRES', 'LINE_QTY': 'LINE_QTY_GTS', 'Qty in Doc. UoM': 'LINE_QTY_CUSRES', 'VAT': 'VAT_RATE_GTS', 'Import VAT Tax Rate': 'VAT_RATE_CUSRES', 'UNIT_PRICE': 'UNIT_PRICE_GTS', 'Unit Price': 'UNIT_PRICE_CUSRES', 'DUTY': 'DUTY_RATE_GTS', 'Duty Rate': 'DUTY_RATE_CUSRES', 'TOTAL_INV_AMT': 'TOTAL_INV_AMT_GTS', 'Invoice total value': 'TOTAL_INV_AMT_CUSRES', 'Total statistical vl': 'TOTAL_STATISTICAL_CUSRES', 'Harbor Maintence Fee (Amount)': 'HARBOR MAINTENCE FEE (AMOUNT)'}, inplace=True)

    df_final = df_final[['SHIP_TO_COUNTRY', 'ENTRY_NUMBER', 'END_USER', 'ENTRY_DATE','INVOICE_NUMBER_GTS','INVOICE_NUMBER_CUSRES','PRODUCT_NUMBER_GTS','PRODUCT_NUMBER_CUSRES','COMPARE_PRODUCT_NUMBER','COUNTRY_ORIGIN_GTS','COUNTRY_ORIGIN_CUSRES','COMPARE_COUNTRY_ORIGIN','INCOTERMS_GTS','INCOTERMS_CUSRES','COMPARE_INCOTERMS','CURRENCY_GTS','CURRENCY_CUSRES','COMPARE_CURRENCY','HTS_GTS','HTS_CUSRES','COMPARE_HTS','LINE_AMOUNT_GTS','LINE_AMOUNT_CUSRES','COMPARE_LINE_AMOUNT','TOTAL_INV_AMT_GTS','TOTAL_INV_AMT_CUSRES','COMPARE_TOTAL_INV_AMT','TOTAL_STATISTICAL_CUSRES','DUTY_RATE_GTS','DUTY_RATE_CUSRES','COMPARE_DUTY_RATE', 'DUTY_GTS','DUTY_CUSRES','COMPARE_DUTY', 'VAT_RATE_GTS', 'VAT_RATE_CUSRES', 'COMPARE_VAT_RATE', 'VAT_GTS', 'VAT_CUSRES', 'COMPARE_VAT','LINE_QTY_GTS','LINE_QTY_CUSRES','COMPARE_LINE_QTY','UNIT_PRICE_GTS','UNIT_PRICE_CUSRES','COMPARE_UNIT_PRICE','HARBOR MAINTENCE FEE (AMOUNT)', 'HTS Class', 'COMPARE_HTS_CLASS']]
    return df_final

def compare_columns(c1, c2):
    if c1 == 'DDP':
        return 'Error GTS'
    elif c1 == '85285220000F' and c2 == '84714900200J':
        return 'SYSTEM'
    elif c1 == '84715010900N' and c2 == '84714900200J':
        return 'SYSTEM'
    if c1 != 'Blank' and c2 == 'Blank':
        return 'Blanks in Cusres'
    elif c1 == 'Blank' and c2 != 'Blank':
        return 'Blanks in GTS'
    elif c1 == c2:
        return 'Equal'
    else:
        return 'Different'

def compare_columns2(c1, c2):
    if c1 == 0 and c2 != 0:
        return 'Zero in GTS'
    elif c1 != 'Blank' and c2 == 'Blank':
        return 'Blanks in Cusres'
    elif c1 == 'Blank' and c2 != 'Blank':
        return 'Blanks in GTS'
    if abs(c1 - c2) < 10:
        return 'Equal'
    else:
        return 'Different'

def compare_master_file(a, b):
    if a == 'Blank':
        r = 'MISSING IN GTS'
    elif b == 'Blank':
        r = 'MISSING IN CUSRES'
    elif a == b:
        r = 'EQUAL'
    else:
        r = 'DIFFERENT'
    return r

def compare_columns4(c1, c2):
    if c1 != 'Blank' and c2 == 'Blank':
        return 'Blanks in class'
    elif c1 == 'Blank' and c2 != 'Blank':
        return 'Blanks in GTS'
    elif c1 == c2:
        return 'Equal'
    else:
        return 'Different'

def send_email(list_criterio, list_qty_equal, list_qty_different, list_qty_blanks_cusres, week, range_date, total_rows, file_absolute_report, file_absolute_img):
    list_html = []
    for i, j, k, w in zip(list_criterio, list_qty_equal, list_qty_different, list_qty_blanks_cusres):
        s = f'<tr> <td>{i}</td> <td>{j}</td><td>{k}</td><td>{w}</td></tr>'
        list_html.append(s)
    s = ''.join(list_html)


    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "rodrigo.piris@expeditors.com; edgardo.abraham@expeditors.com; sebastian.smalc@expeditors.com; hugo.vignale@expeditors.com; karen.salviz@hp.com; rafael.patino@hp.com; karla.botello.tirado@hp.com"
    message.CC = "carolina-otero@hp.com"
    message.Subject = 'Discrepancias Declaraciones vs Documentos HP Semana ' + week + ' ' + range_date
    message.Attachments.Add(file_absolute_report)
    message.Attachments.Add(file_absolute_img)
    html_body = """
        <style>
            table, th, td {
                border: 1px solid #0096D6;
                border-collapse: collapse;
                padding: 0 5px 0 5px
            }
            th {
                color: white;
                background-color: #001373;
            }
        </style>
        Hola,<br>
        <br>
        <p>Hemos identificado las siguientes diferencias entre nuestros sistemas y sus datos (CUSRES), 
        por favor revisarlos y confirmar acción al equipo local de Customs operations de HP.</p>
        <br>
        <p>Adjunto la base de datos, por favor filtre por “Different”,  para conocer el detalle en cada Criterio: </p>
        <br>
        <p> <strong> TOTAL REGISTROS:   """ + total_rows + """</strong></p> 
        <br>
        <table>

        <tr>
            <th>CRITERIO</th>
            <th>QTY EQUAL</th>
            <th>QTY DIFFERENT</th>
            <th>QTY BLANKS IN CUSRES</th>
        </tr>     

        """ + s + "</table>"

    message.HTMLBody = html_body
    message.Save()
    message.Send()


def get_cusres_qty(df_ww_cusres, df_gts, country):
    first_day, today = establish_range()

    df_dmr = pd.read_excel('./input/DMR_LA_GTN.xlsx')
    df_dmr = df_dmr[['Country',	'CBL Number', 'HP Delivery No', 'RBHP', 'Incoterm']]
    df_dmr.rename(columns={'Country': 'DESTINATION COUNTRY', 'CBL Number': 'HAWB', 'HP Delivery No': 'OUTBOUND DELIVERY', 'RBHP': 'RECEIVED BY HP'}, inplace=True)
    df_dmr = df_dmr[~df_dmr['RECEIVED BY HP'].isnull()]
    df_dmr.drop_duplicates(inplace= True)
    df_dmr['RECEIVED BY HP'] = pd.to_datetime(df_dmr['RECEIVED BY HP'] ,format='%Y-%m-%d')
    mask = (df_dmr['RECEIVED BY HP'] >= first_day) & (df_dmr['RECEIVED BY HP'] <= today)
    df_dmr = df_dmr.loc[mask]
    df_dmr['DESTINATION COUNTRY'] = df_dmr['DESTINATION COUNTRY'].str.replace('Argentina', country)

    df_master_file = pd.read_excel('../../db/Shared Documents/- AIR CT MASTER FILES -/CT Multimodal Master File.xlsx')
    df_master_file = df_master_file[['DESTINATION COUNTRY',	'HAWB',	'OUTBOUND DELIVERY', 'RECEIVED BY HP', 'Incoterm']]
    df_master_file.drop_duplicates(inplace= True)
    df_master_file['RECEIVED BY HP'] = pd.to_datetime(df_master_file['RECEIVED BY HP'] ,format='%Y-%m-%d', errors='coerce')
    df_master_file.dropna(subset=['RECEIVED BY HP'], inplace=True)
    mask = (df_master_file['RECEIVED BY HP'] >= first_day) & (df_master_file['RECEIVED BY HP'] <= today)
    df_master_file = df_master_file.loc[mask]

    df_master_file = pd.concat([df_master_file, df_dmr])
    df_master_file = df_master_file[df_master_file['Incoterm'].isin(['DAP', 'DDP'])]
    df_master_file['OUTBOUND DELIVERY'] = df_master_file['OUTBOUND DELIVERY'].astype(str)
    df_master_file['OUTBOUND DELIVERY'] = df_master_file['OUTBOUND DELIVERY'].apply(lambda x: x[:-2] if x[-2:] == '.0' else x)
    df_master_file = df_master_file[df_master_file['DESTINATION COUNTRY'] == country]
    df_gts['SAP_DELIVERY_ID'] = df_gts['SAP_DELIVERY_ID'].astype(str)
    df_master_file = df_master_file.merge(df_gts[['SAP_DELIVERY_ID', 'INVOICE_NUMBER']].drop_duplicates(), left_on='OUTBOUND DELIVERY', right_on='SAP_DELIVERY_ID', how='left')
    df_master_file = df_master_file.merge(df_ww_cusres[['House Airway Bill', 'IC Invoice Number', 'Entry Date']].drop_duplicates(), left_on='INVOICE_NUMBER', right_on='IC Invoice Number', how='left')
    df_master_file.fillna('Blank', inplace=True)
    df_master_file[['HAWB', 'House Airway Bill']] = df_master_file[['HAWB', 'House Airway Bill']].astype(str)
    df_master_file['House Airway Bill'] = df_master_file['House Airway Bill'].apply(lambda x: x.lstrip('0'))
    if df_master_file.shape[0] != 0:
        df_master_file['COMPARE MASTER FILE'] = df_master_file.apply(lambda x: compare_master_file(x['INVOICE_NUMBER'], x['IC Invoice Number']), axis=1)
    return df_master_file


def separate_data_with_same_part_number(df_ww_cusres, df_gts):
    
    # How many part number there are for every invoice, create field with this information
    df_gts_gr = df_gts['key'].value_counts().reset_index()
    df_gts_gr.rename(columns={'index': 'key', 'key': 'pn_count'}, inplace=True)
    df_ww_cusres_gr = df_ww_cusres['key'].value_counts().reset_index()
    df_ww_cusres_gr.rename(columns={'index': 'key', 'key': 'pn_count'}, inplace=True)
    df_gts = df_gts.merge(df_gts_gr, on='key', how='inner')
    df_ww_cusres = df_ww_cusres.merge(df_ww_cusres_gr, on='key', how='inner')

    # take part number that it's more than 1 time in a invoice
    df_gts_same_pn = df_gts[df_gts['pn_count'] > 1]
    df_ww_cusres_same_pn = df_ww_cusres[df_ww_cusres['pn_count'] > 1]
    
    # take part number that it's only 1 time in a invoice
    df_gts = df_gts[df_gts['pn_count'] == 1]
    df_ww_cusres = df_ww_cusres[df_ww_cusres['pn_count'] == 1]

    # remove part number that it's one time in a system and several times in other one and add them in the dataframes with same part number in a invoice 
    gts_same_pn = list(df_gts_same_pn['key'].drop_duplicates())
    ww_cusres_same_pn = list(df_ww_cusres_same_pn['key'].drop_duplicates())

    df_gts_same_pn2 = df_gts[df_gts['key'].isin(ww_cusres_same_pn)]
    df_ww_cusres_same_pn2 = df_ww_cusres[df_ww_cusres['key'].isin(ww_cusres_same_pn)]

    df_gts = df_gts[~df_gts['key'].isin(ww_cusres_same_pn)]
    df_ww_cusres = df_ww_cusres[~df_ww_cusres['key'].isin(gts_same_pn)]

    df_gts_same_pn = pd.concat([df_gts_same_pn, df_gts_same_pn2])
    df_ww_cusres_same_pn = pd.concat([df_ww_cusres_same_pn, df_ww_cusres_same_pn2])

    # sort dataframes
    df_gts_same_pn.sort_values(by=['INVOICE_NUMBER','PRODUCT_NO', 'COO', 'LINE_QTY'], inplace=True)
    df_ww_cusres_same_pn.sort_values(by=['IC Invoice Number','HP Product Number', "Product's Country of Origin (CIC) in ISO Code", 'Qty in Doc. UoM'], inplace=True)

    list_ww_cusres_same_pn = list(df_ww_cusres_same_pn['key'].drop_duplicates())
    df_f1 = pd.DataFrame()

    # concat horizontally for each invoice 
    for k in list_ww_cusres_same_pn:
        df1 = df_ww_cusres_same_pn[df_ww_cusres_same_pn['key'] == k]
        df2 = df_gts_same_pn[df_gts_same_pn['key'] == k]
        df = pd.concat([df1.reset_index(drop=True), df2.reset_index(drop=True)], axis = 1)
        df_f1 = pd.concat([df_f1, df])
    df_f1 = df_f1.iloc[:, :-2]
    df_f2 = df_ww_cusres_same_pn.merge(df_gts_same_pn, on='key', how='right')
    df_f2.drop(columns=['pn_count_y'], inplace=True)
    df_f2.rename(columns={'pn_count_x': 'pn_count'}, inplace=True)
    df_f2= df_f2[df_f2['IC Invoice Number'].isnull()]

    df_final_same_part_number = pd.concat([df_f1, df_f2])

    
    
    df_f1 = df_ww_cusres.merge(df_gts, on= 'key', how='left')
    df_f2 = df_ww_cusres.merge(df_gts, on= 'key', how='right')
    df_f2 = df_f2[df_f2['HP Product Number'].isnull()]
    df_final_different_part_number = pd.concat([df_f1, df_f2])
    
    df_final = pd.concat([df_final_different_part_number, df_final_same_part_number])
        
    df_final = df_final[['SHIP_TO_COUNTRY', 'Entry Date', 'INVOICE_NUMBER', 'IC Invoice Number', 'PRODUCT_NO' ,'HP Product Number', 'End Usr Name Add 1', 'COO', "Product's Country of Origin (CIC) in ISO Code", 'TERMS_OF_DELIVERY', 'Incoterms', 'CURRENCY', 'Entry Value Currency', 'HTS', 'HTS assigned', 'HTS Class', 'LINE_AMOUNT', 'Total statistical vl', 'TOTAL_INV_AMT', 'Invoice total value', 'LINE_QTY', 'Qty in Doc. UoM', 'Duty Charge Amount','Import VAT Tax amount', 'Import VAT Tax Rate', 'Item Total Amount', 'UNIT_PRICE', 'Unit Price', 'USD converted Import VAT Tax Amount', 'Duty Rate', 'Harbor Maintence Fee (Amount)', 'Entry Number', 'Exchange Rate (KF)']]

    return df_final


def run():
    df_gts_qty = pd.read_excel('./input/GTS_qty.xlsx')
    df_gts_dq = pd.read_excel('./input/GTS_dq.xlsx')
    df_ww_cusres_qty = pd.read_excel('./input/cusres_qty.xlsx')
    df_ww_cusres_dq = pd.read_excel('./input/cusres_dq.xlsx')
    df_duty = get_db_additional('AR')

    df_ww_cusres_qty = get_cusres('AR', df_ww_cusres_qty)
    df_ww_cusres_dq = get_cusres('AR', df_ww_cusres_dq)
    df_gts_qty = get_gts('AR', df_gts_qty)
    df_gts_dq = get_gts('AR', df_gts_dq)

    df_master = get_cusres_qty(df_ww_cusres_qty, df_gts_qty, 'AR')
    df_final = separate_data_with_same_part_number(df_ww_cusres_dq, df_gts_dq)
    df_final = compare(df_final, df_duty)

    df_final['SHIP_TO_COUNTRY'] = 'AR'
    df_g = df_final.groupby(['ENTRY_NUMBER'])['END_USER'].nunique().reset_index()
    df_g.rename(columns={'END_USER': 'COUNT END USER X ENTRY NUMBER'}, inplace=True)
    df_final = df_final.merge(df_g, on='ENTRY_NUMBER', how='inner')
    df_final = df_final[(df_final['INVOICE_NUMBER_CUSRES'].str.startswith('9')) | (df_final['INVOICE_NUMBER_CUSRES'].str.startswith('CR'))]
    
    with pd.ExcelWriter(f'../../PBI/Kill Random Review/Import Declaration Reconciliation Report Argentina.xlsx', index=False) as writer:
        df_final.to_excel(writer, index=False, sheet_name='CUSRES DQ')
        df_master.to_excel(writer, index=False, sheet_name='CUSRES QTY')
    
    columns = list(filter(lambda x: x.startswith('COMPARE'), list(df_final.columns)))


    for c in columns:
        df_final = df_final[df_final[c] != 'Blanks in GTS']
        df_final = df_final[df_final[c] != 'Zero in GTS']
        df_final = df_final[df_final[c] != 'Error GTS']

    with pd.ExcelWriter(f'output/Import Declaration Reconciliation Report Argentina.xlsx', index=False) as writer:
        df_final.to_excel(writer, index=False, sheet_name='CUSRES DQ')
        df_master.to_excel(writer, index=False, sheet_name='CUSRES QTY')

    df_final = df_final[columns]

    data = []
    base = {'Equal': 0, 'Different': 0, 'Blanks in Cusres':  0, 
            'Blanks in GTS': 0}
    for i in columns:
        d = {'Criterio': i} | dict(df_final[i].value_counts())
        c = base | d
        data.append(c)

    data_email = pd.DataFrame(data)
    
    list_criterio = data_email['Criterio'].tolist()
    list_qty_equal = data_email['Equal'].tolist()
    list_qty_different = data_email['Different'].tolist()
    #list_qty_blanks_gts = data_email['Blanks in GTS'].tolist()
    list_qty_blanks_cusres = data_email['Blanks in Cusres'].tolist()
    first_day, today = establish_range()
    range_date = first_day + ' al ' + today
    total_rows = str(df_final.shape[0])

    file_path_report = pathlib.Path('output/Import Declaration Reconciliation Report Argentina.xlsx')
    file_absolute_report = str(file_path_report.absolute())
    file_path_img = pathlib.Path('input/LAR Import Declaration Reconciliation - (AR).png')
    file_absolute_img = str(file_path_img.absolute())


    #send_email(list_criterio, list_qty_equal, list_qty_different, list_qty_blanks_cusres, week, range_date, total_rows, file_absolute_report, file_absolute_img)

if __name__ == '__main__':
    run()