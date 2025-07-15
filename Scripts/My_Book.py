# Author: Rodrigo Bedolla Fuerte
# Department: Order management
# Bedolla Fuerte, R. (Dec 06, 2021). My_Book (version No 1.6 | last update by Ana Barraza (June 18, 2025)). Cd. Juarez: Foxconn eCMMS S.A. DE C.V. V1.1

import win32com.client as win32
import pandas as pd
import os
import datetime
import calendar
from pathlib import Path
import numpy as np
import re
import subprocess
import json

#Shared Folder path
def share_path():

    #def_path = '\\\\10.19.16.56\\Order Management\\OM Projects'
    def_path = '\\\\10.19.17.32\\CygnusFiles\\OM_RPA\\OM Projects'
    #def_path = path_home()+'\\Desktop\\Raw_Files'
    
    return def_path

#Convert xls file to xlsx file
def convert_xlsx(file):
    try:

        fname = path()+'\\Files\\'+file
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(fname)

        excel.DisplayAlerts = False

        wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                    
    except Exception as e:
        #FileFormat = 56 is for .xls extension
        excel.Application.Quit()
        fname = path()+'\\Files\\'+file
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(fname)

        excel.DisplayAlerts = False

        wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close() 
        
#Get current time stamp 
def get_time():

    time_stamp = datetime.now()

    return time_stamp

#Convert txt File 1 dimention to a array list
def txt_array(z_file):
    
    with open(share_path()+'\\Files_Format\\'+z_file) as f:
        content = f.readlines()
        # you may also want to remove whitespace characters like `\n` at the end of each line
    content = [x.strip() for x in content]
    return content


#
def local_txt_array(z_file):
    
    with open(path()+'\\Files\\'+z_file) as f:
        content = f.readlines()
        # you may also want to remove whitespace characters like `\n` at the end of each line
    content = [x.strip() for x in content]
    return content

def txt_array_2d(z_file):

    with open(share_path()+'\\Files_Format\\'+z_file) as textFile:

        lines = [line.split() for line in textFile]

    return lines

#Create new txt File or append data to existing file
def create_txt(value,file_name,write_tipe):

    if write_tipe == 'append':

        with open(share_path()+'\\Files_Format\\'+file_name, 'a') as f:
            f.write('\n'+value)
    
    else:

        with open(share_path()+'\\Files_Format\\'+file_name, 'w') as f:
            f.write(value)
#ZSD5 Format
def zsd5_format(file_route):

    df_zsd5 = pd.read_excel(file_route)
    df_zsd5.rename(columns={"WO": "WORK ORDER"}, inplace=True)

    #identify Job finished! and drop empty rows
    value = df_zsd5["TYPE"].iloc[len(df_zsd5.index)-1]
    if value == 'Job finished!':
        df_zsd5=df_zsd5.drop(df_zsd5.index[len(df_zsd5.index)-2:len(df_zsd5.index)])

    return df_zsd5

#ZSD6 Format
def zsd6_format(file_route):

    zsd6_columns = txt_array('zsd6_columns.txt')
    df1 = pd.read_excel(file_route)

    df1=df1.drop(df1.index[[0]])
    df1.columns = df1.iloc[0]
    df1 = df1.drop(df1.index[[0]])
    df1 = df1.loc[:, df1.columns.notnull()]
    df1.columns = df1.columns[:0].tolist() + zsd6_columns

    #identify Job finished! and drop empty rows
    value = df1["TYPE"].iloc[len(df1.index)-1]
    if value == 'Job finished!':
        df1=df1.drop(df1.index[len(df1.index)-2:len(df1.index)])

    return df1

#ZSD6a Format
def zsd6a_format(file_route):

    zsd6a_columns = txt_array('zsd6a_columns.txt')

    df2 = pd.read_excel(file_route)

    df2=df2.drop(df2.index[[0]])
    df2.columns = df2.iloc[0]
    df2=df2.drop(df2.index[[0]])
    df2 = df2.loc[:, df2.columns.notnull()]
    df2.columns = df2.columns[:0].tolist() + zsd6a_columns

    value = df2["SO DATE"].iloc[len(df2.index)-1]

    if value == 'Job finished!':
        df2=df2.drop(df2.index[len(df2.index)-2:len(df2.index)])

    return df2



#Remove multiple column list from specific Dataframe
def drop_list_of_columns(column_list,df):

    for col in column_list: 
        for indice in df.columns:
            if col in indice and (len(col)==len(indice)):   
                del df[indice]
    return df
    
    #try this
    #return dataset.drop(cols, axis=1)

#Get specific column removing duplicates and export to txt file (sap input)
def sap_input(df_file,column):

    column_names = [column]
    df_sales_orders = pd.DataFrame(columns = column_names)

    if column == 'SO':
        df_sales_orders = df_file[column].drop_duplicates().astype(int)
    else:
        df_sales_orders = df_file[column].drop_duplicates().astype(str)

    df_sales_orders.to_csv(path()+'\\Files\\' + str(column) + '.txt', header=None, index=None) #Guardar archivo de txt
    return df_sales_orders

def previous_labor_day():

    current_day = datetime.date.today()
    holidays=txt_array('holidays.txt')
    final_date = current_day + datetime.timedelta(days=-1)

    if calendar.day_name[final_date.weekday()] == 'Saturday':
        final_date = final_date + datetime.timedelta(days=-1)
    elif calendar.day_name[final_date.weekday()] == 'Sunday':
        final_date = final_date + datetime.timedelta(days=-2) 
    while final_date.strftime('%m/%d/%Y') in holidays:
        final_date = final_date + datetime.timedelta(days=-1)
        if calendar.day_name[final_date.weekday()] == 'Saturday':
            final_date = final_date + datetime.timedelta(days=-1)
        elif calendar.day_name[final_date.weekday()] == 'Sunday':
            final_date = final_date + datetime.timedelta(days=-2)

    return pd.Timestamp(final_date)

def format_date(format):

    current_day = datetime.datetime.now()
    time_stamp = datetime.date.strftime(current_day, '%m-%d-%Y %H:%M:%S')
    default_date = datetime.date.today()
    formatted_date = datetime.date.strftime(current_day, "%m/%d/%Y")
    filedate = formatted_date.replace('/','')
    month_day = datetime.date.strftime(current_day, '%m/%d')
    month_name_day =datetime.date.strftime(current_day,'%b %d')
    previous_date = previous_labor_day().strftime("%m/%d/%Y")
    previous_formatted_day = previous_date.replace('/','')
    #datetime.datetime.now()

    #Dates Format
        # 1.- format date mm-dd-yyyy hh:mm:ss:ffffff
        # 2.- format date mm/dd/yyyy
        # 3.- format date mmddyyyy
        # 4.- format date mm/dd
        # 5.- format date mm-dd-yyyy
        
        # 7. -format previous date mm/dd/yyyy
        # 8. -format previous date mmddyyyy

    if format == 1:                
        return time_stamp
    elif format == 2:
        return formatted_date
    elif format == 3:
        return filedate
    elif format == 4:
        return month_day
    elif format == 5:
        return pd.Timestamp(default_date)
    elif format == 6:
        return month_name_day
    elif format == 7:
        return previous_date
    elif format == 8:
            return previous_formatted_day
        
        
        
def base_sku_column(df):

    for i in range(0,len(df.index)):

        material = df['MATERIAL'][i]

        if 'FG' in material:
            df.at[i,'BASE SKU'] = material[:material.find('FG')]
        elif '#' in material:
            df.at[i,'BASE SKU'] = material[:material.find('#')]
        else:
            df.at[i,'BASE SKU'] = material

    return df

def project(file):

    #Parameters 
    # file = Destination file

    wo_types = pd.read_excel(share_path()+'\\Master Template\\Material Master - WO Types.xlsx', sheet_name='WO TYPES')

    file = file.merge(wo_types[['WO TYPE','COMPLEXITY']], on='WO TYPE', how='left')

    for i in range(0,len(file.index)):

        base_sku = file['MATERIAL'][i]
        complexity = file['COMPLEXITY'][i]

        if complexity == 'HPSD' and base_sku.find('FG') > 0:
            file.loc[i,'PROJECT'] = 'HPSD CTO'
        elif complexity == 'HPSD':
            file.loc[i,'PROJECT'] = 'HPSD BTO'
        elif complexity == 'VALIDAR' and base_sku[6:7] == 'R':
            file.loc[i,'PROJECT'] = 'REMAN TRADE'
        elif complexity == 'VALIDAR':
            file.loc[i,'PROJECT'] = 'DIRTY ORDER'
        else:
            file.loc[i,'PROJECT'] = complexity

    file['PROJECT'] = np.where(file['MATERIAL'].astype(str).str.contains('BD505A'),'PPS',file['PROJECT'])

    drop_list_of_columns(['COMPLEXITY'],file)
    
    return(file)

def family(file):

    sku_summary = pd.read_excel(share_path()+'\\Master Template\\Material Master - WO Types.xlsx', sheet_name='Material Master')
    file = file.merge(sku_summary[['BASE SKU','FAMILY']], on='BASE SKU', how='left')
    file['FAMILY'] = file['FAMILY'].fillna('FAMILY NOT FOUND')

    return file

def clean_source():
    log_file_route = path()+'\\Files'
    for i in os.listdir(log_file_route):
        os.remove(path()+'\\Files\\' + i)

def primary_key(df):
    
    #CREATE PRIMARY KEY
    #Find WO's and concatenate PO + ITEM + BASE SKU ON MISSING WO's

    df['PRIMARY KEY'] = df.apply(
        lambda row: str(row['PO']) + str(row['ITEM']) + str(row['BASE SKU']) 
        if pd.isna(row['WORK ORDER']) else row['WORK ORDER'], axis=1
    )
    
    df['PRIMARY KEY'] = df['PRIMARY KEY'].astype(str).str.replace(r"\.0$", "", regex=True)
    df = df[['PRIMARY KEY'] + [col for col in df.columns if col != 'PRIMARY KEY']]

    return df

def dates_operations(operation,days):

    current_day = datetime.datetime.now()

    if operation == 'sum':

        current_day = datetime.date.today()
        end_date = current_day + datetime.timedelta(days=+days)

        return end_date

    if operation == 'less':
        
        current_day = datetime.date.today()
        end_date = current_day + datetime.timedelta(days=-days)

        return end_date

def week_day():

    current_day = datetime.datetime.now()
    week_day = calendar.day_name[current_day.weekday()]

    return week_day

def M2_dailys(df):
    
    #BACKLOG_COLUMN
    #column to create buckets for new orders(curren day), new orders(previous day), new orders (weekends) and Backlog for aged orders

    for i in range(0,len(df.index)):

        a = df['SO DATE'][i].date()

        # 11/29/2021 -> change value 4 to 3 in monday condition
        
        if a == format_date(5):
            df.loc[i,'OPEN BACKLOG']  = datetime.date.strftime(format_date(5),'%b %d')
        elif previous_labor_day() <= a <= dates_operations('less',1):
            df.loc[i,'OPEN BACKLOG'] = datetime.date.strftime(previous_labor_day(),'%b %d')
        else:
            df.loc[i,'OPEN BACKLOG'] = 'BACKLOG'

    return df

def dailys(df):
    
    #BACKLOG_COLUMN
    #column to create buckets for new orders(curren day), new orders(previous day), new orders (weekends) and Backlog for aged orders

    for i in range(0,len(df.index)):
        
        df.loc[:, 'SO DATE'] = pd.to_datetime(df['SO DATE'], errors='coerce')
        a = df['SO DATE'][i]

        # 11/29/2021 -> change value 4 to 3 in monday condition


        if a == format_date(5):
            df.loc[i,'BACKLOG']  = datetime.date.strftime(format_date(5),'%b %d')
        elif previous_labor_day() <= a <= pd.Timestamp(dates_operations('less',1)):
            df.loc[i,'BACKLOG'] = datetime.date.strftime(previous_labor_day(),'%b %d')
        else:
            df.loc[i,'BACKLOG'] = 'BACKLOG'

    
    #COMPLEXITY CATEGORY
    #PPS, SERVERS Y RACKS
    for i in range(0,len(df.index)):

        bts_reman = df['SHIP TO'][i]
        dirty_orders = df['PROJECT'][i]
        complexity = df['PROJECT'][i]
        reman_trade = df['FAMILY'][i]
        racks = df['FAMILY'][i]

        if 'DO NOT SHIP' in bts_reman:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'BTS REMAN'
        elif 'DIRTY ORDER' in dirty_orders:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'DIRTY ORDER'
        elif 'RACK' in racks:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'RACKS'
        elif 'PPS' in complexity:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'PPS'
        elif 'CTO' in complexity or 'BTO' in complexity:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'SERVERS'
        elif 'OPTION' in reman_trade or 'BUY' in reman_trade:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'PPS'
        elif 'REMAN TRADE' in complexity:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'SERVERS'
    
    #COMPLEXITY
    #PPS, BTO, SIMPLE CTO, COMPLEX CTO, BLADES Y HPSD's

    for i in range(0,len(df.index)):

        bts_reman = df['PROJECT'][i]
        dirty_orders = df['PROJECT'][i]
        complexity = df['PROJECT'][i]
        fmx_family = df['FAMILY'][i]
        racks = df['FAMILY'][i]
        
        if 'DO NOT SHIP' in bts_reman:
            df.loc[i,'COMPLEXITY CATEGORY']  = 'BTS REMAN'
        elif 'DIRTY ORDER' in dirty_orders:
            df.loc[i,'COMPLEXITY']  = 'DIRTY ORDER'
        elif 'RACK' in racks:
            df.loc[i,'COMPLEXITY']  = 'RACKS'
        elif 'BL' == fmx_family[:2] and (('PPS' in complexity) == False):
            df.loc[i,'COMPLEXITY']  = 'BLADES'
        elif 'PPS' in complexity:
            df.loc[i,'COMPLEXITY']  = 'PPS'
        elif 'BTO' in complexity:
            df.loc[i,'COMPLEXITY']  = 'BTO'
        elif 'HPSD' in complexity:
            df.loc[i,'COMPLEXITY']  = 'HPSD'
        elif 'sCTO' in complexity:
            df.loc[i,'COMPLEXITY']  = 'SIMPLE CTO'
        elif 'cCTO' in complexity:
            df.loc[i,'COMPLEXITY']  = 'COMPLEX CTO'
        elif 'OPTION' in fmx_family or 'BUY' in fmx_family:
            df.loc[i,'COMPLEXITY']  = 'PPS'
        elif 'REMAN TRADE' in complexity:
            df.loc[i,'COMPLEXITY']  = 'BTO'

    return df

def path(): #Current project path delimited by '\'

    path = str(Path(__file__).parent.parent)
    return path

def path_home(): #Root user path delimited by '\'

    home = str(Path.home())
    return home

def windows_path(): #Current project path delimited by '\\'

    path = str(Path(__file__).parent.parent).replace('\\','\\\\')
    return path

def windows_path_home(): #Root user path delimited by '\\'

    home = str(Path.home()).replace('\\','\\\\')
    return home

def root_path(jumps_back):

    root = str(Path(__file__).parents[jumps_back])

    return root

def delete_local_files(): #Delete file from 'Files' folder for specific project

    log_file_route = path()+'\\Files'
    for i in os.listdir(log_file_route): 
    #    if (i != 'SHIP_STATUS.xlsx') & (i != 'zsd6.xlsx') & (i != 'zsd6a.xlsx'):
        print("Removed: "+i)
        os.remove(path()+'\\Files\\' + i)

def sap_decoding_zsd6_files(flag,file_route):

    if flag == 'zsd6':

        df =  pd.read_csv(file_route, skiprows=[0,1], sep='\\t', thousands=',', engine='python', encoding='ISO-8859-1')

        df = df.dropna(axis=1, how='all')

        df.rename(columns={'TYPE.1': 'WO TYPE', 'COUN' : 'COUNTRY'}, inplace=True)
        df.columns = df.columns.str.lstrip()
        df = df[txt_array('zsd6_columns.txt')]

        if df["TYPE"].iloc[len(df.index)-1] == 'Job finished!':

            df=df.drop(df.index[len(df.index)-1:len(df.index)])

        return df

    elif flag == 'zsd6a':

        df =  pd.read_csv(file_route, skiprows=[0,1], sep='\\t', thousands=',', engine='python', encoding='ISO-8859-1')

        df = df.dropna(axis=1, how='all')
        df.columns = df.columns.str.lstrip()
        df.rename(columns={'OPEN': 'OPEN QTY', 'CO' : 'COUNTRY','TYPE' : 'WO TYPE','ACK' : 'RE-ACK'}, inplace=True)

        try:
            df['OPEN QTY'] = np.where(df['OPEN QTY']!=df['WO QTY'],df['WO QTY'],df['OPEN QTY'])
        except Exception as e:
            print(e)

        df = df[txt_array('zsd6a_columns.txt')]

        if df["SO DATE"].iloc[len(df.index)-1] == 'Job finished!':

            df=df.drop(df.index[len(df.index)-1:len(df.index)])

        return df
    
    else:

        return False

#03162022 Adding SO ID to do merge with previous Master
def primary_key_by_so(df):
    
    #CREATE SO ID
    #Find WO's and concatenate SO + ITEM + BASE SKU ON MISSING WO's and SO QTY for reman orders

    for i in range(0,len(df.index)):

        try:

            deletion_flag = df['DELETION FLAG'][i]

        except Exception as e:
            
            deletion_flag = True
        
        work_order_key = df['WORK ORDER'][i]
        second_key = str(df['SO'][i]) + str(df['ITEM'][i]) + str(df['BASE SKU'][i])
        reman_id = str(df['SO'][i]) + str(df['ITEM'][i]) + str(df['BASE SKU'][i])+str(df['SO QTY'][i])

        if deletion_flag == True:

            if df['BASE SKU'][i][6:7] == 'R' and (work_order_key != work_order_key or len(str(work_order_key)) < 8):
                df.loc[i,'SO ID'] = reman_id
            elif work_order_key != work_order_key:
                df.loc[i,'SO ID'] = second_key
            else:
                df.loc[i,'SO ID'] = work_order_key

        else:

            if df['BASE SKU'][i][6:7] == 'R' and deletion_flag == 'X':
                df.loc[i,'SO ID'] = reman_id
            elif work_order_key != work_order_key:
                df.loc[i,'SO ID'] = second_key
            else:
                df.loc[i,'SO ID'] = work_order_key
    
    df = df[ ['SO ID'] + [ col for col in df.columns if col != 'SO ID' ] ]
    df['SO ID'] = df['SO ID'].astype(str).str.replace("\\.0$", "", regex=True)

    return df

def owner_column(df):

    wo_types = pd.read_excel(share_path()+'\\Master Template\\Material Master - WO Types.xlsx', sheet_name = 'WO TYPES')
    wo_types = wo_types.rename(columns = {'PROJECT':'OWNER'})
    df = df.merge(wo_types[['WO TYPE','OWNER']] , on = 'WO TYPE',how = 'left').drop_duplicates().reset_index(drop=True)

    for i in range(0,len(df.index)):

        base_sku = df['BASE SKU'][i]
        project_column = df['PROJECT'][i]

        if base_sku[6:7].upper() == 'R':
            df.loc[i,'OWNER'] = 'CESAR RAMIREZ'
        elif 'HPSD' in project_column:
            df.loc[i,'OWNER'] = 'DANIELA OLIVAS'
        elif project_column == 'DIRTY ORDER':
            df.loc[i,'OWNER'] = 'OMAR CAMILO'
    
    return df

def count_item_qty(df,column_name,column_result):

    item_qty = df[column_name].value_counts()
    item_qty = item_qty.to_frame().reset_index()
    item_qty.columns = [column_name,column_result]
    df = df.merge(item_qty,on=column_name,how='left').drop_duplicates().reset_index(drop=True)

    return df

def cookie_cygnus():

    #Login to CyGNUS and save cooke with credentials
    subprocess.call('sh bash_scripts/Login.sh')

def extract_ssr_number_zjm2():

    with open(path()+'\\json_files\\Cygnus_Files.json', "r", encoding='utf-8') as file:
        file_content = file.read()

    pattern = r'"item3":"(\\d+)"'
    ssr_number = ''
    match = re.search(pattern, file_content)

    if match:

        ssr_number = match.group(1)
    
    return str(ssr_number)

def cyg_logout():
    
    subprocess.call('sh bash_scripts/Logout.sh')

def current_date():

    dt = datetime.datetime.strptime(str(dates_operations('less',0)), '%Y-%m-%d')

    return dt

def complexities(df):

    df['COMPLEXITY'] = np.where(df['PROJECT'].str.contains('PPS'),'PPS',
                       np.where(df['FAMILY'].str.contains('RACK'),'RACKS',
                       np.where(df['BASE SKU'].str.contains('BD505A'),'PPS',
                       np.where(df['PROJECT'].str.contains('cCTO'),'COMPLEX CTO',
                       np.where(df['PROJECT'].str.contains('BTO'),'BTO',
                       np.where(df['PROJECT'].str.contains('CTO'),'CTO',df['PROJECT']))))))
    
    return df

def analysis_level_column_id(df):

    df['PO + ITEM'] = df['PO'].astype(str) + df['ITEM'].astype(str)
    nested_items = df['PO + ITEM'][(df['SHIP TYPE']=='SP')].to_list()
    nested_items = set(x for x in nested_items if nested_items.count(x) > 1)
    df['ANALYSIS_LEVEL'] = np.where(df['SHIP TYPE']=='SC',df['PO'],np.where(df['PO + ITEM'].isin(nested_items),df['PO + ITEM'],df['WORK ORDER']))

    df = df.drop(columns={'PO + ITEM'})

    return df

def worst_complexity(df):

    worst_complexity = df.copy()
    worst_complexity = worst_complexity[['SHIP TYPE','FAMILY','WORK ORDER','ANALYSIS_LEVEL','PROJECT']].drop_duplicates(subset='WORK ORDER')
    worst_complexity['ORDER_COMPLEXITY'] = np.where(worst_complexity['FAMILY'].str.contains('RACK'),'RACKS',
                                            np.where(worst_complexity['PROJECT'].str.contains('DISMANTLED'),'DISMANTLED',
                                            np.where(worst_complexity['PROJECT'].str.contains('PPS'),'PPS','SERVERS')))

    worst_complexity['CATEGORY_LEVEL'] = np.where(worst_complexity['ORDER_COMPLEXITY'].str.contains('RACK'),1,
                                            np.where(worst_complexity['ORDER_COMPLEXITY'].str.contains('DISMANTLED'),4,
                                            np.where(worst_complexity['ORDER_COMPLEXITY'].str.contains('PPS'),3,2)))

    worst_complexity = worst_complexity.sort_values(by='CATEGORY_LEVEL', ascending=True).drop_duplicates(subset='ANALYSIS_LEVEL', keep='first').reset_index(drop=True)

    df = df.merge(worst_complexity[['ANALYSIS_LEVEL','ORDER_COMPLEXITY']], on='ANALYSIS_LEVEL', how='left')

    return df

def remove_decimals(df,column_name):
    print('Removing decimals from column: '+column_name)
    df[column_name] = df[column_name].astype(str).str.replace('\\.0$', '', regex=True)

    return df

def get_time():

    time_stamp = datetime.datetime.strptime(format_date(1),'%m-%d-%Y %H:%M:%S')

    return time_stamp

def sql_parameters():
    """
    Reads the database connection string from a text file and returns it.

    Returns:
        str: The database connection string.
    """
    with open(share_path()+'\\Files_Format\\db_connection.txt', 'r') as file:
        conn_str = ''.join(line.strip() for line in file)

    return conn_str

def convert_date(date_str):
    # Convert the date string to a datetime object
    date_obj = datetime.datetime.strptime(date_str, '%m/%d/%Y')
    # Format the datetime object as YYYY-MM-DD
    formatted_date = date_obj.strftime('%Y-%m-%d')
    return formatted_date



def sq01_sap_output_format(data_source):

    df =  pd.read_csv(data_source, skiprows=[0,1,2,3,5], sep='\\t', thousands=',', engine='python', encoding='ISO-8859-1').fillna('-')
    
    df = df[~(df['Work Order']=='-')].reset_index(drop=True)
    df['Work Order'] = df['Work Order'].astype(np.int64)

    return df

def json_load(data_source):

    data = open(data_source,'r') #read json file downloaded
    json_string = json.load(data)
    return json_string

def po_viewer_request(po_list,Trantype):
    
    subprocess.call('sh API_connection/Request.sh '+str(po_list)+' '+str(Trantype))
    
    with open(path()+'\\json_files\\Cygnus_API.json', 'r', encoding='utf-8') as f:
            data = json.load(f)

    if Trantype != 'POViewerDetail':

        try:
            df = pd.DataFrame(json.loads(pd.DataFrame([data])['JSONResponse'][0]))
            df['DATE'] = pd.to_datetime(df['DATE'], format='ISO8601')
            df['PO DATE'] = pd.to_datetime(df['PO DATE'], format='ISO8601')
            #df = df.sort_values(by='DATE', ascending=False).drop_duplicates(subset='PO', keep='first')

            return df

        except:

            json_file = str(json_load(path()+'\\json_files\\Cygnus_API.json'))

            return json_file

    else:

        try:

            json_data = pd.DataFrame(json.loads(pd.DataFrame([data])['JSONResponse'][0]))
            pov_history = pd.DataFrame(columns=txt_array('POV_Detail.txt'))
            grouped = json_data.groupby('PO')
            b=1
            for po,group in grouped:

                group = group.reset_index(drop=True)
                po_viewer_detail = pd.DataFrame(pd.DataFrame(group)['PoLog'][0])
                po_viewer_detail['PO'] = po
                #po_viewer_detail = po_viewer_detail[po_viewer_detail['SAP_ITEM'].astype(int)>0]
                pov_history = pd.concat([pov_history,po_viewer_detail])
                #print(po_viewer_detail)

            pov_history.sort_values(by='LAST EDIT DATE', ascending=False).drop_duplicates(subset=['SAP_ITEM', 'PO'],keep='first').reset_index(drop=True)
            #pov_history.to_excel(path()+'\\Files\\PO_Detail.xlsx', index=False)

            return pov_history

        except:

            json_file = str(json_load(path()+'\\json_files\\Cygnus_API.json'))

            return json_file

def api_po_viewer(df,api_flag):

    if api_flag == 'POViewerHeader':
        series_input = 500
    else:
        series_input = 500

    first = 0
    last = series_input
    sub_value = series_input
    batch = 0
    cyg_input = [1]

    df1 = pd.DataFrame({'PO' : []})
    
    df = df[['PO']]
    df = df.drop_duplicates().reset_index(drop=True)

    df['PO'] = df['PO'].astype(str)
    df['PO'] = df['PO'].str.replace("\\.0$", "",regex=True)

    while ((sub_value % series_input) == 0) & (len(cyg_input)>0):

        df1 = df['PO'].iloc[first:last]
        sub_value = len(df1.axes[0])
        first = last
        last = last + sub_value
        df1=[str(int) for int in df1]
        cyg_input = ",".join(df1)

        if batch==0:
            final_df=po_viewer_request(cyg_input,api_flag)
        else:
            df_cyg=po_viewer_request(cyg_input,api_flag)
            
            if 'DATA NOT FOUND' in final_df or 'DATA NOT FOUND' in df_cyg:
                print('DATA NOT FOUND')
            else:
                frames = [final_df, df_cyg]
                final_df = pd.concat(frames)

        #print(final_df)

        batch+=1

    if final_df.empty == False:
        #final_df['WO'] = final_df['WO'].astype(np.int64)
        return final_df.reset_index(drop=True)
    else:
        return False
        