# Author: Ana Barraza Reyes
# Department: Order management
# Client: HPE Planning
# Ana Barraza, R. (Junio 16, 2025). Read_files (version No 2.2 | last update by Ana Barraza (Junio 18, 2025)). Cd. Juarez: Foxconn eCMMS S.A. DE C.V. V1.1


#IMPORTING LIBRARIES---------------------------------------------------------------------------------------
from My_Book import *
from SAP import *
from emailSender import *

# FUNCTION TO READ FILES AND SEND EMAIL WITH THE ATPS -----------------------------------------------------
def read_files_ATPS():
    
    #Deleting files of path to create new ones and calling function to get info from SAP(ZATPRESULT)------
    delete_local_files()
    saplogin(3)
    

    #CALLING FILES ----------------------------------------------------------------------------------------
    # EXCEL File "Signal 855 Confirmation" We only need column 'RECOMMIT QTY'       SRC = Source (N)
    df_signal_855 = pd.read_excel(share_path()+'\\OM_RPAs_Files\\Backup\\Open\\Open '+str(format_date(3))+'.xlsx', usecols={'WORK ORDER','RECOMMIT QTY'})

    # EXCEL File "BACKLOG SEQUENCE ANALYSIS" We only need columns 'CATEGORY','BACKLOG SEQUENCE', 'DNT'
    #try:
        #df_blSEQ = pd.read_excel(share_path() + '\\Backlog_Sequence\\Backlog_Sequence_' + str(format_date(3)) + '.xlsx',usecols={'WORK ORDER','CATEGORY','BACKLOG SEQUENCE','DNT'})
    df_blSEQ = pd.read_excel(share_path() + '\\Backlog_Sequence\\Backlog_Sequence_' + str(format_date(8)) + '.xlsx',usecols={'WORK ORDER','CATEGORY','BACKLOG SEQUENCE','DNT'})

    # SAP Spreadsheet download from "ZATPRESULT" IN 902 "HPE" we only need columns 'Order', 'Material', 'Material description', 'Purch.doc.','Component','Reqmts qty'
    df_zatpresult =  pd.read_csv(path() +'\\Files\\zatpresult_woDetails.xls', skiprows=[0,1], sep='\\t', thousands=',' , engine='python', encoding='ISO-8859-1', usecols={'Order','Material','Material description','Purch.doc.','Component','Reqmts qty'})

    # EXCEL File "MPS 4AM" We only need columns 'PC', 'SHIP TYPE', 'PO', 'COMPLEXITY', 'QTY', 'RDD'. 'SINGLE_GATED_FLAG', 'MATERIAL P','SKU DESCRIPTION','DELIVERY D','ACTUAL SCHEDULED DATE'. The Analysis Level is the column that will be used to filter the MPS 4AM file to remove the duplicates
    df_mps4AM = pd.read_excel(share_path() + '\\Planning\\MPS_Base_Record\\MPS_'+str(format_date(3))+'.xlsx', usecols={'WORK ORDER','PC','SHIP TYPE','PO','COMPLEXITY','QTY','RDD','SINGLE_GATED_FLAG','MATERIAL P','SKU DESCRIPTION','DELIVERY D','ACTUAL SCHEDULED DATE','ANALYSIS_LEVEL'})


    #MAKING COPYs --------------------------------------------------------------------------------------------
    df_signal_855 = df_signal_855.copy()
    df_backlog_sequence = df_blSEQ.copy()
    df_sap_zatpresult = df_zatpresult.copy()
    df_mps_4am = df_mps4AM.copy()


    #RENAME OF COLUMNS ---------------------------------------------------------------------------------------
    df_sap_zatpresult = df_sap_zatpresult.rename(columns={'Order':'WORK ORDER'})  
            #print('zatpresult shape',df_sap_zatpresult.shape) # Para ver el numero de filas y columnas del dataframe


    # convert WORK ORDER to string and remove decimals--------------------------------------------------------
    df_signal_855 = remove_decimals(df_signal_855,'WORK ORDER')
    df_backlog_sequence = remove_decimals(df_backlog_sequence,'WORK ORDER')
    df_sap_zatpresult = remove_decimals(df_sap_zatpresult,'WORK ORDER')
    df_mps_4am = remove_decimals(df_mps_4am,'WORK ORDER')
    df_mps_4am = remove_decimals(df_mps_4am,'ANALYSIS_LEVEL')


    #DROP DUPLICATE COLUMNS-------------------------------------------------------------------------------------
    df_analysis_level = df_mps_4am[['WORK ORDER','ANALYSIS_LEVEL']].drop_duplicates(subset='WORK ORDER').reset_index(drop=True)
    df_signal_855 = df_signal_855.drop_duplicates(subset='WORK ORDER').reset_index(drop=True)
    df_backlog_sequence = df_backlog_sequence.drop_duplicates(subset='WORK ORDER').reset_index(drop=True)

    # Long Pole Analysis: substract critcal short based on longest recovery date ------------------------------
    df_mps_4am = df_mps_4am[df_mps_4am['MATERIAL P']!='-'].reset_index(drop=True)
    df_mps_4am = df_mps_4am.sort_values(by=['ANALYSIS_LEVEL','DELIVERY D','MATERIAL P'], ascending=[True,False,True]).drop_duplicates(subset='ANALYSIS_LEVEL').reset_index(drop=True)
    df_mps_4am = df_mps_4am.drop(columns={'WORK ORDER'})


    #Making merge adding columns and removing blanks and nulls-------------------------------------------------
    df_sap_zatpresult = df_sap_zatpresult.merge(df_analysis_level, on='WORK ORDER', how='left')
    df_sap_zatpresult = df_sap_zatpresult.merge(df_signal_855, on='WORK ORDER', how='left').fillna(0)
    df_sap_zatpresult = df_sap_zatpresult.merge(df_backlog_sequence, on='WORK ORDER', how='left').fillna('-')
    df_sap_zatpresult = df_sap_zatpresult.merge(df_mps_4am, on='ANALYSIS_LEVEL', how='left').fillna('-')


    #Adding new columns with default values--------------------------------------------------------------------
    new_columns = ['PO QTY', 'NEW RECOVERY DATE', 'CTB']
    for col in new_columns:
        df_sap_zatpresult[col] = '-'


    #Reorganizate changing names and printing------------------------------------------------------------------
    df_sap_zatpresult = df_sap_zatpresult [txt_array('HPE_PLANNING_WO_DETAIL_SUMMARY.txt')]

    df_sap_zatpresult = df_sap_zatpresult.rename(columns={
        'BACKLOG SEQUENCE': 'SEQ',
        'DNT': 'NO DONATE',
        'QTY': 'WO QTY',
        'SINGLE_GATED_FLAG': 'SINGLE',
        'WORK ORDER': 'ORDER',
        'Material': 'MATERAL',
        'Material description': 'MATERIAL DESCRIPTION',
        'Purch.doc.': 'PURCH.DOC.',
        'Component': 'COMPONENT',
        'Reqmts qty': 'REQMTS QTY',
        'MATERIAL P': 'LONG POLE PN',
        'SKU DESCRIPTION': 'DESCRIPTION',
        'DELIVERY D': 'ETA',
        'ACTUAL SCHEDULED DATE': 'ACTUAL RECOVERY DATE'
    })

    df_sap_zatpresult.to_excel(path()+'\\Files\\ATPS_'+format_date(3)+'.xlsx',index=False)

    # Making the mail -----------------------------------------------------------------------------------
    recipient_email = "Manuel.SantosTorres@FII-NA.com; Alejandro.Prado@fii-na.com; ana.barraza@fii-na.com;"
    email_copy = "Rodrigo.Bedolla@FII-NA.com; Erik.CarbajalR@FII-NA.com; Bryan.Rodriguez@FII-NA.com;"
    # recipient_email = "ana.barraza@fii-na.com; jonathan.ponce@fii-na.com;"
    # email_copy = ""
    subject = "ATPS "+format_date(2)
    file_attachments = [path()+'\\Files\\ATPS_'+format_date(3)+'.xlsx']


    # Sending the mail -----------------------------------------------------------------------------------
    send_mail_with_excel(recipient_email, email_copy, subject, file_attachments)
