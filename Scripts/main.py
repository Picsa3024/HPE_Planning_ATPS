# Author: Rodrigo Bedolla Fuerte
# Department: Order management
# Bedolla Fuerte, R. (Dec 06, 2021). main (version No 2.1 | last update by Ana Barraza (June 18, 2025)). Cd. Juarez: Foxconn eCMMS S.A. DE C.V. V1.1

import schedule
import time
from read_files import *
from Execution_log import Execution_log
from Email_Alerts import *
from My_Book import week_day,txt_array

flag = 0
error_count = 0
start = 0

def main():

    global start
    global error_count

    start = get_time()
    
    if week_day() not in txt_array('Weekend_Execution.txt'):
        
        # var1 = 5 / int(local_txt_array('possibleError.txt')[0])
        # print(var1)
        read_files_ATPS()
        Execution_log(start,'SUCCESS','','-')
    
    else:

        print('WEEEKEND: '+str(txt_array('Weekend_Execution.txt')))
    
    error_count = 0

    schedule.clear()
    job()  

def job():

    global error_count

    try:

        if error_count == 0:

            schedule.every().day.at("07:28").do(main)      

        elif error_count <= 5:

            schedule.every(5).minutes.do(main)

        else:

            schedule.every(30).minutes.do(main) 

        while True:

            schedule.run_pending()
            time.sleep(1)

    except Exception as error:

        df_table = Execution_log(start,'FAIL',error,'-')
        
        df_table = df_table[['SCRIPT NAME','START', 'FINISH','EXECUTION TIME','PASS/FAIL','FAILURE DESCRIPTION']]

        error_count = error_count + 1


        bi_team = 'ana.barraza@fii-na.com'
        #bi_team = 'ana.barraza@fii-na.com; rodrigo.bedolla@fii-na.com; erik.carbajalr@fii-na.com; bryan.rodriguez@fii-na.com'
        error_subject = 'TESTING - ATPS WO DETAIL SUMMARY REPORT'

        if error_count <= 5:

            if error_count == 1:

                send_mail_alert(bi_team, error_subject+' ERROR' , 'Proximo intento en 5 min | numero de intento: '+str(error_count), df_table)

            else:

                send_mail_alert(bi_team, error_subject+' ERROR' , 'Proximo intento en 5 min | numero de intento: '+str(error_count), df_table)

        elif error_count <= 10:

            send_mail_alert(bi_team, error_subject+' ERROR' , 'Proximo intento en 30 min | numero de intento: '+str(error_count), df_table)
        
        else:

            send_mail_alert(bi_team, error_subject+' CRITICAL ERROR' , 'Ultimo intento | numero de intento: '+str(error_count), df_table)

        if error_count <= 10:

            schedule.clear()
            job()  
job()