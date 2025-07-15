# Author: Ana Barraza Reyes
# Department: Order management
# Client: HPE Planning
# Ana Barraza, R. (June 16, 2025). emailSender (version No 1.1 | last update by Ana Barraza (June 18, 2025)). Cd. Juarez: Foxconn eCMMS S.A. DE C.V. V1.1

 
#IMPORTING LIBRARIES---------------------------------------------------------------------------------------
from email.message import EmailMessage
from My_Book import *
import win32com.client as win32
from Signature import get_html_signature


# FUNCTION TO SEND EMAIL WITH EXCEL ATTACHMENTS---------------------------------------------------------
def send_mail_with_excel(recipient_email, email_copy,subject,file_attachments):

    #The variables are defined in read_files.py---------------------------------------------------------
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient_email # Email list
    mail.CC = email_copy # A quien va copiado
    mail.Subject = subject # Asunto del correo

    # Mail that show the message will appear in the body -----------------------------------------------
    mail.HTMLBody = """<html>
                        <head>
                            <style>
                            </style>
                        </head>
                            <body>
                                <p><b>Buen día Team!</b><br>

                                <!-- Format date 2 send the date in  MM/DD/YYYY  format -->
                                Adjunto el ATPS del dia de hoy ("""+str(format_date(2)) +""")</p><br>
                                <p>Les deseo un excelente dia,<br>Saludos.</p>
                                <br>

                                <!-- CALLING SIGNATURE HTML -->
                                """ + get_html_signature('Ana Barraza','ana.barraza@fii-na.com','') + """ <br>
                            </body>
                    </html>"""
                    
    # Sending the email with attachments------------------------------------------------------------------
    for file_path in file_attachments:
        mail.Attachments.Add(file_path)
        mail.Send()

