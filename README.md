
HPE PLANNING WO DETAIL SUMMARY ---------------------------------------------------------------------------------------
Author: Ana Barraza Reyes; Order Management (OM) Department | Client: HPE Planning 902 | Released date: June 16, 2025 
 

Objective: The purpose of RPA is to provide a structured table with specific fields, properly filtered for later
emailing.


1 REQUIRED FIELDS:

    RECOMMIT QTY, CATEGORY, SEQ, NO DONATE, PC, SHIP TYPE, PO, COMPLEXITY, WO QTY, RDD, SINGLE, ORDER, MATERAL, 
    MATERIAL DESCRIPTION, PURCH.DOC., PO QTY, COMPONENT, REQMTS QTY, LONG POLE PN, DESCRIPTION, ETA, ACTUAL RECOVERY DATE
    NEW RECOVERY DATE, CTB


2 GETTING DATA:

    2.1 - ON SERVER
        2.1.1 - EXCEL File "Signal 855 Confirmation" We only need column 'RECOMMIT QTY' 
        2.1.2 - EXCEL File "BACKLOG SEQUENCE ANALYSIS" We only need columns 'CATEGORY','BACKLOG SEQUENCE', 'DNT'
        2.1.3 - EXCEL File "MPS 4AM" We only need columns 'PC', 'SHIP TYPE', 'PO', 'COMPLEXITY', 'QTY', 'RDD'. 
                'SINGLE_GATED_FLAG', 'MATERIAL P','SKU DESCRIPTION','DELIVERY D','ACTUAL SCHEDULED DATE'. 
                The Analysis Level is the column that will be used to filter the MPS 4AM file to remove the 
                duplicates

    2.2 GETTING FROM SAP
        2.2.1 - SAP Spreadsheet download from "ZATPRESULT" IN 902 "HPE" we only need columns 'Order', 'Material', 
                'Material description', 'Purch.doc.','Component','Reqmts qty'


3 SEQUENCE
    
    3.1 - Import libraries
    3.2 - Execute the function 
        3.2.1 - Execute function to eliminate all older data from our folder
        3.2.2 - Execute function to download neccesary data from SAP 902 (HPE) and save it as 
                zatpresult_woDetails.xls in our Files folder
        3.2.3 - Get column 'RECOMMIT QTY' from "Signal 855 Confirmation" (is in our server, just read file) 
        3.2.4 - Get columns 'WORK ORDER','CATEGORY','BACKLOG SEQUENCE','DNT' from BACKLOG SEQUENCE ANALYSIS (is in 
                our server just read file)
        3.2.5 - Get columns 'Order', 'Material', 'Material description', 'Purch.doc.','Component','Reqmts qty' 
                from zatpresult_woDetails.xls (is our Files folder prevosuly downloaded from SAP 902 | HPE)
        3.2.6 - GET columns 'PC', 'SHIP TYPE', 'PO', 'COMPLEXITY', 'QTY', 'RDD'. 'SINGLE_GATED_FLAG', 'MATERIAL P',
                'SKU DESCRIPTION','DELIVERY D','ACTUAL SCHEDULED DATE' from "MPS 4AM" (is in our server just read 
                file)
        3.2.7 - Make a copy of each Excel that will be used
        3.2.8 - Rename columns 'Order' to 'WORK ORDER' to avoid duplicates 
        3.2.9 - Eliminate decimals from columns to clear the final table
        3.2.10 - Convert Order to string to remove decimals
        3.2.11 - Drop duplicate columns from Indexes "Work Order"
        3.2.12 - Substract critcal short based on longest recovery date    
        3.2.13 - Make merge adding columns and removing blanks and nulls 
        3.2.14 - Adding new columns with default values (This columns they requested as blank as new columns)   
        3.2.15 - Reorganizate the order of fields, changing  names of each field to give the most exact result to user
                 as its original report and finally printing the new excel table as ATPS_(DATE)
        3.2.16 - Make the mail with the names of recipients, copy recipients, subect and mail
        3.2.17 - Send the mail


Foxconn eCMMS S.A. DE C.V. | Last update: July 15, 2025 -------------------------------------------------------------
