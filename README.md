# Daily-Report
This repository contain source code and an excel file where I automated a process of preparing Daily report of IMPS transactions data 
of my client as per their requirement and format.
The Daily report excel file contain 4 sheets in which IMPS transaction data is shared to the client on every 3-hourly basis. Please find
elaborated description and purpose of every worksheet as follows:
1. IMPS_Server_Status:  This worksheet shows a visual representation just to get a quick idea of various application server running status.

2. Hourly_Count: In this worksheet, IMPS transaction data is maintained separately for INWARD and OUTWARD transactions. For Outward transactions,
there are many different channal are used to process the transaction of different purpose such as Retail Internet Banking,  Corporate Internet
Banking etc. Hence the total, successful and unsuccessful transactions data of every hour is maintained in pre-specified format set by client.

3. Hourly_Txn_Error: In this sheet, the errors/reasons due to which transactions failed are maintain separately for INWARD and OUTWARD transactions
with their respective count and id for every hour which provides a overview of different business and technical errors.

4. EOD_Count: In this sheet, whole day's data of different status is maintained for INWARD and OUTWARD transactions.

So I have automated this report using Python programming language. I used 'openpyxl' module of python to work on Excel files. I used 
cx_oracle package to connect python with oracle database. The code automated the data extraction from db and filling the same in existing 
excel file as per format. Also the 'next day dated file creation.py' file code is used to generate the same empty excel file except 
EOD_count for next day.

Note: I have changed some of original names such channel names, table name etc with dummy names to prevent breach of any kind of client
data/terms. 
