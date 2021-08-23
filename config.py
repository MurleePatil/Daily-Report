PATH= 'C:\\Users\\7\\Desktop\\Daily reports\\2019\\November'
PATH1='C:\\Users\\7\\Desktop'
PATH2='C:\\Users\\7\\Desktop\\MNO channel count\\2019'
PATH3='C:\\Users\\7\\Desktop\\MNO count\\2019\\MNO channel error description'

HOST='localhost'
USER='root'
PASSWORD='root@123'
DATABASE='db1'
CHANNELS=['ABC','PQR','XYZ','MNO']


dict1={'ISSUING BANK CBS OR NODE OFFLINE':'Beneficiary bank node is offline and reversal is successful',
        'INVALID TRANSACTION TYPE AND REVERSAL IS SUCCESSFUL':'Invalid transaction type',
        'INVALID AMOUNT FIELD AND REVERSAL IS SUCCESSFUL':'Invalid amount field',       
        'A/C RESTRICTDebit freeze':'Debit freeze on respective account',
        'CLEARED BAL/FUNDS/DP NOT AVAILABLE':'Insufficient balance in remitter account',
        'BENEFICIARY ACCOUNT BLOCKED/FROZEN AND REVERSAL IS SUCCESSFUL':'Beneficiary account is frozen',
        'BENEFICIARY ACCOUNT IS CLOSED AND REVERSAL IS SUCCESSFUL':'Beneficiary account is closed',       
        '':'Null',
        'Invalid beneficiary account':'Invalid mmid/mob/account no',
        'ACCOUNT CLOSED MMID FAILURE':'Beneficiary account is closed',               
        'TRANSACTION EXCEEDS TOTAL CREDIT':'Total credit limit for the day breached the maximum threshold value',
        'ACCOUNT IS LOCKED-PLS TRY AGAIN':'Account was locked at CBS at that point of time',
        }

    


