'''
    I AM USING PYTHON 3.11.
    
    I strictly followed the wordings from the original excel you send to me. please make sure didn't change any wording from the excel 
    For example I used "Deafult Multiplier = " at B14 of tab of IRR Calculation of the excel. even it is a typo.
    Furthermore I accepted "Default_Multiplier" from the input
    
    I read default value from 'Loan IRR.xlsx'    UNLESS you input a different number like below:
    date_format = "%Y-%m-%d" or "%m/%d/%Y" if you want input a different 'Issue_Date'
    INPUT_NAMES_LIST = ['Valuation_Date', 'Grade', 'Issue_Date', 'Term', 'CouponRate',
                    'Invested', 'Outstanding_Balance', 'Recovery_Rate',
                    'Purchase_Premium', 'Servicing_Fee', 'Earnout_Fee',
                    'Deafult_Multiplier', 'Default_Multiplier', 'Prepay_Multiplier']
    
    i.e.
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1
    python3 main_tian.py --Valuation_Date xxx --Grade C2 --Default_Multiplier 2
    python3 main_tian.py --Valuation_Date xxx --Grade C1 --Default_Multiplier 2 --Issue_Date 2023-08-09
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1 --Issue_Date 10/07/2023
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1 --Invested 1000

    fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
    
    xlsx file was saved at the same directory of python codes
    xlsx = pd.ExcelFile('Loan IRR.xlsx')
    the content in the Excel will need to exactly match what you send me in the email.
    OUTPUT_NAMES_LIST = ['Months', 'Paymnt_Count', 'Paydate', 'Scheduled_Principal', 'Scheduled_Interest',
                     'Scheduled_Balance',
                     'Prepay_Speed', 'Default_Rate', 'Recovery', 'Servicing_CF', 'Earnout_CF', 'Balance', 'Principal',
                     'Default', 'Prepay', 'Interest_Amount', 'Total_CF']


    '''
