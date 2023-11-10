    '''
    I am using Python 3.11.

    The xlsx file was saved in the same directory as the Python codes:
    xlsx = pd.ExcelFile('Loan IRR.xlsx')
    The content in the Excel file will need to exactly match what you sent me in the email.
    
    I strictly adhered to the wording from the original Excel file you sent to me. 
    Please ensure that no wording from the Excel file has been altered. 
    For example I used "Deafult Multiplier = " at B14 of tab of IRR Calculation of the excel. even it is a typo.
    Furthermore I accepted "Default_Multiplier" from the input
    
    I read default value from 'Loan IRR.xlsx'    UNLESS you input a different number as shown below:
    date_format = "%Y-%m-%d" or "%m/%d/%Y" if you want input a different 'Issue_Date'
    INPUT_NAMES_LIST = ['Valuation_Date', 'Grade', 'Issue_Date', 'Term', 'CouponRate',
                    'Invested', 'Outstanding_Balance', 'Recovery_Rate',
                    'Purchase_Premium', 'Servicing_Fee', 'Earnout_Fee',
                    'Deafult_Multiplier', 'Default_Multiplier', 'Prepay_Multiplier']
    
    For example, you can use the following commands:
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1
    python3 main_tian.py --Valuation_Date xxx --Grade C2 --Default_Multiplier 2
    python3 main_tian.py --Valuation_Date xxx --Grade C1 --Default_Multiplier 2 --Issue_Date 2023-08-09
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1 --Issue_Date 10/07/2023
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1 --Invested 1000

    I've adjusted the handling of dates to align with the DATE function in Excel in case the 'Issue_Date_data.day' equals 31.
    

    OUTPUT_NAMES_LIST = ['Months', 'Paymnt_Count', 'Paydate', 'Scheduled_Principal', 'Scheduled_Interest',
                     'Scheduled_Balance',
                     'Prepay_Speed', 'Default_Rate', 'Recovery', 'Servicing_CF', 'Earnout_CF', 'Balance', 'Principal',
                     'Default', 'Prepay', 'Interest_Amount', 'Total_CF']


    '''
