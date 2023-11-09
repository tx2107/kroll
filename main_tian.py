import datetime
from math import pow
import pandas as pd
import argparse
import numpy as np

OUTPUT_NAMES_LIST = ['Months', 'Paymnt_Count', 'Paydate', 'Scheduled_Principal', 'Scheduled_Interest',
                     'Scheduled_Balance',
                     'Prepay_Speed', 'Default_Rate', 'Recovery', 'Servicing_CF', 'Earnout_CF', 'Balance', 'Principal',
                     'Default', 'Prepay', 'Interest_Amount', 'Total_CF']
INPUT_NAMES_LIST = ['Valuation_Date', 'Grade', 'Issue_Date', 'Term', 'CouponRate',
                    'Invested', 'Outstanding_Balance', 'Recovery_Rate',
                    'Purchase_Premium', 'Servicing_Fee', 'Earnout_Fee',
                    'Deafult_Multiplier', 'Prepay_Multiplier']


# fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
def create_valid_date(year, month, day):
    try:
        return datetime.datetime(year, month, day)
    except ValueError:
        # If the day is out of range for the month, roll over to next month
        # Subtract the number of days in the current month and then add the desired days
        if month == 12:  # Roll-over to next year if month is December
            days_in_month = (datetime.datetime(year + 1, 1, 1) - datetime.timedelta(days=1)).day
            day_diff = day - days_in_month
            return datetime.datetime(year + 1, 1, day_diff)
        else:
            # Roll-over to next month
            days_in_month = (datetime.datetime(year, month + 1, 1) - datetime.timedelta(days=1)).day
            day_diff = day - days_in_month
            return datetime.datetime(year, month + 1, day_diff)


def get_prepay_speed(xlsx_df, sheet_name, term, month):
    # Read the "Prepay" sheet into a DataFrame
    df = pd.read_excel(xlsx_df, sheet_name=sheet_name)

    if month == 1:
        return 0
    try:
        return df[term][month - 1]
    except KeyError:
        return None


def get_default_rate(xlsx_df, sheet_name, term, grade, month):
    # Read the "Charged Off" sheet into a DataFrame
    df = pd.read_excel(xlsx_df, sheet_name=sheet_name)
    product = f"{term}-{grade}"
    try:
        return df[product][month - 1]
    except KeyError:
        return None


def pmt(rate, nper, pv):
    # Calculate the fixed payment (PMT)
    if rate == 0:
        return -1 * pv / nper
    else:
        pvif = pow(1 + rate, nper)
        pmt_value = rate / (pvif - 1) * -(pv * pvif)

        return pmt_value


def ipmt(rate, per, nper, pv, ipmt_dict=dict()):
    # Calculate the interest part of the payment
    if per in ipmt_dict:
        return ipmt_dict[per]
    if per == 0:
        return 0
    interest_payment = -1 * (pv + sum_principle_paid(rate, per, nper, pv)) * rate
    ipmt_dict[per] = interest_payment
    return interest_payment


def ppmt(rate, per, nper, pv):
    # Calculate the principal part of the payment
    if per == 0:
        return 0
    payment = pmt(rate, nper, pv)
    principal_payment = payment - ipmt(rate, per, nper, pv)
    return principal_payment


def sum_principle_paid(rate, per, nper, pv):
    # Sum the principal paid over a number of periods.
    principle_paid = 0
    for i in range(1, per):
        principle_paid += ppmt(rate, i, nper, pv)
    return principle_paid


def npv_calculation(rate, cashflows):
    total_value = 0
    for i, cashflow in enumerate(cashflows):
        total_value += cashflow / (1 + rate) ** i
    return total_value


def irr_calculation(cashflows, iterations=1000, guess=0.1):
    """Computes the IRR by iteratively solving the NPV equation."""

    # Initial guess for IRR
    rate = guess

    # Loop until we reach our desired precision or the number of iterations
    for i in range(iterations):
        # Calculate NPV with the current rate
        net_present_value = npv_calculation(rate, cashflows)

        # Calculate the derivative (slope) of the NPV function at the current rate
        derivative_npv = sum([-i * cashflow / (1 + rate) ** (i + 1) for i, cashflow in enumerate(cashflows)])

        # Newton-Raphson method formula to improve our guess
        rate -= net_present_value / derivative_npv

        # Exit loop if we're close enough to zero
        if abs(net_present_value) < 1e-10:
            break

    return rate


def irr_flow_preparation(Valuation_Date: str = "12/31/2017", Grade: str = "C4", Issue_Date: str = "8/24/2015",
                         Term: int = 36,
                         CouponRate: float = .280007632124385,
                         Invested: float = 7500.00, Outstanding_Balance: float = 3228.61, Recovery_Rate: float = 0.08,
                         Purchase_Premium: float = 0.051422082, Servicing_Fee: float = 0.025,
                         Earnout_Fee: float = 0.025,
                         Deafult_Multiplier: float = 1.00, Prepay_Multiplier: float = 1.00,
                         xlsx_df=pd.ExcelFile('Loan IRR.xlsx')) -> float:
    out_put_dict = dict()
    # initiating all columns
    for name in OUTPUT_NAMES_LIST:
        if name != 'Paydate':
            out_put_dict[name] = list[float]()

    out_put_dict['Paydate'] = list[datetime.datetime]()
    out_put_dict['Months'] = list(range(1, Term + 2))
    out_put_dict['Paymnt_Count'] = list(range(Term + 1))
    date_format = "%Y-%m-%d"
    Issue_Date_data = datetime.datetime.strptime(Issue_Date, date_format)
    for month_i in out_put_dict['Months']:
        # fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
        Paydate_i = create_valid_date(Issue_Date_data.year + (Issue_Date_data.month + month_i - 1) // 12,
                                      (Issue_Date_data.month + month_i - 1 - 1) % 12 + 1, Issue_Date_data.day)
        out_put_dict['Paydate'].append(Paydate_i)

    rate = CouponRate / 12

    for month_i in out_put_dict['Months']:
        principal_payment = ppmt(rate, month_i - 1, Term, -1 * Invested)
        interest_payment = ipmt(rate, month_i - 1, Term, -1 * Invested)
        out_put_dict['Scheduled_Principal'].append(principal_payment)
        out_put_dict['Scheduled_Interest'].append(interest_payment)
        out_put_dict['Scheduled_Balance'].append(Invested - sum_principle_paid(rate, month_i, Term, -1 * Invested))
        out_put_dict['Prepay_Speed'].append(get_prepay_speed(xlsx_df, "Prepay", Term, month_i))
        out_put_dict['Default_Rate'].append(get_default_rate(xlsx_df, "Charged Off", Term, Grade, month_i))
        out_put_dict['Earnout_CF'].append(Earnout_Fee / 2 * Invested if month_i == 13 or month_i == 19 else 0)
        if month_i == 1:
            out_put_dict['Default'].append(0)
            out_put_dict['Prepay'].append(0)
            out_put_dict['Principal'].append(0)
            out_put_dict['Balance'].append(Invested)
            out_put_dict['Recovery'].append(0)
            out_put_dict['Servicing_CF'].append(0)
            out_put_dict['Interest_Amount'].append(0)
            out_put_dict['Total_CF'].append(-1 * Invested * (1 + Purchase_Premium))

        else:
            out_put_dict['Default'].append(
                out_put_dict['Balance'][-1] * out_put_dict['Default_Rate'][-2] * Deafult_Multiplier)
            out_put_dict['Prepay'].append(
                (out_put_dict['Balance'][-1] - (((out_put_dict['Balance'][-1] - out_put_dict['Scheduled_Interest'][
                    -1]) / out_put_dict['Scheduled_Balance'][-2]) * out_put_dict['Scheduled_Principal'][
                                                    -1])) * out_put_dict['Prepay_Speed'][-1] * Prepay_Multiplier)
            out_put_dict['Principal'].append(
                (out_put_dict['Balance'][-1] - out_put_dict['Default'][-1]) / out_put_dict['Scheduled_Balance'][-2] *
                out_put_dict['Scheduled_Principal'][-1] + out_put_dict['Prepay'][-1])
            out_put_dict['Balance'].append(
                out_put_dict['Balance'][-1] - out_put_dict['Default'][-1] - out_put_dict['Principal'][-1])
            out_put_dict['Recovery'].append(out_put_dict['Default'][-1] * Recovery_Rate)
            out_put_dict['Servicing_CF'].append(
                (out_put_dict['Balance'][-2] - out_put_dict['Default'][-1]) * Servicing_Fee / 12)
            out_put_dict['Interest_Amount'].append(
                (out_put_dict['Balance'][-2] - out_put_dict['Default'][-1]) * CouponRate / 12)
            out_put_dict['Total_CF'].append(
                out_put_dict['Principal'][-1] + out_put_dict['Interest_Amount'][-1] + out_put_dict['Recovery'][-1] -
                out_put_dict['Servicing_CF'][-1] - out_put_dict['Earnout_CF'][-1])

    result_data = list()
    for name in OUTPUT_NAMES_LIST:
        result_data.append(out_put_dict[name])

    # Convert to a pandas DataFrame
    df = pd.DataFrame(result_data).T
    df.columns = OUTPUT_NAMES_LIST
    # Set option to display all columns
    pd.set_option('display.max_columns', None)

    df['Paydate'] = pd.to_datetime(df['Paydate']).dt.date
    # Set option to display all rows
    pd.set_option('display.max_rows', None)
    # print(df.dtypes)
    print('------------------------------------full--data------------------------------------')
    print(df)
    print('---------------------------------------head---------------------------------------')
    # Print the first few rows in a nice format
    print(df.head())
    print('---------------------------------------tail---------------------------------------')
    # Print the last few rows in a nice format
    print(df.tail())
    print()
    print()
    result = irr_calculation(out_put_dict['Total_CF']) * 12 * 100
    return result


def main():
    import argparse
    input_dict = dict()
    parser = argparse.ArgumentParser(description='Process IRR.')
    for name in INPUT_NAMES_LIST:
        parser.add_argument('--{}'.format(name), type=str)
    # parser.add_argument('--grade', type=str, help='Grade')
    args = parser.parse_args()

    for name in INPUT_NAMES_LIST:
        arg_value = getattr(args, name, None)
        input_dict[name] = arg_value
    # read default value from 'Loan IRR.xlsx'
    xlsx_df_irr = pd.ExcelFile('Loan IRR.xlsx')
    df = pd.read_excel(xlsx_df_irr, sheet_name='IRR Calculation')
    for name in INPUT_NAMES_LIST[:-2]:
        indices = np.where(df.values == name)
        for row, col in zip(*indices):
            input_dict[name] = str(df.loc[row].iloc[col + 1])

    # Deafult Multiplier = , Prepay Multiplier =  is very unique
    for name in INPUT_NAMES_LIST[-2:]:
        print(name)
        if input_dict[name] is None:
            indices = np.where(df.values == name.replace("_", " ") + " = ")
            for row, col in zip(*indices):
                input_dict[name] = str(df.loc[row].iloc[col + 1])

    for name in INPUT_NAMES_LIST:
        print("{}---{}".format(name, input_dict[name]))
    result_irr = irr_flow_preparation(Valuation_Date=input_dict['Valuation_Date'][:10], Grade=input_dict['Grade'],
                                      Issue_Date=input_dict['Issue_Date'][:10], Term=int(input_dict['Term']),
                                      CouponRate=float(input_dict['CouponRate']),
                                      Invested=float(input_dict['Invested']),
                                      Outstanding_Balance=float(input_dict['Outstanding_Balance']),
                                      Recovery_Rate=float(input_dict['Recovery_Rate']),
                                      Purchase_Premium=float(input_dict['Purchase_Premium']),
                                      Servicing_Fee=float(input_dict['Servicing_Fee']),
                                      Earnout_Fee=float(input_dict['Earnout_Fee']),
                                      Deafult_Multiplier=float(input_dict['Deafult_Multiplier']),
                                      Prepay_Multiplier=float(input_dict['Prepay_Multiplier']), xlsx_df=xlsx_df_irr)
    print(f"The IRR is: {result_irr:.6f}%")  # 5.4903063297685728


if __name__ == "__main__":
    '''
    I AM USING PYTHON 3.11.
    I read default value from 'Loan IRR.xlsx'
    I strictly followed the wordings from the original excel you send to me. please make sure didn't change any wording from the excel 
    For example i used "Deafult Multiplier = " at B14 of the excel. even it is a typo.
    
    unless you input a different number like below:
    INPUT_NAMES_LIST = ['Valuation_Date', 'Grade', 'Issue_Date', 'Term', 'CouponRate',
                    'Invested', 'Outstanding_Balance', 'Recovery_Rate',
                    'Purchase_Premium', 'Servicing_Fee', 'Earnout_Fee',
                    'Deafult_Multiplier', 'Prepay_Multiplier']
    
    python3 main_tian.py --Valuation_Date xxx --Grade C4 --Deafult_Multiplier 1
    python3 main_tian.py --Valuation_Date xxx --Grade C2 --Deafult_Multiplier 2

    fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
    xlsx file was saved at the same directory of python codes
    xlsx = pd.ExcelFile('Loan IRR.xlsx')
    the content in the Excel will need to exactly match what you send me in the email.


    '''
    main()
