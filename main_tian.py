import datetime
from math import pow
import pandas as pd


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


def irr_flow_preparation(Valuation_Date: str = "12/31/2017", Grade: str = "C4", Issue_Date="8/24/2015", Term: int = 36,
                         CouponRate: float = .280007632124385,
                         Invested: float = 7500.00, Outstanding_Balance: float = 3228.61, Recovery_Rate: float = 0.08,
                         Purchase_Premium: float = 0.051422082, Servicing_Fee: float = 0.025,
                         Earnout_Fee: float = 0.025,
                         Deafult_Multiplier: float = 1.00, Prepay_Multiplier: float = 1.00,
                         xlsx_df=pd.ExcelFile('Loan IRR.xlsx')) -> float:
    names_list = ['Months', 'Paymnt_Count', 'Paydate', 'Scheduled_Principal', 'Scheduled_Interest', 'Scheduled_Balance',
                  'Prepay_Speed', 'Default_Rate', 'Recovery', 'Servicing_CF', 'Earnout_CF', 'Balance', 'Principal',
                  'Default', 'Prepay', 'Interest_Amount', 'Total_CF']

    Months = list(range(1, Term + 2))
    Paymnt_Count = list(range(Term + 1))
    date_format = "%m/%d/%Y"
    Issue_Date_data = datetime.datetime.strptime(Issue_Date, date_format)
    Paydate = list()
    for month_i in Months:
        # fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
        Paydate_i = create_valid_date(Issue_Date_data.year + (Issue_Date_data.month + month_i - 1) // 12,
                                      (Issue_Date_data.month + month_i - 1 - 1) % 12 + 1, Issue_Date_data.day)
        Paydate.append(Paydate_i)

    rate = CouponRate / 12
    Scheduled_Principal = list()
    Scheduled_Interest = list()
    Scheduled_Balance = list()
    Prepay_Speed = list()
    Default_Rate = list()
    Default = list()
    Principal = list()
    Balance = list()
    Prepay = list()
    Recovery = list()
    Servicing_CF = list()
    Earnout_CF = list()
    Interest_Amount = list()
    Total_CF = list()
    for month_i in Months:
        total_payment = pmt(rate, Term, -1 * Invested)
        principal_payment = ppmt(rate, month_i - 1, Term, -1 * Invested)
        interest_payment = ipmt(rate, month_i - 1, Term, -1 * Invested)
        Scheduled_Principal.append(principal_payment)
        Scheduled_Interest.append(interest_payment)
        Scheduled_Balance.append(Invested - sum_principle_paid(rate, month_i, Term, -1 * Invested))
        Prepay_Speed.append(get_prepay_speed(xlsx_df, "Prepay", Term, month_i))
        Default_Rate.append(get_default_rate(xlsx_df, "Charged Off", Term, Grade, month_i))
        Earnout_CF.append(Earnout_Fee / 2 * Invested if month_i == 13 or month_i == 19 else 0)
        if month_i == 1:
            Default.append(0)
            Prepay.append(0)
            Principal.append(0)
            Balance.append(Invested)
            Recovery.append(0)
            Servicing_CF.append(0)
            Interest_Amount.append(0)
            Total_CF.append(-1 * Invested * (1 + Purchase_Premium))

        else:
            Default.append(Balance[-1] * Default_Rate[-2])
            Prepay.append(
                (Balance[-1] - (((Balance[-1] - Scheduled_Interest[-1]) / Scheduled_Balance[-2]) * Scheduled_Principal[
                    -1])) * Prepay_Speed[-1] * Prepay_Multiplier)
            Principal.append(
                (Balance[-1] - Default[-1]) / Scheduled_Balance[-2] * Scheduled_Principal[-1] + Prepay[-1])
            Balance.append(Balance[-1] - Default[-1] - Principal[-1])
            Recovery.append(Default[-1] * Recovery_Rate)
            Servicing_CF.append((Balance[-2] - Default[-1]) * Servicing_Fee / 12)
            Interest_Amount.append((Balance[-2] - Default[-1]) * CouponRate / 12)
            Total_CF.append(Principal[-1] + Interest_Amount[-1] + Recovery[-1] - Servicing_CF[-1] - Earnout_CF[-1])

    result_data = list()
    for name in names_list:
        result_data.append(locals()[name])

    # Convert to a pandas DataFrame
    df = pd.DataFrame(result_data).T
    df.columns = names_list
    # Set option to display all columns
    pd.set_option('display.max_columns', None)

    # Set option to display all rows
    pd.set_option('display.max_rows', None)
    # Assuming `data_frame` is your large DataFrame
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
    result = irr_calculation(Total_CF) * 12 * 100
    return result


if __name__ == "__main__":
    '''
    assumption: date_format is a string with format of  "%m/%d/%Y"
    fixed the date of 31st issues to match the DATE function in excel in case Issue_Date_data.day = 31
    xlsx file was saved at the same directory of python codes
    xlsx = pd.ExcelFile('Loan IRR.xlsx')
    the content in the Excel will need to exactly match what you send me in the email.
    '''
    xlsx_df_irr = pd.ExcelFile('Loan IRR.xlsx')
    result_irr = irr_flow_preparation(Valuation_Date="12/31/2017", Grade="C4", Issue_Date="8/24/2015", Term=36,
                                      CouponRate=.280007632124385,
                                      Invested=7500.00, Outstanding_Balance=3228.61, Recovery_Rate=0.08,
                                      Purchase_Premium=0.051422082, Servicing_Fee=0.025, Earnout_Fee=0.025,
                                      Deafult_Multiplier=1.00, Prepay_Multiplier=1.00, xlsx_df=xlsx_df_irr)
    print(f"The IRR is: {result_irr:.6f}%")  # 5.4903063297685728
