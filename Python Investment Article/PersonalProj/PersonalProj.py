import xlwings as xw
import numpy_financial as npf


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

@xw.func
def final_calculator(interest_rate, sip, number_of_years, principal_amount=0):
    months = int(number_of_years * 12)
    invested_amt = principal_amount
    monthly_interest = interest_rate / 12
    
    for month in range(months):
        invested_amt = invested_amt + invested_amt * monthly_interest
        invested_amt  += sip
    return invested_amt



if __name__ == "__main__":
    xw.Book("PersonalProj.xlsm").set_mock_caller()
    main()
