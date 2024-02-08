import xlwings as xw
import pandas as pd
from nsepython import *
from datetime import datetime, timedelta



def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

@xw.func
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    return 2 * (x + y)


@xw.func
@xw.ret(expand='table')
def indexhistoryforweek(symbol):
    wb = xw.Book.caller()
    sheet = wb.sheets[1]
    start_date = (datetime.today() - timedelta(days=7)).strftime("%m/%d/%Y")
    end_date = datetime.now().strftime("%m/%d/%Y")
    df = pd.DataFrame(index_history(symbol,start_date,end_date))
    return df

@xw.func
def gethighofadate(symbol, dt):
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    start_date = (datetime.today() - timedelta(days=7)).strftime("%m/%d/%Y")
    end_date = datetime.now().strftime("%m/%d/%Y")
    df = pd.DataFrame(index_history(symbol,start_date,end_date))
    dt = dt.strftime("%d %b %Y")
    filt = (df['HistoricalDate'] == dt)
    try:
        result = df.loc[filt,'HIGH'].values[0]
        return float(result)
    except IndexError:
        return "NA"
    except ValueError as ve:
        return ve
    
@xw.func
def getlowofadate(symbol, dt):
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    start_date = (datetime.today() - timedelta(days=7)).strftime("%m/%d/%Y")
    end_date = datetime.now().strftime("%m/%d/%Y")
    df = pd.DataFrame(index_history(symbol,start_date,end_date))
    dt = dt.strftime("%d %b %Y")
    filt = (df['HistoricalDate'] == dt)
    try:
        result = df.loc[filt,'LOW'].values[0]
        return float(result)
    except IndexError:
        return "NA"

@xw.func
@xw.ret(expand='table')
def derivativehistorycall(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType):
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    start_date = start_date.strftime("%d-%m-%Y")
    end_date = end_date.strftime("%d-%m-%Y")
    expiry_date = expiry_date.strftime("%d-%b-%Y")
    df = pd.DataFrame(derivative_history(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType))
    return df

@xw.func
def derivativehistorycallgetlowvalue(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType):
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    start_date = start_date.strftime("%d-%m-%Y")
    end_date = end_date.strftime("%d-%m-%Y")
    expiry_date = expiry_date.strftime("%d-%b-%Y")
    try:
        df = pd.DataFrame(derivative_history(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType))
        result = (df['FH_TRADE_LOW_PRICE'].min())
        return float(result)
    except:
        return "NA"

@xw.func
def derivativehistoryputgetlowvalue(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType):
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    start_date = start_date.strftime("%d-%m-%Y")
    end_date = end_date.strftime("%d-%m-%Y")
    expiry_date = expiry_date.strftime("%d-%b-%Y")
    try:
        df = pd.DataFrame(derivative_history(symbol,start_date,end_date,instrumentType,expiry_date,strikePrice,optionType))
        result = (df['FH_TRADE_LOW_PRICE'].max())
        return float(result)
    except:
        return "NA"


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("test.xlsm").set_mock_caller()
    main()
