# Imports of libraries, pandas options

import pandas as pd
import lxml
import matplotlib.pyplot as plt
import xlsxwriter

pd.options.mode.chained_assignment = None   # default='warn', to avoid pd error when editing with for loop
# there's no for loops as of now, but it might be needed later


def deduplicate(df):  # all row titles are duplicated when scraped from marketwatche
    for index, row in df.iterrows():
        df.iat[index, 0] = row[0][:int(len(row[0])/2)]

# Creation of marketwatch data links


marketwatchdata = {}

print('================= Financial Data Scraper =================\n'
      'By Andrea Sapei - 2020\n\n'
      'This program is designed to scrape financial data from the\n'
      'Internet, using MarketWatch as a source, and put this data\n'
      'in an Excel(.xlsx) file for easy use in models or analysis\n')

ticker = input("Insert symbol to research: ")
# For stocks on NYSE or NASDAQ just write the ticker
# For stocks on other exchanges Marketwatch does not always have data
# If you want to find the data they might have, use 'ticker'?countrycode='countrycode'
# It is not very user friendly now but it'll get better

# ================ Data Scraping Begins ================

# Creates urls to begin scraping data
urlincome = 'https://www.marketwatch.com/investing/stock/' + ticker + '/financials'
urlbalance = 'https://www.marketwatch.com/investing/stock/' + ticker + '/financials/balance-sheet'
urlcashflow = 'https://www.marketwatch.com/investing/stock/' + ticker + '/financials/cash-flow'
urlkeydata = 'https://www.marketwatch.com/investing/stock/' + ticker
urlprofile = 'https://www.marketwatch.com/investing/stock/' + ticker + '/company-profile'

# Scrape of marketwatch data, not a perfect source but will do for this program

print("Gathering data...")
income = pd.read_html(urlincome)  # see above for actual urls
balancesheet = pd.read_html(urlbalance)
cashflow = pd.read_html(urlcashflow)
keydata = pd.read_html(urlkeydata)
profile = pd.read_html(urlprofile)

# general formatting

profile[0] = profile[4].rename(columns={'0': 'Financial ratios'})  # lots of generic data, ratios, insider transactions
profile[1] = profile[5].rename(columns={'0': 'Revenue ratios'})  # etc. etc., needs to be formatted in a better way
profile[2] = profile[6].rename(columns={'0': 'Assets and revenues'})  # than this.
profile[3] = profile[7].rename(columns={'0': 'Margin and returns'})
profile[4] = profile[8].rename(columns={'0': 'Debt to assets'})  # sometimes some parts also stop working for
# profile[5] = profile[9]                                             #some reason, this program isn't perfect
# profile[6] = profile[10].drop(columns = ['Unnamed: 0'])

income_statement = income[4].rename(columns={'Item  Item': 'Income statement'})

balance_sheet_assets = balancesheet[4].rename(columns={'Item  Item': 'Assets'})
balance_sheet_liabilities = balancesheet[5].rename(columns={'Item  Item': 'Liabilities and Equity'})

cashflow_operations = cashflow[4].rename(columns={'Item  Item': 'Cash-flow from Operating activities'})
cashflow_investing = cashflow[5].rename(columns={'Item  Item': 'Cash-flow from Investing activities'})
cashflow_financing = cashflow[6].rename(columns={'Item  Item': 'Cash-flow from Financing activities'})

dayprice = keydata[1].rename(columns={'Close': 'Stock price'})
performance = keydata[4].rename(columns={'0': 'Performance'})
competitors = keydata[5].rename(columns={'Name': 'Competitors'})

print("Deduplicating...")
deduplicate(income_statement)
deduplicate(balance_sheet_assets)
deduplicate(balance_sheet_liabilities)
deduplicate(cashflow_operations)
deduplicate(cashflow_investing)
deduplicate(cashflow_financing)

print("Formatting numbers...")
#after eons and eons, I was finally able to convert the B to billions, M to millions and K to thousands, efficiently
#conversion of income statement
income_statement.iloc[:, [1,2,3,4,5]] = income_statement.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)

income_statement.iloc[:, [1,2,3,4,5]] = '=' + income_statement.iloc[:, [1,2,3,4,5]].astype(str)

#conversion of balance sheet
balance_sheet_assets.iloc[:, [1,2,3,4,5]] = balance_sheet_assets.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)
balance_sheet_liabilities.iloc[:, [1,2,3,4,5]] = balance_sheet_liabilities.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)

balance_sheet_assets.iloc[:, [1,2,3,4,5]] = '=' + balance_sheet_assets.iloc[:, [1,2,3,4,5]].astype(str)
balance_sheet_liabilities.iloc[:, [1,2,3,4,5]] = '=' + balance_sheet_liabilities.iloc[:, [1,2,3,4,5]].astype(str)

#conversion of cashflow statement
cashflow_operations.iloc[:, [1,2,3,4,5]] = cashflow_operations.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)
cashflow_investing.iloc[:, [1,2,3,4,5]] = cashflow_investing.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)
cashflow_financing.iloc[:, [1,2,3,4,5]] = cashflow_financing.iloc[:, [1,2,3,4,5]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0'}, regex=True)


cashflow_operations.iloc[:, [1,2,3,4,5]] = '=' + cashflow_operations.iloc[:, [1,2,3,4,5]].astype(str)
cashflow_investing.iloc[:, [1,2,3,4,5]] = '=' + cashflow_investing.iloc[:, [1,2,3,4,5]].astype(str)
cashflow_financing.iloc[:, [1,2,3,4,5]] = '=' + cashflow_financing.iloc[:, [1,2,3,4,5]].astype(str)

#General data gets converted to num, to facilitate excel use

competitors.iloc[:,[2]] = competitors.iloc[:,[2]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0', '\$':'='}, regex=True)
dayprice.iloc[:,[0]] = dayprice.iloc[:,[0]].replace({'B': '*1000000', 'M': '*1000', 'K': '', '-':'0', '\$':'='}, regex=True)

#Saving model to Excel
try:
    datawriter = pd.ExcelWriter(ticker + '.xlsx', engine='xlsxwriter')   #xlsx automatically changes to ticker.xlsx
except:
    datawriter = pd.ExcelWriter('somethingwentwrong.xlsx', engine='xlsxwriter') #don't put weird stuff as ticker, there's no error handler


#formatting to make excel file easier to read will be added later on

print("Exporting...")
profile[0].to_excel(datawriter, sheet_name='Key data',startrow=2, startcol=4)
profile[1].to_excel(datawriter, sheet_name='Key data',startrow=13, startcol=-1)
profile[2].to_excel(datawriter, sheet_name='Key data',startrow=19, startcol=-1)
profile[3].to_excel(datawriter, sheet_name='Key data',startrow=2, startcol=8)
profile[4].to_excel(datawriter, sheet_name='Key data',startrow=2, startcol=12)
profile[5].to_excel(datawriter, sheet_name='Key data',startrow=13, startcol=4)
profile[6].to_excel(datawriter, sheet_name='Key data',startrow=13, startcol=9)

dayprice.to_excel(datawriter, sheet_name='Key data',startrow=2, startcol=-1)
performance.to_excel(datawriter, sheet_name='Key data',startrow=5, startcol=-1)
competitors.to_excel(datawriter, sheet_name='Key data',startrow=20, startcol=4)

income_statement.to_excel(datawriter, sheet_name='Analysis',startrow=2, startcol=-1)
balance_sheet_assets.to_excel(datawriter, sheet_name='Analysis',startrow=2, startcol=7)
balance_sheet_liabilities.to_excel(datawriter, sheet_name='Analysis',startrow=42, startcol=7)

cashflow_operations.to_excel(datawriter, sheet_name='Analysis',startrow=2, startcol=15)
cashflow_investing.to_excel(datawriter, sheet_name='Analysis',startrow=24, startcol=15)
cashflow_financing.to_excel(datawriter, sheet_name='Analysis',startrow=43, startcol=15)

try:
    datawriter.save()
    print('Output complete')
except:
    print('Saving to file failed')
