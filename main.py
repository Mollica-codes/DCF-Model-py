import json
from urllib.request import urlopen
import pandas as pd
# from openpyxl import *
# make sure pandas, openpyxl, and urllib are installed

api_key = "Add Financial Modelling Prep API key"

ticker = input("Enter company ticker symbol: ")

# cash flow statement df
cf_url = f'https://financialmodelingprep.com/api/v3/cash-flow-statement/{ticker}?limit=5&apikey={api_key}'
cf_response = urlopen(cf_url)
cf_data = cf_response.read().decode("utf-8")
cf_data = json.loads(cf_data)
cf_df = pd.DataFrame(cf_data)
cf_df = cf_df[["depreciationAndAmortization", "changeInWorkingCapital", "investmentsInPropertyPlantAndEquipment"]]

depreciationGrowth = []
averageDepreciationGrowth = 0
nwcGrowth = []
averageNwcGrowth = 0
capexGrowth = []
averageCapexGrowth = 0
for i in range(4):
    depreciationGrowth.append("{:.2f}".format(((cf_df.iloc[i, 0] - cf_df.iloc[i + 1, 0]) / cf_df.iloc[i, 0]) * 100))
    averageDepreciationGrowth += float(depreciationGrowth[i])

    nwcGrowth.append("{:.2f}".format(((cf_df.iloc[i, 1] - cf_df.iloc[i + 1, 1]) / cf_df.iloc[i, 1]) * 100))
    averageNwcGrowth += float(nwcGrowth[i])

    capexGrowth.append("{:.2f}".format(((cf_df.iloc[i, 2] - cf_df.iloc[i + 1, 2]) / cf_df.iloc[i, 2]) * 100))
    averageCapexGrowth += float(capexGrowth[i])

averageDepreciationGrowth = averageDepreciationGrowth / 4
averageNwcGrowth = averageNwcGrowth / 4
averageCapexGrowth = averageCapexGrowth / 4

depreciationGrowth.append(0)
nwcGrowth.append(0)
capexGrowth.append(0)

cf_df['depreciationGrowth'] = depreciationGrowth
cf_df['nwcGrowth'] = nwcGrowth
cf_df['capexGrowth'] = capexGrowth

cols = list(cf_df.columns.values)
cf_df = cf_df[cols[0:1] + [cols[-3]] + cols[1:2] + [cols[-2]] + cols[2:3] + [cols[-1]]]
# income statement df
i_url = f'https://financialmodelingprep.com/api/v3/income-statement/{ticker}?limit=5&apikey={api_key}'
i_response = urlopen(i_url)
data = i_response.read().decode("utf-8")
data = json.loads(data)
df = pd.DataFrame(data)

df = df[["calendarYear", 'revenue', 'incomeBeforeTax', 'incomeTaxExpense']]
df['taxRateFromEBIT'] = df.loc[:, 'incomeTaxExpense'] / df.loc[:, 'incomeBeforeTax']
df['percentOfRev'] = df.loc[:, 'incomeBeforeTax'] / df.loc[:, 'revenue']

revenueGrowth = []
averageGrowth = 0
for i in range(4):
    revenueGrowth.append(float((df.iloc[i, 1] - df.iloc[i + 1, 1]) / df.iloc[i, 1]))
    averageGrowth += float(revenueGrowth[i])
averageGrowth = averageGrowth / 4

revenueGrowth.append(0)
df['revenueGrowth'] = revenueGrowth

cols = list(df.columns.values)
df = df[cols[0:2] + [cols[-1]] + cols[2:3] + [cols[-2]] + cols[3:5]]

averageEBIT = df.iloc[:, 4].mean()

historical_dcf = pd.concat([df, cf_df], axis=1, join='inner')

historical_dcf["DandAPercentofRev"] = historical_dcf.loc[:, 'depreciationAndAmortization'] / historical_dcf.loc[:, 'revenue']
historical_dcf["NWCPercentofRev"] = historical_dcf.loc[:, 'changeInWorkingCapital'] / historical_dcf.loc[:, 'revenue']
historical_dcf["CapExPercentofRev"] = historical_dcf.loc[:, 'investmentsInPropertyPlantAndEquipment'] / historical_dcf.loc[:, 'revenue']

cols = list(historical_dcf.columns.values)
historical_dcf = historical_dcf[cols[0:8] + [cols[-3]] + cols[8:10] + [cols[-2]] + cols[10:12] + [cols[-1]]]

avgDepreciationPofR = historical_dcf.iloc[:, 8].mean()
avgNWCPofR = historical_dcf.iloc[:, 11].mean()
avgCapexPofR = historical_dcf.iloc[:, 14].mean()
avgTaxPofEBIT = historical_dcf.iloc[:, 6].mean()

historical_dcf = historical_dcf.sort_values('calendarYear')
historical_dcf = historical_dcf.reset_index(drop=True)
pd.set_option('display.max_columns', None)
pd.options.display.float_format = '{:.3f}'.format

print(historical_dcf)
print("This is the average revenue growth: " + "{:.2f}".format(averageGrowth * 100) + "%")
print("This is the average EBIT growth: " + "{:.2f}".format(avgTaxPofEBIT * 100) + "%")
print("this is the average tax percent of EBIT: " + "{:.2f}".format(averageEBIT * 100) + "%")
print("This is the average Depreciation percent of revenue: " + "{:.2f}".format(avgDepreciationPofR * 100) + "%")
print("This is the average NWC percent of revenue: " + "{:.2f}".format(avgNWCPofR * 100) + "%")
print("This is the average CapEx percent of revenue: " + "{:.2f}".format(avgCapexPofR * 100) + "%")
print("Enter assumptions for DCF model(use whole numbers like 8 for 8%)")
rgr = (int(input("Revenue Growth rate: ")) / 100) + 1
ebitpofr = int(input("EBIT as percent of revenue: ")) / 100
tax_rate = int(input("Tax rate: ")) / 100
depre = int(input("Depreciation as percent of revenue: ")) / 100
nwcpofr = int(input("Net working capital as percent of revenue: ")) / 100
capex = int(input("Capital Expenditures as percent of revenue: ")) / 100
wacc = int(input("WACC: ")) / 100
tgr = int(input("Terminal Growth Rate: ")) / 100

for i in range(5):
    historical_dcf.loc[len(historical_dcf.index)] = [int(historical_dcf.loc[4 + i, 'calendarYear']) + 1,
                                                     int(historical_dcf.loc[4 + i, 'revenue']) * rgr,
                                                     rgr - 1,
                                                     0,
                                                     ebitpofr,
                                                     0,
                                                     tax_rate,
                                                     0,
                                                     depre,
                                                     0,
                                                     0,
                                                     nwcpofr,
                                                     0,
                                                     0,
                                                     capex]


historical_dcf.loc[5:9, 'incomeBeforeTax'] = historical_dcf.loc[5:9, 'revenue'] * historical_dcf.loc[5:9, 'percentOfRev']
historical_dcf.loc[5:9, 'incomeTaxExpense'] = historical_dcf.loc[5:9, 'incomeBeforeTax'] * historical_dcf.loc[5:9, 'taxRateFromEBIT']
historical_dcf.loc[5:9, 'depreciationAndAmortization'] = historical_dcf.loc[5:9, 'revenue'] * historical_dcf.loc[5:9, 'DandAPercentofRev']
historical_dcf.loc[5:9, 'changeInWorkingCapital'] = historical_dcf.loc[5:9, 'revenue'] * historical_dcf.loc[5:9, 'NWCPercentofRev']
historical_dcf.loc[5:9, 'investmentsInPropertyPlantAndEquipment'] = historical_dcf.loc[5:9, 'revenue'] * historical_dcf.loc[5:9, 'CapExPercentofRev']

historical_dcf['FreeCashFlow'] = ''
historical_dcf.loc[5:9, 'FreeCashFlow'] = historical_dcf.loc[5:9, 'revenue'] - historical_dcf.loc[5:9, 'incomeBeforeTax'] - historical_dcf.loc[5:9, 'incomeTaxExpense'] + historical_dcf.loc[5:9, 'depreciationAndAmortization'] - historical_dcf.loc[5:9, 'changeInWorkingCapital'] - historical_dcf.loc[5:9, 'investmentsInPropertyPlantAndEquipment']
historical_dcf['time'] = ''
historical_dcf.loc[5:9, 'time'] = historical_dcf.loc[5:9, 'calendarYear'] - int(historical_dcf.loc[4, 'calendarYear'])
historical_dcf['presentValueFCF'] = ''
historical_dcf.loc[5:9, 'presentValueFCF'] = (historical_dcf.loc[5:9, 'FreeCashFlow'] / pow(1 + wacc, historical_dcf.loc[5:9, 'time']))
historical_dcf['PVTerminalValue'] = ""
historical_dcf.loc[9, 'PVTerminalValue'] = ((historical_dcf.loc[9, 'FreeCashFlow'] * (1 + tgr)) / (wacc - tgr)) / (pow(1 + wacc, historical_dcf.loc[9, 'time']))
historical_dcf['EnterpriseValue'] = ''
historical_dcf.loc[9, 'EnterpriseValue'] = historical_dcf.loc[5:9, 'presentValueFCF'].sum() + historical_dcf.loc[9, 'PVTerminalValue']

# stock data df
stock_url = f'https://financialmodelingprep.com/api/v3/enterprise-values/{ticker}?limit=5&apikey={api_key}'
response = urlopen(stock_url)
stock_data = response.read().decode("utf-8")
stock_data = json.loads(stock_data)
stock_df = pd.DataFrame(stock_data)
sharesOutstanding = stock_df.loc[0, 'numberOfShares']
# balance sheet df
balance_url = f'https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}?limit=5&apikey={api_key}'
b_response = urlopen(balance_url)
balance_data = b_response.read().decode("utf-8")
balance_data = json.loads(balance_data)
balance_df = pd.DataFrame(balance_data)
netDebt = balance_df.loc[0, 'netDebt']

historical_dcf['netDebt'] = ''
historical_dcf.loc[9, 'netDebt'] = netDebt
historical_dcf['sharesOutstanding'] = ''
historical_dcf.loc[9, 'sharesOutstanding'] = sharesOutstanding
historical_dcf['EquityValue'] = ''
historical_dcf.loc[9, 'EquityValue'] = historical_dcf.loc[9, 'EnterpriseValue'] - historical_dcf.loc[9, 'netDebt']
historical_dcf['sharePrice'] = ''
historical_dcf.loc[9, 'sharePrice'] = historical_dcf.loc[9, 'EquityValue'] / historical_dcf.loc[9, 'sharesOutstanding']

enterpriseVal = historical_dcf.loc[9, 'EnterpriseValue']
equityVal = historical_dcf.loc[9, 'EquityValue']
implied_share_price = historical_dcf.loc[9, 'sharePrice']

final_df = historical_dcf.drop(columns=['depreciationGrowth', 'nwcGrowth'])
pd.set_option('display.max_columns', None)
pd.options.display.float_format = '{:.3f}'.format
print(final_df)

print("Enterprise Value: " + "{:.2f}".format(enterpriseVal))
print("Equity Value: " + "{:.2f}".format(equityVal))
print("Implied Share Price: " + "{:.2f}".format(implied_share_price))

final_df['WACC'] = ""
final_df.loc[9, 'WACC'] = wacc
final_df["TGR"] = ""
final_df.loc[9, "TGR"] = tgr
excel = input("Do you want to export to Excel (y/n): ")


def export_to_excel():
    if excel == 'y':
        excel_name = input("What do you want to name your file? (add .xlsx extension) ")
        data_to_excel = pd.ExcelWriter(excel_name)
        final_df.to_excel(data_to_excel, sheet_name=f"{ticker}_DCF", index=False)
        data_to_excel.save()
        print("Your file has been created and saved!")
    else:
        print("I hope you found this model useful. Thanks.")


export_to_excel()

