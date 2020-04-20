import pandas as pd
import xlsxwriter
import math

country = "United Kingdom.xlsx"

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(country)
worksheet = workbook.add_worksheet()

df = pd.read_excel(r'No_Duplicate.xlsx')
df = df.fillna(0)

# print(df)

rowID = 0

worksheet.write(rowID, 0, 'Country')
worksheet.write(rowID, 1, 'Abbreviation')
worksheet.write(rowID, 2, 'Email')
worksheet.write(rowID, 3, 'Source')
worksheet.write(rowID, 4, 'Anomaly')
rowID += 1

for row, column in df.iterrows():
    filename = df.loc[row, 'File Name']
    if filename == country:
        if str(df.loc[row, 'Country']) == "0":
            worksheet.write(rowID, 0, "")
        else:
            worksheet.write(rowID, 0, df.loc[row, 'Country'])
        worksheet.write(rowID, 1, df.loc[row, 'Abbreviation'])
        email = df.loc[row, 'Email']
        email = email.strip()
        if len(email.split("@")) > 2:
            print(str(row) + " > " + email)
        worksheet.write(rowID, 2, email)
        worksheet.write(rowID, 3, df.loc[row, 'Source'])
        if str(df.loc[row, 'Anomaly']) == "0":
            worksheet.write(rowID, 4, "")
        else:
            worksheet.write(rowID, 4, df.loc[row, 'Anomaly'])
        rowID += 1
    # print(row)

workbook.close()
