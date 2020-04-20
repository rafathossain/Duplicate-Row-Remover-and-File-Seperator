import pandas as pd

df = pd.read_excel(r'Combined.xlsx')

# print(df)

# Select all duplicate rows based on one column
# duplicateRowsDF = df[df.duplicated(['Email', 'Source'], keep='first')]

# print(duplicateRowsDF.head())

df = df.drop_duplicates(subset=['Email', 'Source'], keep="first")

# print(df)

duplicateRowsDF = df[df.duplicated(['Email'], keep=False)]

print("Duplicates: " + str(len(duplicateRowsDF)))

dropList = []

for row, column in duplicateRowsDF.iterrows():
    email = df.loc[row, 'Email']
    # print(email)
    newSource = df.loc[row, 'Source']
    for r2, c2 in duplicateRowsDF.iterrows():
        if r2 > row:
            if email == df.loc[r2, 'Email']:
                # print(df.loc[r2, 'Email'])
                newSource = newSource + ", " + df.loc[r2, 'Source']
                if r2 not in dropList:
                    dropList.append(r2)
    df.at[row, 'Source'] = newSource
    print(row)

for key in dropList:
    df.drop(key, inplace=True)

# print(df)

duplicateRowsDF = df[df.duplicated(['Email'], keep=False)]

print("After Complete Duplicates: " + str(len(duplicateRowsDF)))

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('No_Duplicate.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
