import pandas as pd

PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Galatea/output/'
FILE_DATA_XLSX = 'Result.xlsx'
FILE_DATA_INCORRECT = 'incorrect_xlsx.xlsx'

temp_data = 'c:/Users/1/PycharmProjects/Galatea/data/data.xlsx'

# df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_INCORRECT}')
# df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')
df = pd.read_excel(temp_data)

print(df.dtypes)
print((df.dtypes.tolist()))

