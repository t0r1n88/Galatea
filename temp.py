import pandas as pd

PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Galatea/output/'
FILE_DATA_XLSX = 'Result.xlsx'
FILE_DATA_INCORRECT = 'incorrect_xlsx.xlsx'

# df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_INCORRECT}')
df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')

print(type(df))