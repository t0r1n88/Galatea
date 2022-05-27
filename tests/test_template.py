import pytest
import pandas
import os
from Galatea_main import *


PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Galatea/output/'

FILE_DATA_XLSX = 'Result.xlsx'
FILE_DATA_INCORRECT = 'incorrect_xlsx.xlsx'


def test_create_output_file():
    """
    Дано:Путь к созданному файлу и название созданного файла
    Когда: Функция processing_data завершает свою работу
    Тогда: По указанному путь должен быть создан файл с указанным названием

    """
    processing_data(PATH_FOLDER_DATA, FILE_DATA_XLSX)
    assert os.path.isfile(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')

def test_opened_output_xlsx_file():
    """
    Дано:Путь к созданному файлу и название созданного файла
    Когда: Функция processing_data завершает свою работу
    Тогда: Файл который создан после работы функции processing_data должен корректно открываться для обработки pandas.
    Для проверки используется метод isinstance
    """
    # test_df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_INCORRECT}')
    test_df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')
    assert  isinstance(test_df,pandas.core.frame.DataFrame)



