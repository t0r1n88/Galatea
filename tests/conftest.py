import pytest
import pandas as pd
"""
Файл для фикстур
"""
PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Galatea/data/'
NAME_DATA_XLSX = 'data.xlsx'

@pytest.fixture
def test_df():
    """
    Фикстура для создания датафрейма из файла
    """
    df = pd.read_excel(f'{PATH_FOLDER_DATA}{NAME_DATA_XLSX}')
    return df

