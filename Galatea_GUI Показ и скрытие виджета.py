import tkinter
import sys

import pandas as pd
import qrcode
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()

def select_file_docx():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_docx
    file_docx = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx
    # Получаем путь к файлу
    file_data_xlsx = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def processing_qr_code():
    """
    Фугкция для обработки данных
    :return:
    """
    pass



if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x860')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_qr_code = ttk.Frame(tab_control)
    tab_control.add(tab_qr_code, text='Создание QR кодов')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_qr_code,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_qr_code,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data = Button(tab_qr_code, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_qr_code, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    # создаем переключатели
    # Определяем текстовую переменную
    name_column = StringVar()
    checkbox_type = IntVar()
    entry_id = Entry(tab_qr_code,textvariable=name_column)

    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type = LabelFrame(tab_qr_code, text='1) Выберите режим')
    frame_rb_type.grid(column=0, row=1, padx=10)
    entry_id.grid(column=0, row=4, padx=10, pady=10)

    #
    Radiobutton(frame_rb_type, text='А) Обработка стандартной таблицы', variable=checkbox_type,
                value=0, command=lambda: entry_id.grid_remove()).pack()
    Radiobutton(frame_rb_type, text='Б) Обработка произвольной таблицы', variable=checkbox_type,
                value=1, command=lambda: entry_id.grid()).pack()


    entry_id.grid(column=0, row=4, padx=10, pady=10)


    #Создаем кнопку обработки данных

    btn_proccessing_qr = Button(tab_qr_code, text='3) Создать QR', font=('Arial Bold', 20),
                                  command=processing_qr_code
                                  )
    btn_proccessing_qr.grid(column=0, row=6, padx=10, pady=10)

    entry_id.grid_remove() # удаляем поле ввода фио чтобы его не было видно
    window.mainloop()