from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk
import pandas as pd





def select_file_template_contracts():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_contracts
    name_file_template_contracts = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_contracts():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться договор  и генерация падежей фио
    :return: Путь к файлу с данными и словарь с просклоняемыми ФИО
    """
    global name_file_data_contracts
    # Получаем путь к файлу
    name_file_data_contracts = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))

def select_end_folder_contracts():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_contracts
    path_to_end_folder_contracts = filedialog.askdirectory()


def generate_contracts():
    """
    Функция для создания договоров
    :return:
    """
    try:
        # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
        reader = csv.DictReader(open(name_file_data_contracts), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_contracts)

            context = {'ФиоСлушателя': row['ФиоСлушателя'],
                       'ФиоСлушателяРодПадеж': row['ФиоСлушателяРодПадеж'],
                       'НомерДоговора': row['НомерДоговора'],
                       'ДатаПодписанияДоговора': row['ДатаПодписанияДоговора'],
                       'ДолжностьФиоРодительныйПадеж': row['ДолжностьФиоРодительныйПадеж'],
                       'Программа': row['Программа'], 'СрокВМесяцах': row['СрокВМесяцах'],
                       'Профессия': row['Профессия'], 'СрокВЧасах': row['СрокВЧасах'],
                       'ДатаНачалаЗанятий': row['ДатаНачалаЗанятий'],
                       'НачалоОбучения': row['НачалоОбучения'],
                       'КонецОбучения': row['КонецОбучения'], 'ПолнаяСтоимость': row['ПолнаяСтоимость'],
                       'ПерваяЧастьОплаты': row['ПерваяЧастьОплаты'],
                       'ДатаПервойОплаты': row['ДатаПервойОплаты'],
                       'ВтораяЧастьОплаты': row['ВтораяЧастьОплаты'],
                       'ДатаВторойОплаты': row['ДатаВторойОплаты'],
                       'ТретьяЧастьОплаты': row['ТретьяЧастьОплаты'],
                       'ДатаТретьейОплаты': row['ДатаТретьейОплаты'],
                       'ДатаОкончанияДоговора': row['ДатаОкончанияДоговора'],
                       'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       'ФиоПодписывающего': row['ФиоПодписывающего'],
                       'ДатаПодписиДоговора': row['ДатаПодписиДоговора'], 'ДатаРождения': row['ДатаРождения'],
                       'СерияПаспорта': row['СерияПаспорта'], 'НомерПаспорта': row['НомерПаспорта'],
                       'ДатаВыдачиПаспорта': row['ДатаВыдачиПаспорта'], 'Выдан': row['Выдан'],
                       'АдресРегистрации': row['АдресРегистрации'], 'Снилс': row['Снилс'],
                       'КонтактныйТелефон': row['КонтактныйТелефон']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_contracts}/{row["ФиоСлушателя"]}.docx')




        messagebox.showinfo('Miranda', 'Создание договоров успешно завершено!')
    except NameError as e:
        messagebox.showinfo('Miranda', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


def generate_order_enroll():
    """
    Функция для создания сертификатов

    """
    try:
        reader = csv.DictReader(open(name_file_data_order_enroll), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем словарь, где ключ это наименование  группы а значение это список студентов с подобной группой
        df = pd.read_csv(name_file_data_order_enroll, delimiter=';', encoding='cp1251')
        dct_groups = {group: [] for group in df['Группа'].unique()}
        for row in data:
            dct_groups[row['Группа']].append(row)

        for group in dct_groups.items():
            # Перебираем по группам
        # Создаем в цикле документы
            doc = DocxTemplate(name_file_template_order_enroll)
            fio_lst = []
            for row in group[1]:
                fio_lst.append(row['ФиоСлушателя'])


            context = {'ДатаПодписанияПриказа': group[1][0]['ДатаПодписанияПриказа'],
                       'НомерПриказа': group[1][0]['НомерПриказа'],
                       'НазваниеОрганизации': group[1][0]['НазваниеОрганизации'],
                       'Профессия': group[1][0]['Профессия'], 'Группа': group[1][0]['Группа'],
                       'Программа': group[1][0]['Программа'], 'ДолжностьПодписывающего': group[1][0]['ДолжностьПодписывающего'],
                       'НачалоОбучения': group[1][0]['НачалоОбучения'],
                       'ФиоПодписывающего': group[1][0]['ФиоПодписывающего'],
                       'КонецОбучения': group[1][0]['КонецОбучения'],
                       'Исполнитель': group[1][0]['Исполнитель'],
                       'СрокВЧасах': group[1][0]['СрокВЧасах'], 'lst_students': fio_lst}
                       # }          context = {'ДатаПодписанияПриказа': group[0]['ДатаПодписанияПриказа'],
                       # 'НомерПриказа': row['НомерПриказа'],
                       # 'НазваниеОрганизации': row['НазваниеОрганизации'],
                       # 'Профессия': row['Профессия'], 'Группа': row['Группа'],
                       # 'Программа': row['Программа'], 'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       # 'НачалоОбучения': row['НачалоОбучения'],
                       # 'ФиоПодписывающего': row['ФиоПодписывающего'],
                       # 'КонецОбучения': row['КонецОбучения'],
                       # 'Исполнитель': row['Исполнитель'],
                       # 'СрокВЧасах': row['СрокВЧасах'], 'lst_students': lst_students
                       # }
            doc.render(context)
            # doc.save(f'{path_to_end_folder_order_enroll}/{row["dative_case_lastname"]} {row["dative_case_firstname"]}.docx')
            doc.save(f'{path_to_end_folder_order_enroll}/{group[1][0]["Группа"]}.docx')
        messagebox.showinfo('Dodger', 'Создание приказов успешно завершено!')

    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться приказы')


def select_file_template_certificates():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_certificates
    name_file_template_certificates = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_certificates():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_certificates
    name_file_data_certificates = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_certificates():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_certificates
    path_to_end_folder_certificates = filedialog.askdirectory()


def select_file_template_order_enroll():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_order_enroll
    name_file_template_order_enroll = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_order_enroll():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_order_enroll
    name_file_data_order_enroll = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_order_enroll():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_order_enroll
    path_to_end_folder_order_enroll = filedialog.askdirectory()


# Создаем окно
if __name__ == '__main__':
    window = Tk()
    window.title('Miranda')
    window.geometry('640x480')

    # Создаем ФИО в родительском падеже
    # dct_genitive_fio = create_case(name_file_data_contract)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку свидетельства о повышении
    tab_contract = ttk.Frame(tab_control)
    tab_control.add(tab_contract, text='Создание договоров')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_contract, text='Скрипт для создания договоров')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон

    btn_template_contract = Button(tab_contract, text='1) Выберите шаблон договора', font=('Arial Bold', 20),
                                   command=select_file_template_contracts
                                   )
    btn_template_contract.grid(column=0, row=1, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_contract = Button(tab_contract, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                               command=select_file_data_contracts
                               )
    btn_data_contract.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_contract = Button(tab_contract, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_contracts
                                            )
    btn_choose_end_folder_contract.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов

    btn_create_files_contract = Button(tab_contract, text='4) Создать договора', font=('Arial Bold', 20),
                                       command=generate_contracts
                                       )
    btn_create_files_contract.grid(column=0, row=4, padx=10, pady=10)

    # Создаем вкладку для создания приказов о зачислении
    tab_order_enroll = ttk.Frame(tab_control)
    tab_control.add(tab_order_enroll, text='Создание приказов о зачислении')

    # Добавляем виджеты на вкладку
    lbl_hello = Label(tab_order_enroll, text='Скрипт для создания приказов о зачислении')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон

    btn_template_scc = Button(tab_order_enroll, text='1) Выберите шаблон приказа', font=('Arial Bold', 20),
                              command=select_file_template_order_enroll, )
    btn_template_scc.grid(column=0, row=1, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_scc = Button(tab_order_enroll, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                          command=select_file_data_order_enroll)
    btn_data_scc.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_scc = Button(tab_order_enroll, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_order_enroll)
    btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов

    btn_create_files_scc = Button(tab_order_enroll, text=' Создать приказы', font=('Arial Bold', 20),
                                  command=generate_order_enroll)
    btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

    #
    #
    #
    #
    #
    # # Создаем вкладку Создание удостоверений
    # tab_certificate = ttk.Frame(tab_control)
    # tab_control.add(tab_certificate, text='Создание удостоверений')
    #
    # # Добавляем виджеты на вкладку
    # lbl_hello = Label(tab_certificate, text='Скрипт для создания удостоверений')
    # lbl_hello.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать шаблон
    #
    # btn_template_scc = Button(tab_certificate, text='1) Выберите шаблон удостоверения', font=('Arial Bold', 20),
    #                           command=select_file_template_certificates )
    # btn_template_scc.grid(column=0, row=1, padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_data_scc = Button(tab_certificate, text='2) Выберите файл с данными', font=('Arial Bold', 20),
    #                       command=select_file_data_certificates)
    # btn_data_scc.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_scc = Button(tab_certificate, text='3) Выберите конечную папку', font=('Arial Bold', 20),
    #                                    command=select_end_folder_certificates)
    # btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку для запуска функции генерации файлов
    #
    # btn_create_files_scc = Button(tab_certificate, text=' Создать удостоверения', font=('Arial Bold', 20),
    #                               command=generate_certificates)
    # btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()
