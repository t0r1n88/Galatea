from PyQt5 import uic
from PyQt5.QtWidgets import QApplication

def show_message():
    print('Во имя Линди Бут!!!')

Form, Window = uic.loadUiType("base_app.ui")

app = QApplication([])
# Создаем объект окна
window = Window()
# Создаем форму которая будет отображатся в окне
form = Form()
form.setupUi(window)
# Отображаем окно программы
window.show()

# Привязываем функцию к кнопке
form.getFileNamesButton.clicked.connect(show_message)

app.exec()





