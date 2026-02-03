import threading
import toga
import asyncio
import os

from main import MainApp


class MyApp(toga.App):
    def startup(self):
        self.template = ""
        self.data = ""
        
        # Создание основного окна
        self.main_window = toga.MainWindow(resizable=False)
        self.main_window.size = (550, 300)
        self.main_window.content = toga.Box()
        
        # Создание областей для надписей и кнопок
        self.labels_box = toga.Box(direction="column", align_items="start", width=300)
        self.buttons_box = toga.Box(direction="column", align_items="start")
        
        # Создание кнопки выбора шаблона
        self.select_template_btn = toga.Button("Выбрать шаблон", on_press=self.select_template)
        self.select_template_btn.style.margin = 10
        self.select_template_btn.style.flex = 1
        self.select_template_btn.font_family = "Times New Roman"
        self.select_template_btn.font_size = 14
        
        # Создание текстового поля с файлом с шаблоном
        self.template_file_txtInp = toga.TextInput()
        self.template_file_txtInp.margin = 10
        self.template_file_txtInp.font_family = "Times New Roman"
        self.template_file_txtInp.font_size = 14
        self.template_file_txtInp.value = "Файл пока не выбран"
        self.template_file_txtInp.readonly = True
        
        # Создание кнопки выбора файла с данными
        self.select_data_btn = toga.Button("Выбрать файл с данными", on_press=self.select_data)
        self.select_data_btn.style.margin = 10
        self.select_data_btn.style.flex = 1
        self.select_data_btn.font_family = "Times New Roman"
        self.select_data_btn.font_size = 14
        
        # Создание текстового поля с файлом с данными
        self.data_file_txtInp = toga.TextInput()
        self.data_file_txtInp.margin = 10
        self.data_file_txtInp.font_family = "Times New Roman"
        self.data_file_txtInp.font_size = 14
        self.data_file_txtInp.value = "Файл пока не выбран"
        self.data_file_txtInp.readonly = True
        
        # Создание кнопки запуска скрипта
        self.start_btn = toga.Button("Запустить скрипт", on_press=self.start_script)
        self.start_btn.style.margin = (200, 100, -200, -450)
        self.start_btn.style.flex = 1
        self.start_btn.style.font_family = "Times New Roman"
        self.start_btn.style.font_size = 16
        
        # Создание индикатора загрузки
        self.load_indc = toga.ActivityIndicator()
        self.load_indc.style.margin = (135, 200, -200, -275)
        
        # Помещение всех элементов в нудные области
        self.buttons_box.add(
            self.select_template_btn,
            self.select_data_btn
        )
        
        self.labels_box.add(
            self.template_file_txtInp,
            self.data_file_txtInp
        )
        
        self.main_window.content.add(
            self.labels_box,
            self.buttons_box,
            self.start_btn,
            self.load_indc
        )
        
        self.main_window.show()
    
    # Обработка запуска скрипта
    def start_script(self, widget):
        # Обработка события выбора файлов
        if self.template != "":
            if self.data != "":
                self.load_indc.start()
                thread = threading.Thread(target=self.start_gen)
                thread.start()
            else:
                self.info()
        else:
            self.info()
    
    # Запуск скрипта
    def start_gen(self):
        MainApp(self.template, self.data)
        self.load_indc.stop()
    
    # Функция выбора шаблона
    def select_template(self, widget):
        cur_template = toga.OpenFileDialog("Выберите файл", str(os.getcwd()), ["docx", "doc"], False)
        
        task = asyncio.create_task(self.main_window.dialog(cur_template))
        task.add_done_callback(self.dialog_dismissed_template)
    
    # Ответная функция выбора шаблона
    def dialog_dismissed_template(self, task):
        self.template_file_txtInp.value = task.result()
        self.template = task.result()
    
    # Функция выбора файла с данными
    def select_data(self, widget):
        cur_data = toga.OpenFileDialog("Выберите файл", str(os.getcwd()), ["xlsx", "xls"], False)
        
        task = asyncio.create_task(self.main_window.dialog(cur_data))
        task.add_done_callback(self.dialog_dismissed_data)
    
    # Ответная функция выбора файла с данными
    def dialog_dismissed_data(self, task):
        self.data_file_txtInp.value = task.result()
        self.data = task.result()
    
    # Отображение информации пользователю о выборе файла
    def info(self, **kwargs):
        my_info = toga.InfoDialog("Забыли выбрать", "Вы забыли выбрать файл")
        task = asyncio.create_task(self.main_window.dialog(my_info))


if __name__ == '__main__':
    app = MyApp("WrdGen", "localhost")
    app.main_loop()
