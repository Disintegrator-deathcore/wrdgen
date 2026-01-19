from docx import Document
from python_docx_replace import docx_replace, docx_get_keys
import pymorphy3
from openpyxl import load_workbook


class MainApp():
    def __init__(self, doc_name="./test_data/template.docx"):
        # Вводные данные
        self.morph = pymorphy3.MorphAnalyzer()
        self.target_text = {
            "learning_first":"",
            "curs_num":"",
            "group_name":"",
            "name_ro":"",
            "practice_start_fd":"",
            "practice_end_fd":"",
            "learning_sec":"",
            "name_im":"",
            "order_num":"",
            "order_date":"",
            "P_num":"",
            "specialization":"",
            "PM_name":"",
            "practice_start_day":"",
            "practice_start_month":"",
            "practice_start_year":"",
            "practice_end_day":"",
            "practice_end_month":"",
            "practice_start_year":"",
            "hours":"",
            "end_hour":"",
            "gender":""
        }
        # self.students_update()
        self.doc_name = doc_name
        
        # Вызовы функций
        self.data_update()
        self.replace_text(self.students_update())
            
    # Функция заены текста в документе
    def replace_text(self, students):
        # ########
        
        # self.target_text = {
        #     "learning_first": "Обучающегося",
        #     "curs_num": "2",
        #     "group_name": "170 ис",
        #     "learning_sec": "Обучающийся"
        # }
        # ########

        # Заменяем текст в документе для каждого студента
        for student in students:
            self.target_text["name_im"] = student
            self.target_text["name_ro"] = self.remorph_word(student)
            self.target_text["learning_first"] = self.remorph_word(self.target_text["learning_first"])
            self.target_text["learning_sec"] = self.remorph_word(self.target_text["learning_sec"])
            
            # Открываем шаблон документа
            template = Document(self.doc_name)
            
            # Заменяем текст в шаблоне
            docx_replace(template, **self.target_text)
            
            # Сохраняем изменённый документ с новым названием
            template.save(f"results/{self.target_text["name_ro"].split()[0]}.docx")
            
    def remorph_word(self, word):
        if len(word.split()) > 1:
            transformed_words = []
            for part in word.split():
                try:    
                    if part != self.morph.parse(part)[0].normal_form:
                        pass
                
                    # Получаем первую интерпретацию слова и преобразуем её в родительный падеж
                    parsed_word = self.morph.parse(part)[0].inflect({'gent'})
                    
                    # Преобразуем слово в строчную форму и делаем первую букву заглавной
                    if parsed_word:
                        inflected_word = parsed_word.word.capitalize()
                        transformed_words.append(inflected_word)
                        self.target_text["gender"] = self.morph.parse(part)[0].tag.gender
                        
                    elif parsed_word == None:
                        transformed_words.append(part)
                        

                except Exception as e:
                    print(f"Произошла ошибка при обработке слова '{part}': {e}")
            return ' '.join(transformed_words)
        
        else:
            try:
                if self.target_text["gender"] == "femn":
                    word = self.morph.parse(word)[0].inflect({"femn"}).word.capitalize()
                else:
                    word = self.morph.parse(word)[0].inflect({"masc"}).word.capitalize()
                
            except Exception as e:
                print(f"Произошла ошибка при обработке слова '{word}: {e}")
                
            return word
    
    def data_update(self, data_file="test_data/list_students.xlsx"):
        workbok = load_workbook(filename=data_file)
        
        sheet = workbok.active
        
        self.target_text["learning_first"] = sheet["B1"].value
        self.target_text["curs_num"] = sheet["C1"].value
        self.target_text["group_name"] = sheet["D1"].value
        self.target_text["practice_start_fd"] = sheet["E1"].value
        self.target_text["practice_end_fd"] = sheet["F1"].value
        self.target_text["learning_sec"] = sheet["G1"].value
        self.target_text["order_num"] = sheet["H1"].value
        self.target_text["order_date"] = sheet["I1"].value
        self.target_text["P_num"] = sheet["J1"].value
        self.target_text["specialization"] = sheet["K1"].value
        self.target_text["PM_name"] = sheet["L1"].value
        self.target_text["practice_start_day"] = sheet["M1"].value
        self.target_text["practice_start_month"] = sheet["N1"].value
        self.target_text["practice_start_year"] = sheet["O1"].value
        self.target_text["practice_end_day"] = sheet["P1"].value
        self.target_text["practice_end_month"] = sheet["Q1"].value
        self.target_text["hours"] = sheet["R1"].value
        self.target_text["end_hour"] = sheet["S1"].value
        
    
    def students_update(self, data_file="test_data/list_students.xlsx"):
        # Загружаем рабочую книгу
        workbook = load_workbook(filename=data_file)
        
        # Выбираем активный лист (обычно первый лист файла)
        sheet = workbook.active
        
        student_list = []
        
        # Проходим по первому столбцу (столбец 'A') начиная с первой строки
        for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
            cell_value = row[0]
            
            # Проверяем наличие значения в ячейке
            if cell_value is not None:
                student_list.append(cell_value.strip())
        
        return student_list
    
if __name__ == "__main__":
    MainApp()