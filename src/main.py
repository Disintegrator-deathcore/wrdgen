from docx import Document
from python_docx_replace import docx_replace, docx_get_keys
import pymorphy3
from openpyxl import load_workbook


class MainApp():
    def __init__(self, doc_name="./test_data/template.docx"):
        self.morph = pymorphy3.MorphAnalyzer()
        self.target_text = {
            "learning_first": "",
            "curs_num": "",
            "group_name": "",
            "name_ro": "",
            "practice_start_fd": "",
            "practice_end_fd": "",
            "learning_sec": "",
            "name_im": "",
            "order_num": "",
            "order_date": "",
            "P_num": "",
            "specialization": "",
            "PM_name": "",
            "practice_start_day": "",
            "practice_start_month": "",
            "practice_start_year": "",
            "practice_end_day": "",
            "practice_end_month": "",
            "practice_start_year": "",
            "hours": "",
            "end_hour": "",
            "gender": ""
        }
        self.doc_name = doc_name
        self.data_update()
        self.replace_text(self.students_update())
            
    def replace_text(self, students):
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
            template.save(f"results/{self.target_text['name_ro'].split()[0]}.docx")
            
            # Закрываем документ (хотя в python-docx явное закрытие не требуется, но всё равно освобождаем ресурс)
            del template
    
    def remorph_word(self, word):
        if len(word.split()) > 1:
            return self.fn_remorph(word)
        else:
            return self.learning_remorph(word)
    
    # Преобразование одного слова
    def learning_remorph(self, word):
        try:
            parse_result = self.morph.parse(word)[0]
            gender_tag = {"femn" if self.target_text["gender"] == "femn" else "masc"}
            new_word = parse_result.inflect(gender_tag).word.capitalize()
            return new_word
        except Exception as e:
            print(f"Ошибка при обработке слова '{word}': {e}")
            return word
    
    # Преобразование ФИО в дательный падеж
    def fn_remorph(self, full_name):
        parts = full_name.split()
        result_parts = []
        for part in parts:
            try:
                # Парсим и превращаем в родительный падеж
                inflected_part = self.morph.parse(part)[0].inflect({'gent'}).word.capitalize()
                result_parts.append(inflected_part)
            except Exception as e:
                print(f"Ошибка при обработке имени '{part}': {e}")
                result_parts.append(part)
        return ' '.join(result_parts)
    
    # Функция обновления данных в словаре из excel
    def data_update(self, data_file="test_data/list_students.xlsx"):
        workbook = load_workbook(filename=data_file, read_only=True)
        sheet = workbook.active
        
        # Забираем конкретные значения из определённых ячеек
        self.target_text.update({
            "learning_first": sheet["B1"].value,
            "curs_num": sheet["C1"].value,
            "group_name": sheet["D1"].value,
            "practice_start_fd": sheet["E1"].value,
            "practice_end_fd": sheet["F1"].value,
            "learning_sec": sheet["G1"].value,
            "order_num": sheet["H1"].value,
            "order_date": sheet["I1"].value,
            "P_num": sheet["J1"].value,
            "specialization": sheet["K1"].value,
            "PM_name": sheet["L1"].value,
            "practice_start_day": sheet["M1"].value,
            "practice_start_month": sheet["N1"].value,
            "practice_start_year": sheet["O1"].value,
            "practice_end_day": sheet["P1"].value,
            "practice_end_month": sheet["Q1"].value,
            "hours": sheet["R1"].value,
            "end_hour": sheet["S1"].value
        })
        
    # Функция обновления списка студентов из excel
    def students_update(self, data_file="test_data/list_students.xlsx"):
        workbook = load_workbook(filename=data_file, read_only=True)
        sheet = workbook.active
        return [row[0].strip() for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True) if row[0]]
    
if __name__ == "__main__":
    MainApp()