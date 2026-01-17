from docx import Document
from python_docx_replace import docx_replace, docx_get_keys
import pymorphy3


class MainApp():
    def __init__(self, doc_name="./test_data/template.docx"):
        # Вводные данные
        self.morph = pymorphy3.MorphAnalyzer()
        self.target_text = [
            "{{ learning_first }}",
            "{{ curs_num }}",
            "{{ group_name }}",
            "{{ name_ro }}",
            "{{ practice_start_fd }} "
            "{{ practice_end_fd }}",
            "{{ learning_sec }}",
            "{{ name_im }}",
            "{{ order_num }}",
            "{{ order_date }}",
            "{{ P_num }}",
            "{{ specialization }}",
            "{{ PM_name }}",
            "{{ practice_start_day }}",
            "{{ practice_start_month }}",
            "{{ practice_start_year}}",
            "{{ practice_end_day }}",
            "{{ practice_end_month }}",
            "{{ practice_start_year }}",
            "{{ hours }}",
            "{{ end_hour }}"
        ]
        self.doc_name = doc_name
        
        # Вызовы функций
        # self.get_list()
        self.replace_text()
    
    # Ввод пользователем списка студентов
    def get_list(self):
        stud_list = []
        
        stdin = "_"
        
        while stdin != "0":
            stdin = input("Введите список студентов: ")
            
            stud_list.append(stdin)
            
    # Функция заены текста в документе
    def replace_text(self):
        ########
        students = [
            {"name":"Иванов Иван Иванович"},
            {"name":"Андреев Андрей Андреевич"}
        ]
        
        data_dict = {
            "learning_first": "Обучающегося",
            "curs_num": "2",
            "group_name": "170 ис",
            "name_im": "Иванов Иван Иванович",
            "name_ro": "Иванова Ивана Ивановича"
        }
        ########
        
        
        # Заменяем текст в документе для каждого студента
        for student in students:
            data_dict["name_im"] = student["name"]
            data_dict["name_ro"] = self.remorph_word(student["name"])
            
            
            template = Document(self.doc_name)
            
            docx_replace(template, **data_dict)
            
            template.save(f"{student['name'].split()[0]}.docx")
            
    def remorph_word(self, word):
        full_name = ""
            
        for p_name in word.split():
            ch_word = self.morph.parse(p_name)[0].inflect({"gent"})[0]
            ch_word = ch_word[0].upper() + ch_word[1:]
            full_name += f"{ch_word} "
        
        full_name = full_name.strip()
        
        return full_name
    
if __name__ == "__main__":
    MainApp()