from docx import Document
from docx.shared import Pt


class MainApp():
    def __init__(self):
        # self.get_list()
        self.open_doc()
    
    def get_list(self):
        stud_list = []
        
        stdin = "_"
        
        while stdin != "0":
            stdin = input("Введите список студентов: ")
            
            stud_list.append(stdin)
            
    def open_doc(self, doc_name="./test_data/template.docx"):
        students = [
            {"name":"Иванов Иван Иванович"},
            {"name":"Андреев Андрей Андреевич"}
        ]
        
        target_text = "{{ name_im }}"
        
        for student in students:
            template = Document(doc_name)
        
            print(f"Документ открыт, и используется студент: {student['name']}")
            for paragraph in template.paragraphs:
                if target_text in paragraph.text:
                    paragraph.text = paragraph.text.replace(target_text, " ")
                    new_text = paragraph.add_run(student['name'])
                    print(f"поэтому мы заполняем {student['name']}")
                    
                    new_text.font.name = "Times New Roman"
                    new_text.font.size = Pt(20)
                    new_text.bold = True
                    new_text.underline = True
            
            template.save(f"{student['name'].split()[0]}.docx")
    
if __name__ == "__main__":
    MainApp()