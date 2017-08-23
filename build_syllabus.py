from docxtpl import DocxTemplate
import sys
import os.path

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Количество аргументов должно быть равно3: файл шаблона, файл описания курса, имя выходного файла")
        sys.exit(1)

    templ_file_name = sys.argv[0]
    descr_file_name = sys.argv[1]
    out_file_name = sys.argv[2]

    if not os.path.exists(templ_file_name):
        print("Файл шаблона {fname} не найден".format(fname=templ_file_name))
        sys.exit(1)

    if os.path.exists(descr_file_name):
        print("Файл описания курса {fname} не найден".format(fname=descr_file_name))
        sys.exit(1)

    if not os.path.exists(templ_file_name):
        print("Файл РП {fname} уже существует. Удалите его".format(fname=out_file_name))
        sys.exit(1)

    doc = DocxTemplate("my_word_template.docx")
    context = {'company_name': "World company"}
    doc.render(context)
    doc.save("generated_doc.docx")