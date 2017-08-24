from docxtpl import DocxTemplate
from docx.oxml.shared import qn
import docx
import sys
import os.path
import json
import glob

bookmark_name_starts_with = "lab_work_place"

if __name__ == '__main__':
    if len(sys.argv) != 4:
        print("Количество аргументов должно быть равно3: файл шаблона, файл описания курса, имя выходного файла")
        sys.exit(1)

    templ_file_name = sys.argv[1]
    descr_file_name = sys.argv[2]
    out_file_name = sys.argv[3]

    if not os.path.exists(templ_file_name):
        print("Файл шаблона {fname} не найден".format(fname=templ_file_name))
        sys.exit(1)

    if not os.path.exists(descr_file_name):
        print("Файл описания курса {fname} не найден".format(fname=descr_file_name))
        sys.exit(1)

    if os.path.exists(out_file_name):
        print("Файл РП {fname} уже существует. Удалите его".format(fname=out_file_name))
        sys.exit(1)

    try:
        course_file = open(descr_file_name, encoding="utf-8")
        json_string = course_file.read()
        course_description = json.loads(json_string)
        course_file.close()
    except (IOError, OSError, ValueError, EOFError):
        print("Файл описания {fname} содержит ошибки и не соответствует стандарту JSON. Проверьте файл.".format(fname=descr_file_name))
        sys.exit(1)

    doc = DocxTemplate(templ_file_name)
    doc.render(course_description)
    doc.save(out_file_name)
