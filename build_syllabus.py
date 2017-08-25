from docxtpl import DocxTemplate
import sys
import os.path
import json
from xml.sax.saxutils import escape


bookmark_name_starts_with = "lab_work_place"

if __name__ == '__main__':
    if len(sys.argv) != 6:
        print("Количество аргументов должно быть равно 5: файл шаблона, файл описания курса, файл с вопросами теста, файл с вопросами экзамена или зачета, имя выходного файла")
        sys.exit(1)

    templ_file_name = sys.argv[1]
    descr_file_name = sys.argv[2]
    test_file_name = sys.argv[3]
    question_file_name = sys.argv[4]
    out_file_name = sys.argv[5]

    if not os.path.exists(templ_file_name):
        print("Файл шаблона {fname} не найден".format(fname=templ_file_name))
        sys.exit(1)

    if not os.path.exists(descr_file_name):
        print("Файл описания курса {fname} не найден".format(fname=descr_file_name))
        sys.exit(1)

    if not os.path.exists(test_file_name):
        print("Файл теста {fname} не найден".format(fname=test_file_name))
        sys.exit(1)

    if not os.path.exists(question_file_name):
        print("Файл теста {fname} не найден".format(fname=question_file_name))
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
        print("Файл описания {fname} содержит ошибки и/или не соответствует стандарту JSON. Проверьте файл.".format(fname=descr_file_name))
        sys.exit(1)

    test = []
    test_file = open(test_file_name, encoding="utf-8")
    task_text = ""
    answers = []
    for cur_str in test_file:
        if cur_str.startswith(("\n", "\r")):
            pass
        elif cur_str.startswith("+") or cur_str.startswith("-"):
            answers.append(escape(cur_str[0:cur_str.rindex("<<")]).replace("\n", ""))
        else:
            if task_text != "":
                test.append({"task": task_text, "answers": answers})
                task_text = ""
                answers = []
            else:
                task_text = escape(cur_str[0:cur_str.rindex("<<")]).replace("\n", "")

    control_questions = []
    question_file = open(question_file_name, encoding="utf-8")
    for cur_str in question_file:
        if cur_str.startswith(("\n", "\r")):
            pass
        else:
            control_questions.append(escape(cur_str).replace("\n", ""))

    course_description.update({"test": test})
    course_description.update({"control_questions": control_questions})

    doc = DocxTemplate(templ_file_name)
    doc.render(course_description)
    doc.save(out_file_name)
