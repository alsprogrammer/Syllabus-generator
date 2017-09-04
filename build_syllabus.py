from docxtpl import DocxTemplate
import math
import sys
import os.path
import json
from xml.sax.saxutils import escape

bookmark_name_starts_with = "lab_work_place"

study_form = {"очная": "of", "заочная": "zf", "очно-заочная": "ozf"}

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

    for element in course_description["feature"]["elements"]:
        if "label" in element.keys():
            course_description[element["label"]+"_units"] = sum(element["units"])
            course_description[element["label"]+"_hours"] = sum(element["hours"])

    control_questions = []
    question_file = open(question_file_name, encoding="utf-8")
    for cur_str in question_file:
        if cur_str.startswith(("\n", "\r")):
            pass
        else:
            control_questions.append(escape(cur_str).replace("\n", ""))

    course_description.update({"col_labels": ["Всего, зачетных единиц (акад.часов)"]})
    course_description["col_labels"] = course_description["col_labels"] + course_description["feature"]["col_labels"]
    course_description.update({"features": []})
    for cur_feature in course_description["feature"]["elements"]:
        units_sum = sum(cur_feature["units"])
        hours_sum = sum(cur_feature["hours"])
        cols = ["{0}({1})".format(units_sum, hours_sum) if units_sum !=0 else " "]
        for i in range(len(cur_feature["units"])):
            cols.append("{0}({1})".format(cur_feature["units"][i], cur_feature["hours"][i]) if cur_feature["units"][i] != 0 else "")
        course_description["features"].append({"label": cur_feature["name"], "cols": cols})
    course_description["features"].append({"label": "Вид промежуточной аттестации (зачет, экзамен)", "cols": [course_description["control_type"]] + course_description["feature"]["control_type"]})

    course_description.update({"test": test})
    course_description.update({"control_questions": control_questions})

    whole_parts_num = sum(map(lambda x: len(x["module_parts"]), course_description["course_content"]))

    whole_labs_num = sum(map(lambda x: sum(map(lambda y: len(y["lab_works"]), x["module_parts"])), course_description["course_content"]))
    whole_pracs_num = sum(map(lambda x: sum(map(lambda y: len(y["practice_works"]), x["module_parts"])), course_description["course_content"]))

    lessons_per_module = math.floor(course_description["lecture_hours"] / whole_parts_num)
    seminars_per_module = math.floor(course_description["seminar_hours"] / whole_parts_num)
    labs_per_module = math.floor(course_description["lab_hours"] / whole_parts_num)
    pracs_per_module = math.floor(course_description["practice_hours"] / whole_parts_num)
    selftraining_per_module = math.floor(course_description["selftraining_hours"] / whole_parts_num)

    labs_per_part = math.floor(course_description["lab_hours"] / whole_labs_num) if whole_labs_num != 0 else 0
    pracs_per_part = math.floor(course_description["practice_hours"] / whole_pracs_num) if whole_pracs_num != 0 else 0

    for course_module in course_description["course_content"]:
        for module_part in course_module["module_parts"]:
            module_part.update({"module_part_lessons": lessons_per_module})
            module_part.update({"module_part_seminars": seminars_per_module})
            module_part.update({"module_part_lab": labs_per_module})
            module_part.update({"module_part_selftraining": selftraining_per_module})
            module_part["selftraining_hours_f"].update({[study_form[course_description["study_form"]]][0]: selftraining_per_module})

    course_description["course_content"][-1]["module_parts"][-1].update({"module_part_lessons": course_description["lecture_hours"] - lessons_per_module * whole_parts_num + lessons_per_module})
    course_description["course_content"][-1]["module_parts"][-1].update({"module_part_seminars": course_description["seminar_hours"] - seminars_per_module * whole_parts_num + seminars_per_module})
    course_description["course_content"][-1]["module_parts"][-1].update({"module_part_lab": course_description["lab_hours"] - labs_per_module * whole_parts_num + labs_per_module})
    course_description["course_content"][-1]["module_parts"][-1].update({"module_part_selftraining": course_description["selftraining_hours"] - selftraining_per_module * whole_parts_num + selftraining_per_module})
    course_description["course_content"][-1]["module_parts"][-1]["selftraining_hours_f"].update({[study_form[course_description["study_form"]]][0]: course_description["selftraining_hours"] - selftraining_per_module * whole_parts_num + selftraining_per_module})

    for course_module in course_description["course_content"]:
        for module_part in course_module["module_parts"]:
            for lab in module_part["lab_works"]:
                lab.update({study_form[course_description["study_form"]]: labs_per_part})
    if whole_labs_num != 0:
        course_description["course_content"][-1]["module_parts"][-1]["lab_works"][-1].update({study_form[course_description["study_form"]]: labs_per_part + course_description["lab_hours"] - labs_per_part * whole_labs_num})

    for course_module in course_description["course_content"]:
        for module_part in course_module["module_parts"]:
            for lab in module_part["practice_works"]:
                lab.update({study_form[course_description["study_form"]]: pracs_per_part})
    if whole_pracs_num != 0:
        course_description["course_content"][-1]["module_parts"][-1]["practice_works"][-1].update({study_form[course_description["study_form"]]: pracs_per_part + course_description["practice_hours"] - pracs_per_part * whole_pracs_num})

    doc = DocxTemplate(templ_file_name)
    doc.render(course_description)
    doc.save(out_file_name)
