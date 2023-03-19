import docx
from utils_functions.create_title import create_title
from utils_functions.create_table import create_table, add_table_row
from utils_functions.create_image import create_image
from doc_template import add_image_with_title

def add_formatted_content(doc, filename):
    with open(filename, "r", encoding="utf-8") as f:
        content_lines = f.readlines()

    for line in content_lines:
        line = line.strip()
        if line.startswith("-"):
            paragraph = doc.add_paragraph(line[1:].strip(), style="ListBullet")
        else:
            create_title(doc, line, 3)

def create_word_template_with_image(output_path, img_paths):
    doc = docx.Document()

    create_title(doc, "足踝检测评估报告", 1)
    table = create_table(doc, 0, 3)
    add_table_row(table, ["姓名：\t", "年龄：\t", "性别："])

    create_image(doc, img_paths[0], 6)
    doc.add_page_break()

    create_title(doc, "实际足部受力成像情况：", 2)
    create_image(doc, img_paths[1], 6)  

    compare_table = create_table(doc, 1, 2)
    add_image_with_title(doc, img_paths[2], "正常足底受力", compare_table.cell(0, 0))
    add_image_with_title(doc, img_paths[3], "实际足底受力", compare_table.cell(0, 1))

    doc.add_page_break() 

    create_title(doc, "足跟内外翻情况对比：", 2)
    compare_table2 = create_table(doc, 1, 2)
    add_image_with_title(doc, img_paths[4], "正常后跟内外翻情况", compare_table2.cell(0, 0))
    add_image_with_title(doc, img_paths[5], "实际后跟内外翻情况", compare_table2.cell(0, 1))

    create_title(doc, "初步诊断：", 2)
    with open("diagnosis_text.txt", "r", encoding="utf-8") as f:
        diagnosis_text = f.read()
    doc.add_paragraph(diagnosis_text)
    create_image(doc, img_paths[6], 6)

    doc.add_page_break() 

    create_title(doc, "足弓发育四个阶段", 2)
    create_image(doc, img_paths[7], 6)
    add_formatted_content(doc, "content_text.txt")

    doc.save(output_path)
