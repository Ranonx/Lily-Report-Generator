import docx
from utils_functions.create_title import create_title
from utils_functions.create_table import create_table, add_table_row
from utils_functions.create_image import create_image


def create_word_template_with_image(output_path, img_paths):
    doc = docx.Document()

    # Add table with name, age, and gender
    table = create_table(doc, 1, 3)
    add_table_row(table, ["姓名：", "年龄：", "性别："])
    create_image(doc, img_paths[0], 6)
    doc.add_page_break()

    # Add the title on the second page
    create_title(doc, "实际足部受力成像情况：", 2)
    create_image(doc, img_paths[1], 6)  

    # Add a table to place two images horizontally
    compare_table = create_table(doc, 1, 2)
    add_image_with_title(doc, img_paths[2], "正常足底受力", compare_table.cell(0, 0))
    add_image_with_title(doc, img_paths[3], "实际足底受力", compare_table.cell(0, 1))

    doc.add_page_break() 

    create_title(doc, "足跟内外翻情况对比：", 2)
    compare_table2 = create_table(doc, 1, 2)
    add_image_with_title(doc, img_paths[4], "正常后跟内外翻情况 ", compare_table2.cell(0, 0))
    add_image_with_title(doc, img_paths[5], "实际后跟内外翻情况", compare_table2.cell(0, 1))

    doc.save(output_path)


def add_image_with_title(doc, img_path, title_text, cell):
    cell.width = docx.shared.Inches(3)
    cell_img = cell.add_paragraph().add_run()
    create_image(cell_img, img_path, 3)

    title = cell.add_paragraph()
    title.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(title_text)
