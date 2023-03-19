import sys
import os
from pdf_utils import get_pdf_file, extract_page_as_image, extract_page_4d
from docx_utils import create_word_template_with_image

def main(pdf_folder, output_path):
    pdf_file = get_pdf_file(pdf_folder)
    if not pdf_file:
        print('No PDF file found in the folder.')
        sys.exit(1)

    img_folder = "img"
    if not os.path.exists(img_folder):
        os.makedirs(img_folder)

    img1 = extract_page_as_image(pdf_file, 8, crop=(0.1, 0.90), output_path=os.path.join(img_folder, "img1.png"))
    img2 = extract_page_as_image(pdf_file, 1, crop=(0.27, 0.69), output_path=os.path.join(img_folder, "img2.png"))
    img3 = os.path.join(img_folder, "good.png")
    img4 = extract_page_4d(pdf_file, 1, crop=(0.27, 0.50, 0.35, 0.67), output_path=os.path.join(img_folder, "img4.png"))
    img5 = os.path.join(img_folder, "good2.png")
    img6 = extract_page_4d(pdf_file, 1, crop=(0.27, 0.50, 0.65, 0.99), output_path=os.path.join(img_folder, "img6.png"))
    img7 = os.path.join(img_folder, "table.png")

    create_word_template_with_image(output_path, [img1, img2, img3, img4, img5, img6,img7])
    print('Word template created successfully.')

if __name__ == '__main__':
    pdf_folder = 'pdf'
    output_path = 'output.docx'
    main(pdf_folder, output_path)
