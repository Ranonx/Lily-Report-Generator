import os
import fitz  # PyMuPDF

def get_pdf_file(pdf_folder):
    for file in os.listdir(pdf_folder):
        if file.endswith('.pdf'):
            return os.path.join(pdf_folder, file)
    return None

def extract_page_as_image(pdf_path, page_number, zoom=2, crop=None, output_path=None):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)
    if crop:
        page_rect = page.rect
        crop_rect = fitz.Rect(page_rect.x0, crop[0]*page_rect.y1, page_rect.x1, crop[1]*page_rect.y1)
        page.set_cropbox(crop_rect)
    mat = fitz.Matrix(zoom, zoom)  # Create a transformation matrix with the zoom factor
    pix = page.get_pixmap(alpha=False, matrix=mat)  # Apply the transformation matrix to the pixmap
    output_image = output_path or 'page_{}_crop.png'.format(page_number)
    pix.save(output_image)
    return output_image

def extract_page_4d(pdf_path, page_number, zoom=2, crop=(0, 1, 0, 1), output_path=None):
    doc = fitz.open(pdf_path)
    page = doc[page_number - 1]

    page_width, page_height = float(page.rect.width), float(page.rect.height)
    left, top, right, bottom = page_width * crop[2], page_height * crop[0], page_width * crop[3], page_height * crop[1]
    rect = fitz.Rect(left, top, right, bottom)

    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect)
    output_image = output_path or 'page_{}_4d_crop.png'.format(page_number)
    pix.save(output_image)

    return output_image
