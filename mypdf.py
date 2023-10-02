import PyPDF2
import os

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import load_workbook

# pdfmetrics.registerFont(TTFont('Hebrew', 'Assistant-VariableFont_wght.ttf'))
pdfmetrics.registerFont(TTFont('Hebrew', 'arial-hebrew-bold.ttf'))
pdfmetrics.registerFont(TTFont('Hebrew_david', 'DavidLibre-Bold.ttf'))


def reverse_slicing(s):
    if isinstance(s, str):
        return s[::-1]
    return str(s)  # Convert non-string values to strings


# Define a mapping of parameter names/indices to x, y coordinates
parameter_coordinates = {

    "Parameter1": (0, 510),  # שם מלא
    "Parameter2": (-100, -100),  # רחוב+עיר
    "Parameter3": (-100, -100),  # תז
    "Parameter4": (-100, -100),  # מספר
    "Parameter5": (-100, -100),  # כתובת מלאה + עיר
    "Parameter6": (0, 466),  # קלפיות החלפה

}


def generate_individual_pdf(row_data, output_dir):
    pdf_filename = f"{output_dir}/כתב מינוי_{row_data[3]}.pdf"
    template_canvas = canvas.Canvas(pdf_filename, pagesize=(2000, 2000))
    # template_canvas.setFont('Hebrew', 10)
    # first_row_flag = 0
    # second_row_flag = 0
    # third_row_flag = 0
    # print("---------------------------")
    # print(row_data[1])

    for parameter, value in zip(row_data[3:], parameter_coordinates.keys()):
        if parameter:
            x_coord, y_coord = parameter_coordinates[value]
            template_canvas.setFont('Hebrew_david', 16)
            if value in ["Parameter1"]:
                # template_canvas.setFont('Hebrew_david', 16)
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
            if value in ["Parameter2"]:
                # template_canvas.setFont('Hebrew', 20)
                # text_width = template_canvas.stringWidth(parameter, 'Hebrew', 20)
                # new_x_coord = CENTER_X - (text_width / 2)
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
            if value in ["Parameter3"]:
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
                print(x_coord, y_coord)
            if value in ["Parameter4"]:
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
                print(x_coord, y_coord)
            if value in ["Parameter5"]:
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
                print(x_coord, y_coord)
            if value in ["Parameter6"]:
                template_canvas.drawString(x_coord, y_coord, reverse_slicing(parameter))
                print(x_coord, y_coord)

    template_canvas.save()
    return [row_data[1], row_data[2], pdf_filename]
    # return pdf_filename


# Load Excel data and generate individual PDFs
excel_file = 'input_data_new.xlsx'
workbook = load_workbook(excel_file)
sheet = workbook.active
# Initialize empty arrays for each category
# []
ramat_hasharon = {
    "av_bait_vesadran": "רמש כתב מינוי לאב בית וסדרן.pdf",
    "sadran": "רמש כתב מינוי לסדרן.pdf",
    "sadran_hacvana": "רמש כתב מינוי לסדרן הכוונה.pdf",
    "mazkir": "רמש כתב מינוי מזכיר.pdf",
    "mazkir_mahlif": "רמש כתב מינוי מזכיר מחליף.pdf",
}

bat_yam = []
rishon = []
kfar_saba = []
output_directory = 'miniu'

if not os.path.exists(output_directory):
    os.makedirs(output_directory)
# Define the initial template file
initial_template_file_path = 'pdf_files/minui_template.pdf'
for row in sheet.iter_rows(min_row=2, max_col=8, values_only=True):
    [city, job_description, pdf_filename] = generate_individual_pdf(row, output_directory)

    try:
        # Define the template file based on the city and job description
        if city == 'רמת השרון':
            if job_description == 'מזכיר':
                template_file_path = f'pdf_files/{ramat_hasharon["mazkir"]}'
            elif job_description == 'מזכיר מחליף':
                template_file_path = f'pdf_files/{ramat_hasharon["mazkir_mahlif"]}'
            elif job_description == 'סדרן':
                template_file_path = f'pdf_files/{ramat_hasharon["sadran"]}'
            elif job_description == 'סדרן הכוונה':
                template_file_path = f'pdf_files/{ramat_hasharon["sadran_hacvana"]}'
            elif job_description == 'אב בית וסדרן':
                template_file_path = f'pdf_files/{ramat_hasharon["av_bait_vesadran"]}'
            else:
                template_file_path = initial_template_file_path  # Use the initial template file as a default
        else:
            template_file_path = initial_template_file_path  # Use the initial template file as a default

        # Open the PDF files in binary mode
        template_file = open(template_file_path, 'rb')
        data_file = open(pdf_filename, 'rb')

        # Create PDF reader objects for the files
        template_reader = PyPDF2.PdfReader(template_file)
        data_reader = PyPDF2.PdfReader(data_file)

        # Create a PDF writer object to write the merged output
        merged_writer = PyPDF2.PdfWriter()

        # Get the first page of the template PDF
        template_page = template_reader.pages[0]

        # Get the first page of the data PDF
        data_page = data_reader.pages[0]

        # Merge the data page onto the template page
        template_page.merge_page(data_page)

        # Add the merged page to the output PDF
        merged_writer.add_page(template_page)

        # Save the output PDF file
        final_output_filename = f"{output_directory}/כתב מינוי_{row[3]}.pdf"
        with open(final_output_filename, 'wb') as output_file:
            merged_writer.write(output_file)

    except Exception as e:
        print(f"Error processing row: {str(e)}")

    finally:
        # Close the input and output files if they were successfully opened
        if 'template_file' in locals() and template_file is not None:
            template_file.close()
        if 'data_file' in locals() and data_file is not None:
            data_file.close()
