# Import necessary libraries
import pytesseract    # for performing OCR on the image
from PIL import Image    # for opening and manipulating images
from openpyxl import Workbook    # for creating and saving data in Excel file
from pathlib import Path    # for finding image files in a directory

# Set the path for Tesseract OCR executable file
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Define a function to extract text from image using OCR


def extract_text_from_image(image_path):
    img = Image.open(image_path)    # open the image using PIL
    # extract text from image using OCR
    text = pytesseract.image_to_string(img, lang='eng')
    return text.splitlines()    # return extracted text as a list of lines

# Define a function to save the extracted text to an Excel file


def save_text_to_excel(text_lines, output_file):
    wb = Workbook()    # create a new Excel workbook
    ws = wb.active    # select the active worksheet

    # Write each line of text to a new row in the first column of the worksheet
    for i, line in enumerate(text_lines, start=1):
        ws.cell(row=i, column=1, value=line)

    wb.save(output_file)    # save the workbook to the specified output file

# Define the main function to process all images in a specified folder


def main():
    # set the path for input folder containing images
    input_folder = "C:/Users/User/Desktop/img"
    output_file = "output.xlsx"    # set the name for output Excel file

    # set the list of supported image file extensions
    supported_extensions = [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif"]
    image_files = [file for ext in supported_extensions for file in Path(
        input_folder).glob(f"*{ext}")]    # find all image files with supported extensions in the input folder

    if not image_files:    # if no image files are found, print a message and exit
        print("No supported images found in the specified folder.")
        return

    all_text_lines = []    # create an empty list to store all extracted text from images

    # Process each image file, extract text and add it to the list
    for image_file in image_files:
        print(f"Processing file: {image_file}")
        text_lines = extract_text_from_image(str(image_file))
        all_text_lines.extend(text_lines)

    # save all extracted text to Excel file
    save_text_to_excel(all_text_lines, output_file)
    print("Text successfully saved to Excel file.")


# Call the main function if this script is executed directly
if __name__ == "__main__":
    main()
