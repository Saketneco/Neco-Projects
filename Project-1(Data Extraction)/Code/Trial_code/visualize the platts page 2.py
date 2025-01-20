import pdfplumber
#from pdf2image import convert_from_path
from PIL import Image, ImageDraw
import matplotlib.pyplot as plt

def extract_table_lines_and_display(pdf_path, page_number):
    # Convert PDF to image
    #images = convert_from_path(pdf_path, first_page=page_number, last_page=page_number)
    #image = images[0]  # Get the first page image (since we specified a single page)

    # Use pdfplumber to extract table information from the same page
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]  # Page number is 1-based in extract_table_platts_pdf, but 0-based here
        tables = page.extract_tables(table_settings = {
                "horizontal_strategy": "lines_strict",
                "explicit_vertical_lines": [310,425,504,530,550],
                "explicit_horizontal_lines": [80],
                "snap_tolerance": 8,
                "join_tolerance": 1,
                "edge_min_length": 2,
                "min_words_vertical": 2,
                "min_words_horizontal": 2,
                "intersection_tolerance": 2,
                "text_tolerance": 2,
            })

        # If tables are found, extract the boundaries
        if tables:
            print(f"Tables found on page {page_number}:")
            for table in tables:
                print(table)  # You can visualize or analyze this table as needed
            print(f"Number of tables found: {len(tables)}")
        else:
            print(f"No tables found on page {page_number}.")
        
        # Extracting table boundaries (lines)
        table_lines = page.extract_table_settings().get("explicit_vertical_lines", []) + page.extract_table_settings().get("explicit_horizontal_lines", [])
        print("Table lines:", table_lines)

    # Drawing lines on the image
    # draw = ImageDraw.Draw(image)
    # for line in table_lines:
    #     # Here we draw the lines on the image
    #     # This is just a placeholder, you would need to calculate where the lines should be based on the line's coordinates
    #     draw.line((line[0], 0, line[0], image.height), fill="red", width=2)  # Vertical lines in red
    #     draw.line((0, line[1], image.width, line[1]), fill="blue", width=2)  # Horizontal lines in blue

    # Display the image with the table boundaries drawn
    # plt.figure(figsize=(10, 10))
    # plt.imshow(image)
    # plt.axis("off")
    # plt.show()

# Example usage
pdf_path = "g:\My Drive\PDF_files\ICT_20241206.pdf"
page_number = 2  # Example page number to extract tables and display lines
extract_table_lines_and_display(pdf_path, page_number)
