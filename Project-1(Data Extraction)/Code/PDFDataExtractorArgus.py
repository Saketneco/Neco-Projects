import pandas as pd
import os
import pdfplumber


class PDFDataExtractorArgus:
    def __init__(self, pdf_path,download_directory,Argus_data):
        self.pdf_path = pdf_path
        self.table_settings = {
            "text_x_tolerance": 2,
            "text_y_tolerance": 2,
            "explicit_vertical_lines": [306, 348, 367, 394, 489, 518, 550]
        }
        self.download_directory = download_directory
        self.Argus_data = Argus_data

    def extract_data(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            # Select the first page (index 0)
            page = pdf.pages[0]
            
            # Extract the table with the specified settings
            tables = page.extract_tables(table_settings=self.table_settings)

            if tables:
                first_table = tables[0]
                columns = ['Energy', 'Basis', 'Timing', 'Port', '', 'Price', None, '', '±']

                name = None
                price = None

                for i, row in enumerate(first_table):
                    if row[0].lower() == 'south africa':
                        # Ensure i + 1 is within bounds
                        if i + 1 < len(first_table):
                            name = first_table[i + 1][columns.index('Port')]
                            self.Argus_data['Price'] = first_table[i + 1][columns.index('Price')]
                            price =self.Argus_data['Price']
                            #print(price)
                            break  # Exit the loop once the values are found

                return name, price
            else:
                print("No tables found on the page.")
                return None, None

    def save_to_excel(self):
        # Create a DataFrame from the dictionary
        df = pd.DataFrame([self.Argus_data])
        name = "Argus_data.xlsx"
        output_path = f"{self.download_directory}{name}"
        # Write the DataFrame to an Excel file
        df.to_excel(output_path, index=False)
        print("\nData has been written to", os.path.basename(output_path))