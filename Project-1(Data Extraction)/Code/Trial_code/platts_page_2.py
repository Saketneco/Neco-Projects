import pandas as pd
import pdfplumber
import logging



class PDFTableExtractorPlatts:
    def __init__(self, pdf_path,download_directory,filter_data,Platts_Id):
        self.pdf_path = pdf_path
        self.download_directory = download_directory
        self.filter_data = filter_data
        self.Platts_Id = Platts_Id

    def extract_table_platts_pdf(self, page_number):
        with pdfplumber.open(self.pdf_path) as pdf:
            page = pdf.pages[page_number - 1]  # Convert to 0-based index
            #print(page)
            tables = self.extract_tables(page)
            # for table in tables:
            #     for row in table:
            #         print(row)

            print(f"Number of tables extracted: {len(tables)}\n")

            if tables:
                specific_table_index = 0
                if len(tables) > specific_table_index:
                    table = tables[specific_table_index]
                   # print(table)
                    output_list = self.process_table(table ,page)
                   # print(output_list)
                    
                   # self.save_to_excel_with_text_conversion(output_list, "Platts_data.xlsx" ,page)
                else:
                    print(f"\nNo table found at index {specific_table_index}.")
            else:
                print("\nNo tables found on this page.")
                
        return output_list        

    def extract_tables(self, page):
        if page.page_number == 7:
            # Table settings for page 1
            print("page7")
            table_settings_page_7 = {
                "vertical_strategy": "text",
                "horizontal_strategy": "lines_strict",
                "explicit_vertical_lines": [350, 310, 270, 220, 170],
                "explicit_horizontal_lines": [113],
                "snap_tolerance": 8,
                "join_tolerance": 1,
                "edge_min_length": 2,
                "min_words_vertical": 2,
                "min_words_horizontal": 2,
                "intersection_tolerance": 2,
                "text_tolerance": 2,
            }
            return page.extract_tables(table_settings=table_settings_page_7)

        elif page.page_number == 2:
            # Table settings for page 2
            print("page2")
            table_settings_page_2 = {
             "horizontal_strategy": "lines",
            "explicit_vertical_lines": [304.0 ,409,424,500,530,560],
            "explicit_horizontal_lines": [80,160, 179 , 252, 265],
             "snap_tolerance": 3,
             "join_tolerance": 1,
            }
            return page.extract_tables(table_settings=table_settings_page_2)

          


    def process_table(self, table ,page):
        df = pd.DataFrame(table)
        #for row in df:    
        #print(df)
        table_as_list = df.values.tolist()
       # print(table_as_list)
        output_list = {}

        if page.page_number == 7:
            
            for row in self.filter_data.keys():
               # print(table_as_list)
                for i, row2 in enumerate(table_as_list):
                    if row == row2[0]:
                        #print( row,row2[1])
                        column_list = []
                        for j, column in enumerate(table_as_list[1]):
                            if self.filter_data[row][0]  == column.replace("\n", " ") or \
                                self.filter_data[row][1] == column.replace("\n", " "):
                                column_list.append(table_as_list[i][j])

                        output_list[row] = column_list    
                       # print(output_list)
                        break
                    
        elif(page.page_number == 2):  
             print("page==2")
             for row in self.filter_data.keys():
                #print(table_as_list)
                for i, row2 in enumerate(table_as_list):
                    if row == row2[1]:
                       # print(row2[1])
                       # print( self.filter_data[row][0])
                        column_list = []
                        for j, column in enumerate(table_as_list[8]):
                            if self.filter_data[row][0]  == column:
                                #print(column)
                                column_list.append(table_as_list[i][j])
                                column_list.append('')
                               # print(column_list)
                        #
                        output_list[row] = column_list
                        
                        break
        return output_list

    def insert_column(self,row_list,columns):
        
        row_list = [list(item) for item in row_list]
        #print(row_list,columns)
        
        for key,value in self.Platts_Id.items():
            #print(value[0][0])
            if value[0][0] not in columns:
                columns.insert(0,value[0][0])
                columns.insert(len(columns)-2,'Currency')
                #print(columns)
                
            for row in row_list:
                if row[0]== key:
                    
                    row.insert(0,value[0][1])
                   # print(row)
                    row.insert(len(row)-2,"USD")
                   # print(row)
                    break   
            
        return row_list , columns       
            
    # def save_to_csv(self, output_list, output_filename):
    #     columns = ["Comodity Name"] + list(self.filter_data.values())[0]
    #     row_list = [(name, *value) for name, value in output_list.items()]
    #     #print(row_list)
    #     row_list,columns = self.insert_column(row_list,columns)
    #     df = pd.DataFrame(row_list, columns=columns)
    #     #df = df.convert_dtypes(convert_string=True)
    #     #print(df.dtypes)
    #     output_path = f"{self.download_directory}{output_filename}"
    #     df.to_csv(output_path, index=False)
    #     #df.to_excel(output_path, index=False, engine='openpyxl')
    #     print(f"Data successfully written to {output_filename}")


    def save_to_excel_with_text_conversion(self, output_list, output_filename):
        # Define the columns based on provided data
        #print(output_list)
        text_columns =["ID"]
        columns = ["Title"] + list(self.filter_data.values())[0]
        row_list = [(name, *value) for name, value in output_list.items()]

        #print(row_list,columns)
        # Insert additional columns if needed
        row_list, columns = self.insert_column(row_list, columns)
        #print(row_list,columns)
        columns = [col.replace('FOB Australia', 'Price') for col in columns]
        
        # Create a DataFrame
        modified_row_list = [row[:-1] for row in row_list]
        #print(modified_row_list,columns[:-1])
        
        df = pd.DataFrame(modified_row_list, columns=columns[:-1])
        #print(df)

        # Ensure the output filename has the correct extension
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'

        # Create the full output path
        output_path = f"{self.download_directory}{output_filename}"

        # Convert specified columns to text (if any)
        if text_columns:
            for col in text_columns:
                if col in df.columns:
                    df[col] = df[col].astype(str)  # Convert column to string

        # Save to Excel
        df.to_excel(output_path, index=False, engine='openpyxl')

        # # Optionally, set the format to Text in the saved Excel file
        # wb = openpyxl.load_workbook(output_path)
        # ws = wb.active
    
        # # Apply text formatting to specified columns
        # for col in text_columns:
        #     if col in df.columns:
        #         col_index = df.columns.get_loc(col) + 1  # Get column index (1-based)
        #         for row in ws.iter_rows(min_row=2, min_col=col_index, max_col=col_index):  # Skip header
        #             for cell in row:
        #                 cell.number_format = '@'  # Set format to Text

        # wb.save(output_path)
        # print(f"Data successfully written to {output_path}")
        
        
if __name__ == "__main__":
    
    
    DOWNLOAD_DIRECTORY = f"G:\\Shared drives\\Metal Prices\\Market_price\\"
    latest_ict_file = "G:\\My Drive\\PDF_files\\ICT_20241210.pdf"
    
    if not latest_ict_file:
     logging.warning("No ICT file provided for Platts extraction.")

    filter_data = {
        "Premium Low Vol": ["FOB Australia", "CFR India"],
        "Low Vol HCC": ["FOB Australia", "CFR India"],
        "Low Vol PCI": ["FOB Australia", "CFR India"],
        "Mid Vol PCI": ["FOB Australia", "CFR India"],
        "Semi Soft": ["FOB Australia", "CFR India"],
        "Richards Bay-India West" :["$/mt"]
    }

    Platts_Id = {
        "Premium Low Vol": [("ID", "50000141001")],
        "Low Vol HCC": [("ID", "50000141012")],
        "Low Vol PCI": [("ID", "50000151001")],
        "Mid Vol PCI": [("ID", "50000151002")],
        "Semi Soft": [("ID", "50000151003")],
        "Richards Bay-India West" : [("ID", "50000151004")]
    }

    try:
        pdf_extractor = PDFTableExtractorPlatts(latest_ict_file, DOWNLOAD_DIRECTORY, filter_data, Platts_Id)
        
        # Define the indices in a list for better maintainability
        indices_to_process = [7, 2]

        # Extract tables and store them in output_lists
        output_lists = [pdf_extractor.extract_table_platts_pdf(i) for i in indices_to_process]

        # Merge the dictionaries from the output_lists into a single dictionary
        merged_dict = {key: value for d in output_lists for key, value in d.items()}

        # If output_lists is not empty, save the merged data to an Excel file
        if output_lists:
            pdf_extractor.save_to_excel_with_text_conversion(merged_dict, "Platts_data.xlsx")

        print(output_lists)

    except FileNotFoundError:
        logging.error(f"Error: The file '{latest_ict_file}' was not found.")
    except Exception as e:
        logging.error(f"An error occurred during the extraction process: {e}")