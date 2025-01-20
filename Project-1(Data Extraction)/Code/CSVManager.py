import os
import logging
import glob
import pandas as pd


class CSVManager:

    def list_xlsx_files(directory):
        """List the xlsx files in the given directory."""
        return glob.glob(os.path.join(directory, '*.xlsx'))


    def read_xlsx_files(xlsx_files):
        """Read the xlsx files and return data matrices."""
        data_matrices = {}
        for filename in xlsx_files:
            if filename is None:
                continue
            try:
                # Check if the file is a valid Excel file
                xls = pd.ExcelFile(filename)
                if len(xls.sheet_names) == 0:
                    print(f"No worksheets found in {filename}.")
                    continue
                    
                # Read the first sheet (you can modify this to read specific sheets if needed)
                df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
                data_matrices[os.path.basename(filename)] = {
                    'data': df.values.tolist(),
                    'columns': df.columns.tolist()
                }
            except Exception as e:
                print(f"Error reading {filename}: {e}")

        return data_matrices


    def add_ID_column(data_matrices, add_material_code):
        """Add the 'Material Code' column to the data matrices."""
        #print(data_matrices)
        for file_name, data in data_matrices.items():
           # print(file_name,"#######",add_material_code)
            if file_name in add_material_code:
                for i in range(len(add_material_code[file_name])):
                    material_code_name = add_material_code[file_name][i][0]
                    material_code_value = add_material_code[file_name][i][1]

                    data['columns'].insert(1, material_code_name)
                    for row in data['data']:
                        row.insert(1, material_code_value)
        return data_matrices

    def filter_data(data, columns, conditions):
        """Filter the data based on the given conditions."""
        #print(data)
        #print(columns)
        df = pd.DataFrame(data, columns=columns)
        temp_list = []
        #print(df['Sl.no.'])
        for column_name, value in conditions:
            
            if column_name not in df.columns:
                logging.warning(f"\nColumn '{column_name}' does not exist in the DataFrame.")
                return pd.DataFrame()
           # print(column_name)
           # print(value)
           # elif df[column]==value:
                
            df = df[df[column_name] == value].drop(columns=[column_name,'Sl.no.'])
           # temp_list = df.to_numpy().tolist()
          #  print(df)
        return df

    def remove_column(data,columns):
       sl_no_index = columns.index('Sl.no.')
       new_data = [row[:sl_no_index] + row[sl_no_index + 1:] for row in data]
       new_columns = [col for col in columns if col != 'Sl.no.']
       
       return new_data,new_columns
   
    def apply_filter_conditions(data_matrices, filter_conditions):
        """Apply filter conditions to the data matrices."""
        filtered_dfs = []
        for filename, conditions in filter_conditions.items():
            #print(filename,"*********",data_matrices,"\n")
            if filename in data_matrices:
                
                data = data_matrices[filename]['data']
                columns = data_matrices[filename]['columns']
                #new_data,new_column = CSVManager.remove_column(data,columns)
                filtered_df = CSVManager.filter_data(data, columns, conditions)
                #print(filtered_df)
                #new_data,new_column = CSVManager.remove_column(data,columns)
                #print(filtered_df)
                if len(filtered_df) != 0:
                    filtered_dfs.append(filtered_df)
        return filtered_dfs



    def concatenate_and_save(filtered_dfs, output_file):
        """Concatenate filtered dataframes and save to excel."""
        text_columns =["ID"]
        if filtered_dfs:
            final_filtered_df = pd.concat(filtered_dfs, ignore_index=True)
            if 'Sl.no.' in final_filtered_df.columns:
                final_filtered_df['Sl.no.'] = range(1, len(final_filtered_df) + 1)
            if 'Description' in final_filtered_df.columns:
               # print(final_filtered_df.columns)
                final_filtered_df.rename(columns={'Description': 'Title'}, inplace=True)  
            final_filtered_df.to_excel(output_file, index=False, engine='openpyxl')

            # Optionally, set the format to Text in the saved Excel file
            # wb = openpyxl.load_workbook(output_file)
            # ws = wb.active
    
            # # Apply text formatting to specified columns
            # for col in text_columns:
            #     if col in final_filtered_df.columns:
            #         col_index = final_filtered_df.columns.get_loc(col) + 1  # Get column index (1-based)
            #         for row in ws.iter_rows(min_row=2, min_col=col_index, max_col=col_index):  # Skip header
            #             for cell in row:
            #                 cell.number_format = '@'  # Set format to Text

            # wb.save(output_file)
            # print(f"Data successfully written to {output_file}")
        else:
            logging.warning("\nNo data matched the filtering conditions.")


    def delete_pdf_files(directory):
        """Delete all PDF files in the given directory."""
        if not os.path.exists(directory):
            logging.warning(f"\nThe directory '{directory}' does not exist.")
            return

        for filename in os.listdir(directory):
            if filename.endswith('.pdf'):
                file_path = os.path.join(directory, filename)
                try:
                    os.remove(file_path)
                    logging.info(f"Deleted: {file_path}")
                except Exception as e:
                    logging.error(f"\nError deleting {file_path}: {e}")