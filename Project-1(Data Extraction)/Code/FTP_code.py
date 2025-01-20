import ftplib
import os
import re

def download_file(ftp_server, username, password, filename_pattern, local_directory):
    try:
        # Connect to the FTP server
        ftp = ftplib.FTP(ftp_server)
        ftp.login(user=username, passwd=password)

        # List files in the current directory
        files = ftp.nlst()

        # Find files that match the pattern
        matching_files = [f for f in files if re.match(filename_pattern, f)]
        
        if matching_files:
            # Assuming you want the first matching file
            target_filename = matching_files[0]
            print(f"\nFound file: {target_filename}")

            # Create local directory if it doesn't exist
            os.makedirs(local_directory, exist_ok=True)

            # Set the local file path
            local_file_path = os.path.join(local_directory, target_filename)

            # Open a local file for writing
            with open(local_file_path, 'wb') as local_file:
                # Define a callback to write the file in chunks
                def write_file(data):
                    local_file.write(data)

                # Download the file
                ftp.retrbinary(f'RETR {target_filename}', write_file)

            print(f"File downloaded successfully to {local_file_path}")
        else:
            print("No matching files found.")

    except ftplib.all_errors as e:
        print(f"FTP error: {e}")

    finally:
        ftp.quit()


