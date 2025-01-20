# ----------------------------------------------------------------------
# Copyright (c) 2024, Oracle and/or its affiliates. */
#
# NAME
#  app.py
#
# DESCRIPTION
# To create a python application based on this generated code, follow these steps:
# From a command line, create a new directory for your application
# and change into that directory.
# Install python oracledb using:
#   `pip install oracledb`, or 
#   `pip3 install --user oracledb`
# Open app.py in a text editor and replace the contents of the
# app.py file with this code
# Save the changes to app.py, replace the connection password tags:
# <Password>, <ProxyPassword> and <WalletPassword> with the actual passwords,
# and run the resulting application with:
# python app.py

# In case you use a Cloud DB instance and you download the credentials files
# make sure you use a password for the wallet file, and save it, as the driver
# will need it at runtime.

# To enable thick mode of python-driver use the below line after importing oracledb.
# oracledb.init_oracle_client(lib_dir="path/to/instantclient")

# For more info - https://python-oracledb.readthedocs.io/en/latest/
# -----------------------------------------------------------------------

# These are the packages/modules required for this application to run.
import os
import oracledb
import pandas as pd

Deg_list = []
# this is the SQL text that will be executed
sql_text = """
SELECT
    DISTINCT
    DESIGNATION
FROM
    V_EMPDATA
"""

# here we create a python connection,cursor object and execute the statement using  Cursor.execute().
# connections and cursors should be released when they are no longer needed.
# closed automatically when the variable referencing it goes out of scope (and no further references are retained).
# `with` block is a convenient way to ensure this.
with oracledb.connect(
  user = 'jnl',
  password = 'jnl123',
  host = '80.0.1.81', 
  port = 1521, 
  service_name = "jnil",
) as connection:
    with connection.cursor() as cursor:
        for row in cursor.execute(sql_text):
            val = list(row)
            Deg_list.append([element.replace(' ', '') for element in val])
           # print(row)

print("Done!!")

print(Deg_list)

df = pd.DataFrame(Deg_list)

# Specify the file name
file_name = 'List_of_designation.xlsx'


# Write the DataFrame to an Excel file
df.to_excel(file_name, index=False, header=False)
