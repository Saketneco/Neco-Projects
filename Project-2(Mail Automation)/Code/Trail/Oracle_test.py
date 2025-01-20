import os
import sys

# Set the path to the Oracle Instant Client directory
instant_client_path = r"C:\oracle\instantclient_19_8"
os.environ["PATH"] = instant_client_path + ";" + os.environ["PATH"]

# Now import cx_Oracle after setting the PATH
import cx_Oracle  # type: ignore
