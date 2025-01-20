import pymysql

db_config = {
    'user': 'root',
    'password': '',  # Update with your actual password
    'host': 'localhost',
    'port': 3306,
    'database': 'employee_database',
}

try:
    print("Connecting to the database...")
    connection = pymysql.connect(**db_config)
    print("Connection successful.")
    
    # Perform your database operations here

finally:
    if connection:
        connection.close()
        print("Database connection closed.")
