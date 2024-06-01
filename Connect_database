import mysql.connector
import pandas as pd
from datetime import datetime
import os
import re

# Replace these with your actual database connection details
host = "DB_ip_address"
user = "DB_user"
password = "DB_password"
database = "DB_name"

# Adjust connection to DataBase
config = {
    "host": host,
    "user": user,
    "password": password,
    "database": database
}

def clean_data(value):
    """
    Clean data by removing or replacing problematic characters.
    """
    if isinstance(value, str):
        # Replace problematic characters with a space or remove them
        value = re.sub(r'[\x00-\x1f\x7f-\x9f]', ' ', value)
    return value

# Check table in Database
try:
    print("Connecting to the database...")
    cnx = mysql.connector.connect(**config)

    if cnx.is_connected():
        print("Connection established successfully.")
        with cnx.cursor(dictionary=True) as cursor:  # Use dictionary=True to get column names
            print("Fetching table names...")
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()

            # Print table names
            print("Tables in the database:")
            for table in tables:
                print(table)
                
except mysql.connector.Error as err:
    print(f"Error: {err}")
except Exception as e:
    print(f"Unexpected error: {e}")

# Query Database to .xlsx file
try:
    print("Connecting to the database...")
    cnx = mysql.connector.connect(**config)

    if cnx.is_connected():
        print("Connection established successfully.")
        with cnx.cursor(dictionary=True) as cursor:  # Use dictionary=True to get column names
            print("Executing SQL command...")
            query = """
                SELECT 
                FROM 
                WHERE
            """

            cursor.execute(query)

            print("Fetching data...")
            rows = cursor.fetchall()

            print(f"Number of rows fetched: {len(rows)}")

            # Convert to DataFrame
            df = pd.DataFrame(rows)

            # Clean the data
            for column in df.columns:
                df[column] = df[column].apply(clean_data)

            # Save to Excel with proper encoding
            # Get the current date
            current_date = datetime.today().strftime('%Y-%m-%d')
            output_directory = "C:\\Pam\\Proj\\DB_report\\" #path directory
            output_filename = f"phone_report_{current_date}_no_cut.xlsx" 
            
            # Combine the directory path and filename
            output_file = os.path.join(output_directory, output_filename)

            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Data saved to {output_file}")

        print("Closing connection...")
        cnx.close()
        print("Connection closed successfully.")

    else:
        print("Could not connect to the database.")

except mysql.connector.Error as err:
    print(f"Error: {err}")

except Exception as e:
    print(f"Unexpected error: {e}")
