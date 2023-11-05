"""Author: Uygar Tolga Kara - Date: Thu Sep 28 13:30:18 2023."""
# -*- coding: utf-8 -*-

from pyodbc import connect
from os.path import join
from os import getcwd

# Define hardcoded variables
t1 = "Contact the driver and inform to drive to the work shop / garage immediately"
t2 = "Visit the workshop / garage soon"

# Define database file name and path
database_filename = 'DTC Database - ChatGPT.accdb'

# Make database path
database_path = join(getcwd(), database_filename)

# Make new connection string
connection_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    rf'DBQ={database_path};'
)

# Connect and set cursor at database
connection = connect(connection_str, autocommit=True)
cursor = connection.cursor()

tb_names = ['Powertrain', 'Body', 'Chassis', 'Network']

# Start main loop to fill database
for index, tb_name in enumerate(tb_names):

    # Select and gather row data for DTC and description
    cursor.execute(f"SELECT DTC, [DTC Description], Criticality, Recommendation FROM {tb_name}")
    rows_data = cursor.fetchall()

    for idx, row_data in enumerate(rows_data):

        # Inform developer of process
        print(f"{idx+1}/{len(rows_data)}")

        # Gather code and description data
        code, description, criticality, recommendation = row_data

        # Skip row if dtc is ISO/SAE Reserved
        if "ISO/SAE Reserved" in description:
            continue

        if criticality == "Serious":
            recom = t1
        else:
            recom = t2

        # Set operation to update
        update_query = f"""
            UPDATE {tb_name}
            SET
                Recommendation = ?
            WHERE DTC = ? AND [DTC Description] = ?"""

        cursor.execute(update_query, (recom, code, description))
        cursor.commit()

connection.commit()
connection.close()


