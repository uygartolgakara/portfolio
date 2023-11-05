"""Author: Uygar Tolga Kara - Date: Mon Oct 23 21:52:48 2023."""
# -*- coding: utf-8 -*-

import pandas as pd
import pyodbc
from fpdf import FPDF
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import TfidfVectorizer

# Set pandas to suppress notifications
pd.options.mode.chained_assignment = None

print("\nWarning, database and input files needs to have same columns as following: \n")

print("""DTC
DTC Description
DTC Category
FTB Category
Possible Causes
Possible Symptoms
Criticality
Recommendation
Additional Information
Review from Raghu
Review from Coc
Review Status\n""")

access = input("Access database path: ")
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + access + ';')

table_name = input("Powertrain, Body, Chassis or Network? ")
query = "SELECT * FROM " + table_name

df1 = pd.read_sql(query, conn)
conn.close()

excel = input("Path to excel file: ")
df2 = pd.read_excel(excel)

# Assuming 'code' is the column name that you match between input and database
common_codes = set(df1['DTC']).intersection(set(df2['DTC']))

similarities = []

for code in common_codes:
    df1_row = df1[df1['DTC'] == code].iloc[0]
    df2_row = df2[df2['DTC'] == code].iloc[0]

    avg_similarity = 0
    count = 0
    for col in df1.columns:

        val1, val2 = df1_row[col], df2_row[col]
        if pd.notna(val1) and pd.notna(val2):
            try:
                val1, val2 = str(val1).lower().strip(), str(val2).lower().strip()
                if val1 == val2:
                    avg_similarity += 1
                else:
                    tfidf_vectorizer = TfidfVectorizer()
                    tfidf_matrix = tfidf_vectorizer.fit_transform([val1, val2])
                    cosine_sim = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
                    avg_similarity += max(0, min(cosine_sim, 1))  # Ensure cosine_sim is within [0, 1]
                count += 1
            except ValueError:
                continue

    if count > 0:
        avg_similarity /= count
        similarities.append((code, avg_similarity))
    else:
        similarities.append((code, 0))

# Sort the codes by similarity
sorted_similarities = sorted(similarities, key=lambda x: x[1], reverse=True)
print(sorted_similarities)

def generate_pdf_report(data):

    pdf = FPDF()
    pdf.add_page()

    # Set font
    pdf.set_font("Arial", size=12)

    # Add title
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Similarity Report", 0, 1, 'C')

    # Add column headers
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(100, 10, "Label", 1)
    pdf.cell(100, 10, "Similarity (%)", 1)
    pdf.ln()

    # Add data
    pdf.set_font("Arial", size=12)
    for label, similarity in data:
        pdf.cell(100, 10, str(label), 1)
        pdf.cell(100, 10, f"{similarity*100:.2f}%", 1)
        pdf.ln()

    # Save the PDF
    pdf.output("Report.pdf")

generate_pdf_report(sorted_similarities)

# C:\Users\KUY3IB\Desktop\amazon\DTC_Database_Backup.accdb
# Powertrain
# C:\Users\KUY3IB\Desktop\amazon\Powertrain.xlsx