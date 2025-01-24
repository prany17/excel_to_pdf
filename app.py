# imports
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import pandas as pd
import openpyxl
import reportlab
import os

# load the excel file
excel_file = 'student_scores.xlsx'
try:    
    data = pd.read_excel(excel_file)
except FileNotFoundError as error:
    print(f"exception:  {error} ")
except Exception as error:
    print(f"error: {error}")

# check if the columns exists
required_columns = {"Student ID", "Name", "Math", "Science", "English", "History", "Art"}
try:
    if not required_columns.issubset(data.columns):
        raise ValueError(f"The {required_columns} does not exists")
except ValueError as error:
    print(error)

# calculate total and average scores
data["Total Score"] = data[["Math", "Science", "English", "History", "Art"]].sum(axis=1)
data["Average Score"] = data[["Math", "Science", "English", "History", "Art"]].mean(axis=1)

# generate report card pdf's
    # what do we need in the pdf
    # 1. Title ("Report Card")
    # 2. Student Name
    # 3. Total Score
    # 4. Average Score
    # Table of Marks for each subject

# directory to save the report cards
output_dir = "report_cards"
os.makedirs(output_dir, exist_ok=True)

# variables for the data in the df
for _,row in data.iterrows():
    student_id  = row["Student ID"]
    student_name = row["Name"]
    total_score = row["Total Score"]
    average_score = row["Average Score"]
    subject_scores = row[["Math","Science", "English", "History", "Art"]].to_dict()

    # creating the report cards
    
    # pdf name and size
    pdf_filename = os.path.join(output_dir, f"report_card_{student_id}.pdf")
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)

    styles = getSampleStyleSheet()
    elements = []

    # Add student name
    elements.append(Paragraph(f"<b>Report Card</b>", style=styles['Title']))
    elements.append(Paragraph(f"Student Name: {student_name}", style=styles['Normal']))
    elements.append(Paragraph(f"Student Name: {total_score}", style=styles['Normal']))
    elements.append(Paragraph(f"Student Name: {average_score}", style=styles['Normal']))

    # create table for subject scores
    table_data = [["Subject", "Score"]] 
    for subject, score in subject_scores.items():
        table_data.append([subject, score])   

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)

    # Build PDFs
    try:
        doc.build(elements)
        print(f"Report card generated: {pdf_filename}")
    except Exception as e:
        print(f"Error generating pdf for {student_name}: {e}")

print("All report cards generated sucessfully.")
