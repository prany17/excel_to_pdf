import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os

# Load the Excel file
try:
    excel_file = "student_scores.xlsx"  # Replace with the path to your Excel file
    data = pd.read_excel(excel_file)
except FileNotFoundError:
    print("Error: Excel file not found. Please check the file path.")
    exit()
except Exception as e:
    print(f"Error: {e}")
    exit()

# Validate data
required_columns = {"Student ID", "Name", "Math", "Science", "English", "History", "Art"}
if not required_columns.issubset(data.columns):
    print("Error: Missing required columns in the Excel file.")
    exit()

# Group data by student and calculate total and average scores
data["Total Score"] = data[["Math", "Science", "English", "History", "Art"]].sum(axis=1)
data["Average Score"] = data[["Math", "Science", "English", "History", "Art"]].mean(axis=1)
# total_sum_of_ids = data["Student ID"].sum(axis=0)
print(data)

# Generate PDF report cards
output_dir = "report_cards"
os.makedirs(output_dir, exist_ok=True)

for _, row in data.iterrows():
    student_id = row["Student ID"]
    student_name = row["Name"]
    total_score = row["Total Score"]
    average_score = row["Average Score"]
    subject_scores = row[["Math", "Science", "English", "History", "Art"]].to_dict()

    # Create a PDF for each student
    pdf_filename = os.path.join(output_dir, f"report_card_{student_id}.pdf")
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)

    styles = getSampleStyleSheet()
    elements = []

    # Add student name
    elements.append(Paragraph(f"<b>Report Card</b>", styles["Title"]))
    elements.append(Paragraph(f"Student Name: {student_name}", styles["Normal"]))
    elements.append(Paragraph(f"Total Score: {total_score}", styles["Normal"]))
    elements.append(Paragraph(f"Average Score: {average_score:.2f}", styles["Normal"]))

    # Create a table for subject scores
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

    # Build PDF
    try:
        doc.build(elements)
        print(f"Report card generated: {pdf_filename}")
    except Exception as e:
        print(f"Error generating PDF for {student_name}: {e}")

print("All report cards generated successfully.")
