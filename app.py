import os
dir = "U:/temp/qa"
xls_file = os.path.join(dir, "ReviewCases-2024.xlsx")
sheet_name = 'Form1'

print(xls_file)

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Step 1: Read Excel File
excel_file_url = xls_file
data = pd.read_excel(excel_file_url, sheet_name=sheet_name)

# Step 2: Define a Function to Create PDFs
def create_pdf(data_row, filename):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter

    # Add form headers or static content
    c.drawString(100, height - 50, "External Beam and Brachytherapy - Quality Assurance Program Audit Form")

    # Add dynamic content from the Excel row
    c.drawString(100, height - 100, f"Patient Name: {data_row['Patient Name']}")
    c.drawString(100, height - 120, f"Tumor Staged: {data_row['Tumor Staged']}")
    c.drawString(100, height - 140, f"Treatment Type: {data_row['Treatment Type']}")

    # Add more fields as per your requirements

    # Save the PDF
    c.save()


from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def create_pdf2(data_row, filename):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, height - 50, "External Beam and Brachytherapy - Quality Assurance Program Audit Form")
    
    # Subheader
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, height - 80, "Individual Patient Chart and Film Review Form")

    # Draw dynamic content
    c.setFont("Helvetica", 10)
    y = height - 120  # Start position for text

    # General Information
    c.drawString(100, y, f"ID: {data_row['ID']}")
    c.drawString(100, y - 20, f"Start time: {data_row['Start time']}")
    c.drawString(100, y - 40, f"Completion time: {data_row['Completion time']}")
    c.drawString(100, y - 60, f"Email: {data_row['Email']}")
    c.drawString(100, y - 80, f"Name: {data_row['Name']}")
    c.drawString(100, y - 100, f"MRN: {data_row['MRN']}")

    # Checklist (boolean or text responses)
    checklist_items = [
        "Is there a history and physical on the chart?",
        "Are tumors staged?",
        "Is there a pathology report",
        "Are there appropriate x-ray reports",
        "Is there a signed prescription that includes the area to be treated, technique, energy, dose fractionation and the total dose plus limits to critical structures (if applicable)?",
        "Is the plan appropriate for tumor state and type?",
        "Do the treatment fields adequately cover the tumor?",
        "Is there a signed informed consent?",
        "Are there ID photos and field photos",
        "Are there periodic progress notes?",
        "Is there a completion note?"
    ]
    
    y -= 140  # Adjust for spacing
    for item in checklist_items:
        response = data_row.get(item, "N/A")  # Get response or default to "N/A"
        c.drawString(100, y, f"{item}: {response}")
        y -= 20  # Decrease y position for next item

        # Check for page overflow
        if y < 50:
            c.showPage()  # Create a new page
            y = height - 50  # Reset y position

    # Comments
    c.drawString(100, y - 20, f"Comments: {data_row['Comments']}")

    # Save the PDF
    c.save()



# Step 3: Loop through Data and Generate PDFs
output_dir = os.path.join(dir, "output_pdfs")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

for index, row in data.iterrows():
    pdf_filename = os.path.join(output_dir, f"response_{index + 1}.pdf")
    create_pdf2(row, pdf_filename)

print("PDFs generated successfully.")



