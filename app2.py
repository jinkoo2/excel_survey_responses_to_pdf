import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

dir = "U:/temp/qa"
xls_file = os.path.join(dir, "ReviewCases-2024.xlsx")
sheet_name = 'Form1'

# Step 1: Read the Excel File and Select the Sheet 'Form1'
excel_file_url = xls_file
data = pd.read_excel(excel_file_url, sheet_name='Form1')  # Specify the sheet name

# Step 2: Define the Function to Add Pages to the PDF
def add_pdf_page(data_row, pdf_canvas):
    width, height = letter

    # Header
    pdf_canvas.setFont("Helvetica-Bold", 14)
    pdf_canvas.drawString(100, height - 50, "External Beam and Brachytherapy - Quality Assurance Program Audit Form")
    
    # Subheader
    pdf_canvas.setFont("Helvetica-Bold", 12)
    pdf_canvas.drawString(100, height - 80, "Individual Patient Chart and Film Review Form")

    # Draw dynamic content
    pdf_canvas.setFont("Helvetica", 10)
    y = height - 120  # Start position for text

    # General Information
    pdf_canvas.drawString(100, y, f"ID: {data_row['ID']}")
    pdf_canvas.drawString(100, y - 20, f"Start time: {data_row['Start time']}")
    pdf_canvas.drawString(100, y - 40, f"Completion time: {data_row['Completion time']}")
    pdf_canvas.drawString(100, y - 60, f"Email: {data_row['Email']}")
    pdf_canvas.drawString(100, y - 80, f"Name: {data_row['Name']}")
    pdf_canvas.drawString(100, y - 100, f"MRN: {data_row['MRN']}")

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
        pdf_canvas.drawString(100, y, f"{item}: {response}")
        y -= 20  # Decrease y position for next item

        # Check for page overflow
        if y < 50:
            pdf_canvas.showPage()  # Create a new page
            y = height - 50  # Reset y position

    # Comments
    comments = data_row['Comments']
    comments = "None" if pd.isna(comments) else comments
    pdf_canvas.drawString(100, y - 20, f"Comments: {comments}")

    # Finalize the page
    pdf_canvas.showPage()

# Step 3: Create the PDF File with Multiple Pages
output_file = os.path.join(dir, "responses_summary.pdf")
pdf_canvas = canvas.Canvas(output_file, pagesize=letter)

for index, row in data.iterrows():
    add_pdf_page(row, pdf_canvas)

# Save the final PDF
pdf_canvas.save()

print(f"PDF file '{output_file}' created successfully with one page per response.")
