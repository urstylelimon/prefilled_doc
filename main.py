import pandas as pd
from docx import Document
import os

# Load Excel file using pandas
excel_data = pd.read_excel('Membership_Address.xlsx')

# Template Word document path
template_doc_path = 'Membership Information Collection Form.docx'

# Directory to save the generated documents
output_dir = 'generated_docs'
os.makedirs(output_dir, exist_ok=True)  # Create directory if it doesn't exist

def replace_placeholders(doc, placeholders):
    """
    Replaces placeholders in the Word document's paragraphs and tables.
    """
    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        for key, value in placeholders.items():
            if key in para.text:
                para.text = para.text.replace(key, value)

    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in placeholders.items():
                        if key in para.text:
                            para.text = para.text.replace(key, value)

# Loop over each row in the Excel sheet
for index, row in excel_data.iterrows():
    # Extract member data from the row
    name = row['Name']
    address = row['Address']
    member_type = row['MemberType']  # AM, GM, LM, etc.
    member_id = row['MemberID']

    # Load the template Word document
    doc = Document(template_doc_path)

    # Dictionary to map placeholders in the Word document to actual data
    placeholders = {
        '{name}': name,
        '{address}': address,
        '{member_id}': str(member_id),
        '{AM}': '✓' if member_type == 'AM' else '',
        '{GM}': '✓' if member_type == 'GM' else '',
        '{LM}': '✓' if member_type == 'LM' else '',
        '{SM}': '✓' if member_type == 'SM' else '',
        '{DM}': '✓' if member_type == 'DM' else '',
        '{TRLM}': '✓' if member_type == 'TRLM' else '',
        '{TRSM}': '✓' if member_type == 'TRSM' else ''
    }

    # Replace placeholders with actual values
    replace_placeholders(doc, placeholders)

    # Save the filled Word document for each member
    output_filename = f"{name.replace(' ', '_')}_{member_id}.docx"
    output_filepath = os.path.join(output_dir, output_filename)
    doc.save(output_filepath)

    print(f"Generated document for {name}: {output_filepath}")

print("All documents have been generated successfully!")
