import pandas as pd
import fitz  # PyMuPDF for extracting text from PDFs
from docx import Document
from google.colab import files
import os

genai.configure(api_key="AIzaSyCTZUKNEwJFDgbyWNRSrYLRxs8XvLNO4n4")  # Replace with your API key

# Load Project Database
project_df = pd.read_csv("KVA Project Database (2).csv")

# Extract Text from PDFs
def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

cv_text = extract_text_from_pdf("KVA CVs.pdf")
old_proposal_text = extract_text_from_pdf("EOI.pdf")

# Generate Proposal Function
def generate_proposal(project_name):
    project_info = project_df[project_df['Project Name'].str.contains(project_name, case=False, na=False)]

    if project_info.empty:
        return "Project not found in database."

    project_details = project_info.to_dict(orient='records')[0]

    prompt = f"""
    Based on the following old proposal template:
    {old_proposal_text}

    And the project details:
    {project_details}

    And the team members' CVs:
    {cv_text}

    Generate a new tender proposal for the project, keeping the formatting and structure of the old proposal.
    """

    response = genai.chat(model="gemini-pro", messages=[{"role": "user", "content": prompt}])
    return response.text

# Save Proposal as Word Document
def save_proposal_as_word(proposal_text, filename="Generated_Proposal.docx"):
    doc = Document()
    for para in proposal_text.split("\n"):
        doc.add_paragraph(para)
    doc_path = f"/content/{filename}"
    doc.save(doc_path)
    return doc_path

# Example Usage
project_name = "Enter Project Name Here"  # Change this
proposal_text = generate_proposal(project_name)
if "Project not found" not in proposal_text:
    doc_path = save_proposal_as_word(proposal_text)
    files.download(doc_path)  # Download file
