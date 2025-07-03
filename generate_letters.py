import csv
import os
import sys
from datetime import datetime
from docx import Document
from PyPDF2 import PdfMerger
import subprocess
import concurrent.futures

# === CONFIGURATION ===
TEMPLATE_FILE = "New Member Letter Template.docx"  # should be in current directory
OUTPUT_DOCX_DIR = "letters"
OUTPUT_PDF_DIR = "pdfs"
COMBINED_PDF = "Combined_Letters.pdf"

# === FUNCTIONS ===
def normalize_keys(row):
    return {k.strip().lower(): v.strip() for k, v in row.items()}

def determine_account_type(row):
    if row.get("business friend of blue mound state park", "").strip():
        return "business"
    elif row.get("family friend of blue mound state park", "").strip():
        return "family"
    elif row.get("individual friend of blue mound state park", "").strip():
        return "individual"
    elif row.get("volunteer/working friend", "").strip():
        return "volunteer"
    return "individual"

def determine_salutation(row):
    account_type = determine_account_type(row)
    if account_type == "business":
        return row.get("account name", "Friend")
    elif account_type == "family":
        return f"{row.get('last name', 'Friend')} family"
    else:
        return row.get("first name", "Friend")

def determine_name(row):
    account_type = determine_account_type(row)
    if account_type == "business":
        return row.get("account name", "Friend")
    else:
        return f"{row.get('first name', '')} {row.get('last name', '')}".strip()

def fill_template(template_path, output_path, replacements):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for key, val in replacements.items():
            if key in para.text:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)
    doc.save(output_path)

def convert_to_pdf(docx_path, pdf_path):
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir",
                        os.path.dirname(pdf_path), docx_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error converting {docx_path} to PDF: {e}")

def get_amount(row):
    for key in ["amount after fees", "donation", "amount"]:
        val = row.get(key, "").strip()
        if val:
            return val
    return ""

def process_letter(row, today_str, template_file):
    account_type = determine_account_type(row)
    salutation = determine_salutation(row)
    name = determine_name(row)
    filename_base = row.get("account name") or name.replace(" ", "_")
    filename_base = filename_base.replace("/", "-")

    city = row.get("address (city)", "")
    state = row.get("address (state/province)", "")
    postal_code = row.get("address (postal code)", "")
    city_state_zip = f"{city}, {state} {postal_code}".strip()
    if city_state_zip.startswith(","):
        city_state_zip = city_state_zip[1:].strip()
    address_line = row.get("address (street)", "").strip()
    full_address_block = f"{address_line}\n{city_state_zip}".strip().rstrip(",")
    amount = get_amount(row)
    replacements = {
        "SALUTATION": salutation,
        "NAME": name,
        "ADDRESS": full_address_block,
        "CITY_STATE_ZIP": "",  # prevent accidental second line
        "DATE": today_str,
        "AMOUNT": amount or (
            "100" if account_type == "business" else
            "45" if account_type == "family" else
            "25" if account_type == "individual" else
            "25"
        )
    }
    docx_path = os.path.join(OUTPUT_DOCX_DIR, f"{filename_base}.docx")
    pdf_path = os.path.join(OUTPUT_PDF_DIR, f"{filename_base}.pdf")
    fill_template(template_file, docx_path, replacements)
    convert_to_pdf(docx_path, pdf_path)
    print(f"Generated: {pdf_path}")
    return pdf_path

def main(csv_files, template_file=TEMPLATE_FILE):
    if not os.path.exists(template_file):
        print(f"Template file '{template_file}' not found.")
        return

    os.makedirs(OUTPUT_DOCX_DIR, exist_ok=True)
    os.makedirs(OUTPUT_PDF_DIR, exist_ok=True)

    today_str = datetime.today().strftime("%B %d, %Y")
    all_rows = []
    for csv_file in csv_files:
        with open(csv_file, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for raw_row in reader:
                row = normalize_keys(raw_row)
                all_rows.append(row)

    generated_pdfs = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(process_letter, row, today_str, template_file) for row in all_rows]
        for future in concurrent.futures.as_completed(futures):
            pdf_path = future.result()
            generated_pdfs.append(pdf_path)

    # Combine PDFs
    if generated_pdfs:
        merger = PdfMerger()
        for pdf in sorted(generated_pdfs):
            merger.append(pdf)
        merger.write(COMBINED_PDF)
        merger.close()
        print(f"Combined PDF created at: {COMBINED_PDF}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_letters.py file1.csv [file2.csv ...] [TEMPLATE_FILE]")
        sys.exit(1)
    args = sys.argv[1:]
    if len(args) > 1 and args[-1].lower().endswith('.docx') and os.path.isfile(args[-1]):
        template_file = args[-1]
        csv_files = args[:-1]
    else:
        template_file = TEMPLATE_FILE
        csv_files = args
    main(csv_files, template_file)
