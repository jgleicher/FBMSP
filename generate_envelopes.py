import os
import sys
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from PyPDF2 import PdfMerger
from datetime import date


# ----- CONFIGURATION -----
RETURN_ADDRESS = [
    "Friends of Blue Mound State Park",
    "PO Box 287",
    "Mount Horeb, WI 53572"
]
LOGO_FILENAME = "BLUEMOUNDS-F-01.png"
OUTPUT_DIR = "Envelopes"
COMBINED_FILENAME = "Combined_Envelopes.pdf"
ENVELOPE_SIZE = (241, 105)  # #10 envelope in mm (landscape)

# ----- MAIN FUNCTION -----
def generate_envelopes(csv_file):
    if not os.path.isfile(csv_file):
        print(f"❌ Error: File '{csv_file}' not found.")
        return

    if not os.path.exists(LOGO_FILENAME):
        raise FileNotFoundError(f"❌ Logo file '{LOGO_FILENAME}' is required but not found.")

    df = pd.read_csv(csv_file).fillna("")
    if "Account Name" in df.columns:
        df.sort_values("Account Name", inplace=True)
    elif "First Name" in df.columns:
        df.sort_values("First Name", inplace=True)
    else:
        print("⚠️  Warning: Neither 'Account Name' nor 'First Name' columns found. Skipping sort.")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    individual_paths = []

    for _, row in df.iterrows():
        first = str(row.get("First Name", "")).strip()
        last = str(row.get("Last Name", "")).strip()
        account = str(row.get("Account Name", "")).strip()
        if not account:
            account = f"{first} {last}".strip()
        name_line = f"{first} {last}".strip() if first or last else account

        def get_field(row, *keys):
            for key in keys:
                val = str(row.get(key, "")).strip()
                if val:
                    return val
            return ""

        street = get_field(row, "Address (Street)", "Street Address")
        city = get_field(row, "Address (City)", "City")
        state = get_field(row, "Address (State/Province)", "State")
        zip_code = get_field(row, "Address (Postal/Zip Code)", "Zip Code")
        city_state_zip = f"{city}, {state} {zip_code}"

        pdf = FPDF(unit="mm", format=ENVELOPE_SIZE)
        pdf.add_page()
        pdf.set_auto_page_break(False)

        # Logo
        pdf.image(LOGO_FILENAME, x=10, y=2, w=30)

        # Return address slightly higher (start at y=4.5 instead of 6.5)
        pdf.set_font("helvetica", size=10)
        for i, line in enumerate(RETURN_ADDRESS):
            pdf.set_xy(40, 7.5 + i * 4.5)  # Moved down by 3mm
            pdf.cell(0, 5, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")

        # Delivery address centered
        pdf.set_font("helvetica", size=12)
        x_text = (241 - 90) / 2
        y_text = 62.7  # Moved down by 1/2 inch (12.7mm)
        for line in [name_line, street, city_state_zip]:
            pdf.set_xy(x_text, y_text)
            pdf.cell(90, 6, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
            y_text += 6
        safe_name = account.replace(" ", "_").replace("/", "-")
        output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_Envelope.pdf")
        pdf.output(output_path)
        individual_paths.append(output_path)

    # Combine all
    merger = PdfMerger()
    for path in individual_paths:
        merger.append(path)
    combined_path = os.path.join(".", COMBINED_FILENAME)
    merger.write(combined_path)
    merger.close()

    print(f"✅ All envelopes generated in: {OUTPUT_DIR}")
    print(f"📄 Combined PDF: {combined_path}")


# ----- ENTRY POINT -----
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_envelopes.py <members.csv>")
    else:
        generate_envelopes(sys.argv[1])

