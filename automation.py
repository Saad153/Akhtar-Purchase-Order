import pdfplumber
import re


def is_date_format(text):
    """Check if text matches DD.MM.YYYY format"""
    return bool(re.match(r'^\d{2}\.\d{2}\.\d{4}$', text))

def parse_address_text(text):
    """Parse text to separate text before number and the number"""
    import re
    
    # Find the first number in the text
    match = re.search(r'(.*?)\s*(\d+)', text)
    if match:
        text_part = match.group(1).strip()  # Text before number
        number_part = match.group(2)  # The number
        return text_part, number_part
    return text, None

def extract_po_data(pdf_path):
    fields = {
        "PO Number": None,
        "Style Number": None,
        "Description": None,
        "PO Rel Date": None,
        "Unit Price": None,
        "PO Qty": None,
        "HOD Date": None,
        "FFC Code": None,
        "Seller": None,
        "Sourcing Type": None,
        "PO Header Text": None,
        "Country": None,
        "VAS": None,
        "Plant Code": None,
        "Season": None,
        "Brand": None
    }

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    # Extract company to sourcing type text
    company_to_sourcing = re.search(r'Company.*?(?=Sourcing|-)', text, re.DOTALL)
    company_text = company_to_sourcing.group(0) if company_to_sourcing else ""

    # Determine seller based on text content
    if 'arka' in company_text.lower():
        seller = "Arka Global"
    elif 'akhtar' in company_text.lower():
        seller = "Akhtar Textile Industries Pvt Ltd"
    else:
        seller = None


    # Regex patterns for each field
    splitted_text= text.split('Item Total Value')[1]
    country = splitted_text.split('Line Item')[0].split()[-1].split(",")[-1].strip()
    Description_2 = splitted_text.split('Line Item')[0]
    splitted_text = splitted_text.split('Date')[1].split('\n')[1].strip()
    
    description_text = splitted_text.split('Line Item')[0]
    splitted_text = splitted_text.split()
    
    fields["HOD Date"] = next(date for date in splitted_text if is_date_format(date))

    Description_1 = description_text.split(splitted_text[1])[1]
    Description_1 = Description_1.split(fields["HOD Date"])[0].strip()
    Description_2 = Description_2.split(splitted_text[-1])[1].strip().split("\n")[0].strip()
    Description_2, number = parse_address_text(Description_2)

    Description = f"{Description_1} {Description_2}".strip()

    Qty = splitted_text[-4]
    UnitPrice = splitted_text[-2]

    Description_1 = description_text.split(description_text.split()[1])[1]
    Description_1 = Description_1.split(fields["HOD Date"])

    fields["Plant Code"] = splitted_text[splitted_text.index(next(date for date in splitted_text if is_date_format(date))) + 2]
    fields["Style Number"] = splitted_text[1]    
    fields["Description"] = Description    
    
    patterns = {
        "PO Number": r"Purchase Order#\s*(\d+)",
        "PO Header Text": r"PO Header Text\s*-\s*(.*?)(?=Purchase Order Item Details)",
        "FFC Code": r"FFC Code\s*([A-Z0-9]+)",
        "Country": r"Manufacturing Country of Origin.*",
        "Sourcing Type": r"Sourcing Type\s*-\s*([^\n]+)",
        "VAS": r"Line Item VAS\s*Line Item Text\n([\s\S]+?)\nItem#",
        "Season": r"Season\s*([\w\d]+)",
        "Brand": r"Brand\s*([\w\d]+)",
        "PO Rel Date": r"PO Rel Date\s*(\d{2}\.\d{2}\.\d{4})"
    }


    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            fields[key] = match.group(1).strip()

    fields["VAS"] = " ".join(fields["VAS"].split()[1:])
    fields["Seller"] = seller
    fields["Country"] = country
    fields["Unit Price"] = UnitPrice
    fields["PO Qty"] = Qty

    return fields

if __name__ == "__main__":
    pdf_path = "azSXYx.pdf"  # Replace with your PDF path
    data = extract_po_data(pdf_path)
    for key, value in data.items():
        print(f"{key}: {value}")
