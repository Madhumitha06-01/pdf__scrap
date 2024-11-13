import fitz  # PyMuPDF
import pandas as pd
import re

# Path to the PDF file
pdf_path = r'C:\selenium2\Healthshine Catalogue - Copy.pdf'

# Regular expressions to capture other product details
product_name_pattern = r"^(?!.*\b(91|CONTACT|\d{10})\b)([A-Z0-9\-/()&\s]+)$"
highlights_pattern = r"Highlights:\n(.+?)(?=\nFeatures|\n\+91)"  # Highlights up to 'Features' or '+91'
features_pattern = r"Features\n(.+?)(?=\n\+91|\n)"  # Features up to '+91'
mrp_pattern = r"MRP\s+(\d+/-(?=\n))"
material_pattern = r"Material\s*:\s*(.+?)(?=\n)"
dimensions_pattern = r"Width\s*:\s*(\d+\s*CM)\s*Length\s*:\s*(\d+\s*CM)\s*Height\s*:\s*(\d+\s*CM)"  # Dimensions
color_pattern = r"Colour\s*:\s*(.+?)(?=\n)"

# Open the PDF file
pdf_file = fitz.open(pdf_path)
summary_data = []

# Extract product details per page
for page_num in range(0, 3):
    page = pdf_file[page_num]
    text = page.get_text("text")  # Extract text from the page
    
    print(f"\nProcessing page {page_num + 1}...")  # Debug line
    
    # Perform regex searches for each detail
    product_name = re.search(product_name_pattern, text, re.MULTILINE)
    highlight = re.search(highlights_pattern, text, re.DOTALL)
    feature = re.search(features_pattern, text, re.DOTALL)
    mrp = re.search(mrp_pattern, text)
    material = re.search(material_pattern, text)
    color = re.search(color_pattern, text)
    dimensions = re.search(dimensions_pattern, text)
    
    # Print extracted information for debugging
    print(f"Product Name: {product_name.group(0).strip() if product_name else 'N/A'}")
    print(f"Highlights: {highlight.group(1).strip().replace('\n', ', ') if highlight else 'N/A'}")
    print(f"Features: {feature.group(1).strip().replace('\n', ', ') if feature else 'N/A'}")
    print(f"MRP: {mrp.group(1) if mrp else 'N/A'}")
    print(f"Material: {material.group(1).strip() if material else 'N/A'}")
    print(f"Dimensions: {dimensions}")
    print(f"Color: {color.group(1).strip() if color else 'N/A'}")

    # Handle the case where dimensions might be None
    if dimensions:
        # If dimensions are found, use them
        dimensions_values = [" x ".join(dim.strip() for dim in dimensions.groups())]
    else:
        # If dimensions are not found, set to "N/A"
        dimensions_values = ["N/A"]
    
    # Append data to list, using "N/A" where necessary
    summary_data.append({
        "Page Number": page_num + 1,
        "Product Name": product_name.group(0).strip() if product_name else "N/A",
        "Highlights": highlight.group(1).strip().replace("\n", ", ") if highlight else "N/A",
        "Features": feature.group(1).strip().replace("\n", ", ") if feature else "N/A",
        "MRP": mrp.group(1) if mrp else "N/A",
        "Material": material.group(1).strip() if material else "N/A",
        "Dimensions": dimensions_values[0],  # Use the handled dimensions
        "Color": color.group(1).strip() if color else "N/A"
    })

# Convert to DataFrame and save to Excel
df = pd.DataFrame(summary_data)
output_path = r'C:\selenium2\_Product_Details_With_Summary.xlsx'
df.to_excel(output_path, index=False)

print(f"\nData saved to Excel: {output_path}")
