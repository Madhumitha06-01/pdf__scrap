import re
import PyPDF2
import pandas as pd

# Path to the PDF file
file_path = r"Healthshine Catalogue - Copy.pdf"  # Use raw string
output_excel_path = r"C:\selenium2\summary1_output.xlsx"  # Path for the output Excel file

# Initialize a list to store summaries with their page indices
summary_data = []

# Read text from the PDF
with open(file_path, "rb") as file:
    reader = PyPDF2.PdfReader(file)
    for page_index, page in enumerate(reader.pages):
        text = page.extract_text()  # Extract text from the current page

        # Define the summary pattern
        summary_pattern = r"(Healthshine's.+?)(?=\s*\n\s*Highlights:)"

        # Search for the summary in the extracted text
        match = re.search(summary_pattern, text, re.DOTALL)  # re.DOTALL allows '.' to match newline characters
        if match:
            summary = match.group(1).strip()  # Use .strip() to clean up whitespace
            print(f"Summary found on page {page_index + 1}: {summary}")  # Print with page index

            # Append the summary and page index to the list
            summary_data.append({'Page Index': page_index + 1, 'Summary': summary})

# Check if any summaries were found and save to Excel
if summary_data:
    df = pd.DataFrame(summary_data)  # Create a DataFrame from the list
    df.to_excel(output_excel_path, index=False)  # Save DataFrame to Excel
    print(f"Summaries saved to {output_excel_path}")
else:
    print("No summaries found.")
