from pypdf import PdfReader
import pandas as pd

# Specify the path to your PDF file
file_path = r"C:\selenium2\Healthshine Catalogue - Copy.pdf"  # Use raw string
start_page = 1  # Starting page number (1-based)
end_page = 138    # Ending page number (1-based)

try:
    # Open the PDF file
    reader = PdfReader(file_path)
    
    # Check if the desired page range is valid
    if start_page < 1 or end_page > len(reader.pages) or start_page > end_page:
        print(f"Invalid page range: {start_page} to {end_page}. The document has {len(reader.pages)} pages.")
    else:
        all_text = []
        
        # Extract text from the specified range of pages
        for page_number in range(start_page, end_page + 1):  # Inclusive of end_page
            page = reader.pages[page_number - 1]  # Adjust for 0-based index
            text = page.extract_text()
            
            if text:  # Ensure text is not None or empty
                all_text.append(text)
            else:
                print(f"No text found on page {page_number}.\n")
        
        # Create a DataFrame with all extracted text
        df = pd.DataFrame(all_text, columns=['Extracted Text'])  # Each page's text in a separate row
        
        # Save DataFrame to Excel
        output_file = f'extracted_data_{start_page}_to_{end_page}.xlsx'
        df.to_excel(output_file, index=False)
        print(f"Text from pages {start_page} to {end_page} saved to {output_file}.")

except FileNotFoundError:
    print(f"The file {file_path} was not found. Please check the path.")
except Exception as e:
    print(f"An error occurred: {e}")