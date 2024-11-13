import fitz  # PyMuPDF
import pandas as pd

# Path to the PDF file and output Excel file
file_path = r"C:\selenium2\Healthshine Catalogue - Copy.pdf"
excel_path = r"C:\selenium2\extracted02_images.xlsx"

# Open the PDF file
pdf_file = fitz.open(file_path)
page_nums = len(pdf_file)
image_data = []

# Extract images from each page in the PDF
for page_num in range(page_nums):  # Use page_nums to cover all pages dynamically
    page = pdf_file[page_num]
    images = page.get_images(full=True)
    
    # Initialize a list to store image names for each page
    image_names = []

    # Check if the page has images
    if images:
        for img_index, image in enumerate(images, start=1):
            xref = image[0]
            base_image = pdf_file.extract_image(xref)
            image_ext = base_image['ext']
            image_name = f'image_page_{page_num + 1}_{img_index}.{image_ext}'
            
            # Collect image names for this page
            image_names.append(image_name)
    else:
        # If no images are found on the page, use "N/A"
        image_names.append("N/A")
    
    # Append combined data for the page to the list
    image_data.append({
        'Page Number': page_num + 1,
        'Image Names': image_names,  # Store as a list instead of joining
    })

# Create a DataFrame
df = pd.DataFrame(image_data)

# Save the DataFrame to an Excel file, converting lists to strings for Excel compatibility
df['Image Names'] = df['Image Names'].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)

df.to_excel(excel_path, index=False)

print(f'Image data saved to {excel_path}')
