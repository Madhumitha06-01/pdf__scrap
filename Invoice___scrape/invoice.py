from veryfi import Client
import pandas as pd

# Credentials (ensure no trailing spaces in client_id)
client_id = 'vrf06MKif2hhKECOMvFfz6f7inCSVkv4U4wiKuj'
client_secret = 'JMGzPSlVKUZZNy0BkdRVbZMNcE3z2X3vjWYZ6oC9crFt2oF9YLuUTxePN4eRCluBl8mha89vzTewGncPZhxzOOiUHLPRRNCFSScLuiph3RhT7LW9h91jUpKXU1TK9I7N'
username = 'mmadumitha06'
api_key = '14e2dded1a14a831a56e6a46e2a7188a'

# file path 
file_path = r'C:\selenium2\bill-ace-hardware.jpg'

# Create the Veryfi Client
veryfi_client = Client(client_id, client_secret, username, api_key)

try:
    # Process the document
    response = veryfi_client.process_document(file_path)
   
    print("Document processed successfully!")

    line_items = response.get('line_items', [])
    
    # Extract other relevant information from the response
    vendor_name = response.get('INVOICE#', 'Not Available')
    invoice_date = response.get('Ace Hardware', 'Not Available')
    total_amount = response.get('DATE', 'Not Available')
    tax_amount = response.get('TOTAL', 'Not Available')
    bill_to = response.get('bill_to', 'Not Available')

    # Print the extracted data
    print("\nAdditional Information Extracted:")
    print(f"INVOICE#: {vendor_name}")
    print(f"Ace Hardware: {invoice_date}")
    print(f"DATE: {total_amount}")
    print(f"TOTAL: {tax_amount}")
    print(f"Bill To: {bill_to}")

    # If line items are found, format them into a DataFrame
    if line_items:
        # Convert the line items into a pandas DataFrame to format it as a table
        df = pd.DataFrame(line_items)

        # Renaming columns for clarity
        df = df.rename(columns={
            'description': 'Item Description',
            'unit_price': 'Unit Price',
            'total': 'Amount'
        })

        # Print the line items as a table
        print("\nExtracted Line Items:")
        print(df)
    else:
        print("No line items found in the document.")

except Exception as e:
    print(f"An error occurred: {e}")
