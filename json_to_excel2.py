import json
import pandas as pd

# Load JSON data
with open('out.json', 'r') as f:
    data = json.load(f)

# Convert to DataFrame
df = pd.DataFrame(data)

# Create a new DataFrame for extracted data
extracted_data = pd.DataFrame()

# Extract batch_id from posting_instruction_batch
if 'posting_instruction_batch' in df.columns:
    # Extract batch_id
    extracted_data['batch_id'] = df['posting_instruction_batch'].apply(
        lambda x: x.get('id') if isinstance(x, dict) else None
    )
    
    # Function to extract posting data from the correct location
    def extract_posting_fields(batch, field_name):
        """Extract fields from the correct location in the JSON structure"""
        if not isinstance(batch, dict):
            return None
            
        # Get posting instructions
        instructions = batch.get('posting_instructions', [])
        if not instructions or not isinstance(instructions, list) or len(instructions) == 0:
            return None
            
        # Get the first instruction
        instruction = instructions[0]
        if not isinstance(instruction, dict):
            return None
            
        # Try committed_postings first (they seem more reliable)
        committed = instruction.get('committed_postings', [])
        if committed and isinstance(committed, list) and len(committed) > 0:
            # Find the first credit posting (or just use first if no credit)
            for posting in committed:
                if isinstance(posting, dict) and posting.get('credit') is True:
                    if field_name in posting:
                        return posting.get(field_name)
            
            # If no credit posting found, use the first one
            if isinstance(committed[0], dict) and field_name in committed[0]:
                return committed[0].get(field_name)
        
        # Try custom_instruction.postings if committed_postings didn't work
        custom = instruction.get('custom_instruction', {})
        postings = custom.get('postings', []) if isinstance(custom, dict) else []
        
        if postings and isinstance(postings, list) and len(postings) > 0:
            # Find the first credit posting (or just use first if no credit)
            for posting in postings:
                if isinstance(posting, dict) and posting.get('credit') is True:
                    if field_name in posting:
                        return posting.get(field_name)
            
            # If no credit posting found, use the first one
            if isinstance(postings[0], dict) and field_name in postings[0]:
                return postings[0].get(field_name)
        
        return None

    # Extract each field we need
    fields_to_extract = {
        'credit_amount': 'amount',
        'denomination': 'denomination',
        'account_id': 'account_id',
        'account_address': 'account_address',
        'asset': 'asset',
        'phase': 'phase',
        'internal_account_processing_label': 'internal_account_processing_label',
        'posting_instruction_id': 'id'
    }
    
    # Extract each field
    for column_name, field_name in fields_to_extract.items():
        if column_name == 'posting_instruction_id':
            # Special handling for instruction ID which is in a different place
            extracted_data[column_name] = df['posting_instruction_batch'].apply(
                lambda x: x.get('posting_instructions', [{}])[0].get('id') if isinstance(x, dict) and 
                          'posting_instructions' in x and len(x['posting_instructions']) > 0 else None
            )
        else:
            # Regular field extraction
            extracted_data[column_name] = df['posting_instruction_batch'].apply(
                lambda x: extract_posting_fields(x, field_name)
            )

# Extract timestamps
if 'timestamp' in df.columns:
    extracted_data['value_timestamp'] = df['timestamp']
    extracted_data['booking_timestamp'] = df['timestamp']

# Print column information
print("Extracted columns:", extracted_data.columns.tolist())
print("Number of records:", len(extracted_data))

# Export to Excel with formatting
with pd.ExcelWriter('transaction_data_fixed.xlsx', engine='openpyxl') as writer:
    extracted_data.to_excel(writer, index=False, sheet_name='Transactions')
    # Auto-adjust column widths
    for column in extracted_data:
        column_width = max(extracted_data[column].astype(str).map(len).max(), len(column)) + 2
        col_idx = extracted_data.columns.get_loc(column)
        writer.sheets['Transactions'].column_dimensions[chr(65 + col_idx)].width = column_width

print("Excel file created successfully!")