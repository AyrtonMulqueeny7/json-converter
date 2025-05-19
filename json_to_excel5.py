#Importing Necessary Libraries
import json
import pandas as pd
import datetime
from datetime import timezone, timedelta  # ADDED: For timezone handling

#new functions needed as Account type is not in the same place as the other fields

# Load JSON data
with open('out.json', 'r') as f: #r is read mode
    data = json.load(f)

# Convert to DataFrame
df = pd.DataFrame(data) # like a temporary db table in memory


extracted_data = pd.DataFrame() #empty DataFrame to store extracted data

# Extract batch_id from posting_instruction_batch
if 'posting_instruction_batch' in df.columns:
    # Extract batch_id
    extracted_data['batch_id'] = df['posting_instruction_batch'].apply(
        lambda x: x.get('id') if isinstance(x, dict) else None
    )#checks if there is a row & lambda processes each row in the column
    
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
            
        # Try committed_postings may be a more reliable methid for some entries
        committed = instruction.get('committed_postings', [])
        if committed and isinstance(committed, list) and len(committed) > 0:
            # Use the first posting entry
            if isinstance(committed[0], dict) and field_name in committed[0]:
                return committed[0].get(field_name)
        
    
        custom = instruction.get('custom_instruction', {})
        postings = custom.get('postings', []) if isinstance(custom, dict) else []
        # some entries have no value for custom_instruction as some entries are null
        if postings and isinstance(postings, list) and len(postings) > 0:
            # Use the first posting entry
            if isinstance(postings[0], dict) and field_name in postings[0]:
                return postings[0].get(field_name)
        
        return None

    # ADDED NEW FUNCTION: Function to extract from instruction_details
    def extract_instruction_detail(batch, detail_key):
        """Extract a specific key from instruction_details array in order to get account_type"""
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
        
        # Get instruction details
        details = instruction.get('instruction_details', [])
        if not details or not isinstance(details, list):
            return None
        
        # Find the detail with matching key
        for detail in details:
            if isinstance(detail, dict) and detail.get('key') == detail_key:
                return detail.get('value')
        
        return None

    # Extract each field we need - with credit and amount as separate fields? fix
    fields_to_extract = {
        'credit': 'credit',
        'amount': 'amount',
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
            # Special handling for instruction ID which is in a different place/ column
            extracted_data[column_name] = df['posting_instruction_batch'].apply(
                lambda x: x.get('posting_instructions', [{}])[0].get('id') if isinstance(x, dict) and 
                          'posting_instructions' in x and len(x['posting_instructions']) > 0 else None
            )
        else:
            # Regular field extraction
            extracted_data[column_name] = df['posting_instruction_batch'].apply(
                lambda x: extract_posting_fields(x, field_name)
            )
    
    # ADDED NEW CODE: Extract account_type from instruction_details
    extracted_data['account_type'] = df['posting_instruction_batch'].apply(
        lambda x: extract_instruction_detail(x, 'account_type')
    )

# Extract timestamps - MODIFIED to add correctly formatted timestamps with timezone
if 'timestamp' in df.columns:
    # Keep original timestamps
    extracted_data['value_timestamp'] = df['timestamp']
    extracted_data['booking_timestamp'] = df['timestamp']
    
    # Define the timezone offset (+0200)
    tz_offset = timezone(timedelta(hours=2))
    
    # ADDED NEW CODE: Add readable timestamp columns with correct timezone
    extracted_data['readable_value_date'] = df['timestamp'].apply(
        lambda x: datetime.datetime.fromtimestamp(int(x)/1000, tz=tz_offset).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] + ' +0200' 
        if x and pd.notnull(x) else None
    )
    extracted_data['readable_booking_date'] = df['timestamp'].apply(
        lambda x: datetime.datetime.fromtimestamp(int(x)/1000, tz=tz_offset).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] + ' +0200'
        if x and pd.notnull(x) else None
    )

# MODIFIED: Updated column order to include the new readable date columns
column_order = [
    'batch_id', 'credit', 'amount', 'denomination', 'account_id', 
    'account_type',
    'account_address', 'asset', 'phase', 
    'internal_account_processing_label', 'posting_instruction_id',
    'value_timestamp', 'readable_value_date',  # Added readable date after timestamp
    'booking_timestamp', 'readable_booking_date'  # Added readable date after timestamp
]
#order can be customized above as needed
extracted_data = extracted_data[column_order]

# Print column information
print("Extracted columns:", extracted_data.columns.tolist())
print("Number of records:", len(extracted_data))

# Export to Excel with formatting
with pd.ExcelWriter('transaction_data_fixed_final2.xlsx', engine='openpyxl') as writer:  # Change filename here when needed
    extracted_data.to_excel(writer, index=False, sheet_name='Transactions')
    # Auto-adjust column widths
    for column in extracted_data:
        column_width = max(extracted_data[column].astype(str).map(len).max(), len(column)) + 2
        col_idx = extracted_data.columns.get_loc(column)
        writer.sheets['Transactions'].column_dimensions[chr(65 + col_idx)].width = column_width

print("Excel file created successfully!")