import pandas as pd
import os

print("Starting the merging and cleaning process...")

# 1. Setup Rules
EXCLUSIONS = ['Thrive', 'Kalinga']
OUTPUT_FILE = 'Combined_Master_Tracker.xlsx'

STATUS_MAP = {
    'Done': 'Completed',
    'Yes': 'Completed',
    'Uploaded': 'Completed',
    'Active': 'Completed',
    'No': 'Pending',
    'Not Started': 'Pending',
    'Nan': 'Pending',
    'None': 'Pending',
    'Ongoing': 'In Progress',
    'In Progress': 'In Progress'
}

def clean_text(text):
    """Removes hidden spaces and capitalizes the first letter of each word."""
    if pd.isna(text):
        return 'Pending'
    cleaned = str(text).strip().title()
    # Map the cleaned text to our standard terms if it exists in the dictionary
    return STATUS_MAP.get(cleaned, cleaned)

# 2. Find all Excel files
all_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$') and f != OUTPUT_FILE]

if not all_files:
    print("No Excel files found in this folder!")
    exit()

print(f"Found {len(all_files)} files to merge: {', '.join(all_files)}")

# 3. Combine and Clean Data
combined_data = {}

for file in all_files:
    try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name)
            
            # Remove excluded schools to keep lists focused
            if 'School Name' in df.columns:
                df = df[~df['School Name'].isin(EXCLUSIONS)]
            
            # Apply the auto-cleaner to all text columns
            for col in df.columns:
                if df[col].dtype == 'object':
                    # Apply cleaning only to columns that look like they hold statuses
                    if 'status' in col.lower() or col == 'Confirmation':
                        df[col] = df[col].apply(clean_text)
                    else:
                        # Just strip spaces for names and regular text
                        df[col] = df[col].astype(str).str.strip()
                        # Remove 'nan' strings that pandas sometimes creates
                        df[col] = df[col].replace({'nan': '', 'Nan': ''})
            
            # Add to our master dictionary
            if sheet_name not in combined_data:
                combined_data[sheet_name] = []
            combined_data[sheet_name].append(df)
            
    except Exception as e:
        print(f"Error reading {file}: {e}")

# 4. Save to a single Master File
print(f"Merging data and saving to {OUTPUT_FILE}...")

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    for sheet_name, df_list in combined_data.items():
        if df_list:
            # Stick all rows for this sheet together
            final_df = pd.concat(df_list, ignore_index=True)
            # Save it as a tab in the new Excel file
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("✅ Success! Your clean, merged file is ready.")