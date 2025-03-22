import pandas as pd
import json
import os

try:
    # Read Excel File
    df = pd.read_excel('Copy of act.xlsb(1).xlsx', sheet_name='N C')

    # Debug: Print column names to check actual names
    print("🔍 Column Names:", df.columns.tolist())

    # Strip spaces in column names
    df.columns = df.columns.str.strip()

    # Forward fill only if columns exist
    for col in ['SL.NO', 'Act Name', 'Act Number']:
        if col in df.columns:
            df[col] = df[col].ffill()
        else:
            print(f"⚠️ Warning: Column '{col}' not found!")

    # Drop completely empty columns
    df = df.dropna(axis=1, how='all')

    # Remove columns starting with 'Unnamed'
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]

    # Convert to list of dicts, removing NaN
    data = []
    for record in df.to_dict(orient='records'):
        clean_record = {k: v for k, v in record.items() if pd.notna(v)}
        data.append(clean_record)

    # Check if output.json exists → delete to avoid permission issue
    if os.path.exists('output.json'):
        os.remove('output.json')

    # Save as JSON
    with open('output.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print("✅ Successfully cleaned! Null columns removed!")

except Exception as e:
    print(f"❌ Error occurred: {e}")
