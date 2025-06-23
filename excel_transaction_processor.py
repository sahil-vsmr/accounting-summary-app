import pandas as pd
import numpy as np
from collections import defaultdict
import os


abbreviation_map = {
  "TIF Rent": { "Description": "Tiffin", "Category": "Tiffin" },
  "Ext LB": { "Description": "External Labour", "Category": "External Labour" },
  "Petrol": { "Description": "Petrol", "Category": "Transport" },
  "ptr": { "Description": "Petrol", "Category": "Transport" },
  "Tif Ptr": { "Description": "Tiffin", "Category": "Tiffin" },
  "Adv": { "Description": "Pinu", "Category": "Transport" },
  "Pinu": { "Description": "Pinu", "Category": "Transport" },
  "Bike": { "Description": "Bike", "Category": "Transport" },
  "Bharat": { "Description": "Bharat", "Category": "Bharat" },
  "Weed ptr": { "Description": "Weed Petrol", "Category": "Weed" },
  "Weed": { "Description": "Weed", "Category": "Weed" },
  "wd": { "Description": "Weed", "Category": "Weed" },
  "Tif": { "Description": "Tiffin", "Category": "Tiffin" },
  "Gas": { "Description": "Gas", "Category": "Transport" },
  "Plants": { "Description": "Plants", "Category": "Plants" },
  "Seeds": { "Description": "Seeds", "Category": "Seeds" },
  "Help": { "Description": "Helper", "Category": "Helper" },
  "Helper": { "Description": "Helper", "Category": "Helper" },
  "Nanu": { "Description": "Nanu", "Category": "Nanu" },
  "Suresh": { "Description": "Suresh", "Category": "Suresh" },
  "Jeev": { "Description": "Jeevamrut", "Category": "Fertilizer" },
  "Tempo Ptr": { "Description": "Tempo Petrol", "Category": "Transport" }
}



def read_excel_statement(excel_path):
    """
    Read the Excel account statement and extract transaction data
    Skip first 22 rows and start from row 23
    """
    try:
        # Read the Excel file, skipping first 22 rows (header rows)
        df = pd.read_excel(excel_path, skiprows=20)
        
        print(f"Excel file loaded successfully")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        # Display first few rows to understand structure
        print("\nFirst 5 rows (starting from row 20):")
        print(df.head())
        
        return df
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def clean_and_process_transactions(df):
    """
    Clean and process the transaction data
    """
    # Make a copy to avoid modifying original
    df_clean = df.copy()
    
    # Remove any completely empty rows
    df_clean = df_clean.dropna(how='all')
    
    # Reset index
    df_clean = df_clean.reset_index(drop=True)
    
    print(f"\nCleaned data shape: {df_clean.shape}")
    
    return df_clean

def group_transactions_by_narration_suffix(df):
    """
    Group transactions by last 3 letters of narration
    """
    grouped_data = defaultdict(lambda: {'withdrawals': [], 'deposits': [], 'transactions': []})
    
    # Find the narration column (it might be named differently)
    narration_col = None
    withdrawal_col = None
    deposit_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if any(narration in col_lower for narration in ['narration', 'description', 'particulars', 'details']):
            narration_col = col
        elif any(withdrawal in col_lower for withdrawal in ['withdrawal', 'debit']):
            withdrawal_col = col
        elif any(deposit in col_lower for deposit in ['deposit', 'credit']):
            deposit_col = col
    
    print(f"\nIdentified columns:")
    print(f"Narration column: {narration_col}")
    print(f"Withdrawal column: {withdrawal_col}")
    print(f"Deposit column: {deposit_col}")
    
    if not narration_col:
        print("Could not find narration column. Available columns:")
        for i, col in enumerate(df.columns):
            print(f"{i+1}. {col}")
        return None
    
    # Process each transaction
    for index, row in df.iterrows():
        narration = str(row[narration_col]) if pd.notna(row[narration_col]) else ""
        
        if '****' in row['Date']:
            continue

        if "STATEMENT SUMMARY  :-" in row["Date"]:
            break

        matched_key = None
        for key in abbreviation_map.keys():
            if key.lower() in narration.lower():
                matched_key = key
                break

        print(f'Matched Key {matched_key}!!!!!!!!!!!!!')
        if matched_key:
            group_key = matched_key
            group_value = abbreviation_map.values(group_key)
        else:
            group_key = "Other"
            group_value = { "Description": "NA", "Category": "NA" }

            
            # Get withdrawal and deposit amounts
        withdrawal = 0
        deposit = 0
        
        if withdrawal_col and pd.notna(row[withdrawal_col]):
            try:
                withdrawal = float(str(row[withdrawal_col]).replace(',', ''))
            except:
                withdrawal = 0
        
        if deposit_col and pd.notna(row[deposit_col]):
            try:
                deposit = float(str(row[deposit_col]).replace(',', ''))
            except:
                deposit = 0
        
        # Add to grouped data
        if withdrawal > 0:
            grouped_data[group_key]['withdrawals'].append(withdrawal)
        if deposit > 0:
            grouped_data[group_key]['deposits'].append(deposit)
        
        # Store full transaction details
        transaction_info = {
            'Index': index,
            'Narration': narration,
            'Withdrawal': withdrawal,
            'Deposit': deposit,
            'group_key': group_key,
            'group_description': group_value.get("Description"),
            'group_category': group_value.get("Category")
        }
        grouped_data[group_key]['transactions'].append(transaction_info)
    
    return grouped_data

def create_excel_output(grouped_data, output_path):
    """
    Create Excel file with grouped transaction data
    """
    excel_data = []
    
    for suffix, data in grouped_data.items():
        total_withdrawal = sum(data['withdrawals'])
        total_deposit = sum(data['deposits'])
        
        excel_data.append({
            'Narration_Suffix': suffix,
            'Total_Withdrawal': total_withdrawal,
            'Total_Deposit': total_deposit,
            'Net_Amount': total_deposit - total_withdrawal,
            'Transaction_Count': len(data['withdrawals']) + len(data['deposits'])
        })
    
    # Create DataFrame and sort by suffix
    df = pd.DataFrame(excel_data)
    df = df.sort_values('Narration_Suffix')
    
    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Grouped_Transactions', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Grouped_Transactions']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return df

def main():
    excel_file = "Acct Statement_XX1020_23062025.xlsx"
    output_file = "Grouped_Transactions_From_Excel.xlsx"
    
    print("Reading Excel account statement...")
    df = read_excel_statement(excel_file)
    
    if df is None:
        print("Failed to read Excel file.")
        return
    
    print("\nCleaning and processing transactions...")
    df_clean = clean_and_process_transactions(df)
    
    print("\nGrouping transactions by narration suffix...")
    grouped_data = group_transactions_by_narration_suffix(df_clean)
    
    if grouped_data is None:
        print("Failed to group transactions.")
        return
    
    print(f"\nFound {len(grouped_data)} unique narration suffixes")
    
    # Show some examples of grouped data
    print("\nSample grouped data:")
    for i, (suffix, data) in enumerate(list(grouped_data.items())[:5]):
        print(f"{i+1}. Suffix: '{suffix}' - {len(data['transactions'])} transactions")
        print(f"   Withdrawals: {len(data['withdrawals'])} (Total: {sum(data['withdrawals']):.2f})")
        print(f"   Deposits: {len(data['deposits'])} (Total: {sum(data['deposits']):.2f})")
    
    print("\nCreating Excel output...")
    result_df = create_excel_output(grouped_data, output_file)
    
    print(f"Excel file created: {output_file}")
    print(f"Total groups: {len(grouped_data)}")
    
    # Display summary
    print("\nSummary:")
    print(result_df.to_string(index=False))
    
    # Show some example transactions for verification
    print("\nExample transactions by suffix:")
    for suffix, data in list(grouped_data.items())[:3]:
        print(f"\nSuffix: '{suffix}'")
        for transaction in data['transactions'][:2]:  # Show first 2 transactions
            print(f"  - {transaction['Narration']} (W: {transaction['Withdrawal']}, D: {transaction['Deposit']})")

if __name__ == "__main__":
    main() 