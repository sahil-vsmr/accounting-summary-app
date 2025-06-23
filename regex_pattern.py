import re

def get_transaction_regex_pattern():
    """
    Returns the precise regex pattern for extracting transaction data from the PDF
    """
    # Pattern explanation:
    # ^(\d{2}/\d{2}/\d{2}) - Date in DD/MM/YY format at start of line
    # \s+ - One or more spaces
    # ([A-Za-z0-9\-\s\.]+?) - Narration (non-greedy match)
    # \s+ - One or more spaces  
    # (\d{16}) - 16-digit reference number
    # \s+ - One or more spaces
    # (\d{2}/\d{2}/\d{2}) - Value date
    # \s+ - One or more spaces
    # ([0-9,]+\.?\d*|-)? - Withdrawal amount (optional, can be dash)
    # \s+ - One or more spaces
    # ([0-9,]+\.?\d*|-)? - Deposit amount (optional, can be dash)
    # \s+ - One or more spaces
    # ([0-9,]+\.?\d*) - Closing balance
    
    pattern = r'^(\d{2}/\d{2}/\d{2})\s+([A-Za-z0-9\-\s\.]+?)\s+(\d{16})\s+(\d{2}/\d{2}/\d{2})\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*)$'
    
    return pattern

def test_regex_pattern():
    """
    Test the regex pattern with sample transaction lines
    """
    pattern = get_transaction_regex_pattern()
    
    # Sample transaction lines from the PDF
    test_lines = [
        "03/06/25 UPI-SAISAYAJI 0000105885808379 04/06/25 823.90 18,725.31",
        "04/06/25 UPI-MRSHUBHAM 0000552176269091 04/06/25 2,000.00 20,725.31",
        "04/06/25 JANMAR25INSTAALERTCHG7SMS040425-MIR2 MIR2615429257739 04/06/25 1.66 20,723.65",
        "04/06/25 UPI-PADHYEANAND 0000552142975555 04/06/25 4,000.00 24,723.65",
        "04/06/25 IMPS-515511554137-GOOGLEINDIADIGITAL-UTI 0000515511554137 04/06/25 12,575.00 37,298.65",
        "04/06/25 UPI-PADHYEANAND 0000552139871316 04/06/25 250.00 37,548.65",
        "05/06/25 UPI-SEEMAKEDAR 0000105957141453 05/06/25 9,000.00 32,393.65",
        "05/06/25 UPI-MRGULABDADABHAU 0000105971529423 05/06/25 440.00 31,953.65"
    ]
    
    print("Testing regex pattern with sample transaction lines:")
    print("=" * 80)
    
    for i, line in enumerate(test_lines, 1):
        match = re.search(pattern, line)
        if match:
            date, narration, ref_no, value_date, withdrawal, deposit, balance = match.groups()
            print(f"Line {i}: MATCH")
            print(f"  Date: {date}")
            print(f"  Narration: {narration}")
            print(f"  Ref No: {ref_no}")
            print(f"  Value Date: {value_date}")
            print(f"  Withdrawal: {withdrawal}")
            print(f"  Deposit: {deposit}")
            print(f"  Balance: {balance}")
        else:
            print(f"Line {i}: NO MATCH - {line}")
        print()

def get_improved_converter_script():
    """
    Returns the improved converter script with the correct regex
    """
    script = '''
import pandas as pd
import pdfplumber
import re
from collections import defaultdict

def extract_transactions_from_pdf(pdf_path):
    """
    Extract transaction data from PDF account statement with precise regex
    """
    transactions = []
    
    # Precise regex pattern for transaction lines
    transaction_pattern = r'^(\d{2}/\d{2}/\d{2})\s+([A-Za-z0-9\-\s\.]+?)\s+(\d{16})\s+(\d{2}/\d{2}/\d{2})\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*)$'
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\\n')
                    
                    for line in lines:
                        # Skip header lines
                        if any(keyword in line.lower() for keyword in ['from :', 'to :', 'statement', 'account', 'balance brought forward', 'date narration']):
                            continue
                        
                        match = re.search(transaction_pattern, line.strip())
                        
                        if match:
                            date, narration, ref_no, value_date, withdrawal, deposit, balance = match.groups()
                            
                            # Clean up the data
                            withdrawal = withdrawal.replace(',', '') if withdrawal and withdrawal != '-' else '0'
                            deposit = deposit.replace(',', '') if deposit and deposit != '-' else '0'
                            balance = balance.replace(',', '')
                            
                            transactions.append({
                                'Date': date,
                                'Narration': narration.strip(),
                                'Ref_No': ref_no,
                                'Value_Date': value_date,
                                'Withdrawal': float(withdrawal) if withdrawal != '0' else 0,
                                'Deposit': float(deposit) if deposit != '0' else 0,
                                'Balance': float(balance)
                            })
    
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return []
    
    return transactions

def group_transactions_by_narration_suffix(transactions):
    """
    Group transactions by last 3 letters of narration
    """
    grouped_data = defaultdict(lambda: {'withdrawals': [], 'deposits': []})
    
    for transaction in transactions:
        narration = transaction['Narration']
        if len(narration) >= 3:
            suffix = narration[-3:].upper()
            
            if transaction['Withdrawal'] > 0:
                grouped_data[suffix]['withdrawals'].append(transaction['Withdrawal'])
            if transaction['Deposit'] > 0:
                grouped_data[suffix]['deposits'].append(transaction['Deposit'])
    
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
    pdf_file = "Acct Statement_XX1020_19062025.pdf"
    output_file = "Grouped_Transactions.xlsx"
    
    print("Extracting transactions from PDF...")
    transactions = extract_transactions_from_pdf(pdf_file)
    
    if not transactions:
        print("No transactions found. Please check the PDF format.")
        return
    
    print(f"Found {len(transactions)} transactions")
    
    # Show first few transactions for verification
    print("\\nFirst 5 transactions:")
    for i, t in enumerate(transactions[:5]):
        print(f"{i+1}. {t}")
    
    print("\\nGrouping transactions by narration suffix...")
    grouped_data = group_transactions_by_narration_suffix(transactions)
    
    print("Creating Excel output...")
    result_df = create_excel_output(grouped_data, output_file)
    
    print(f"Excel file created: {output_file}")
    print(f"Total groups: {len(grouped_data)}")
    
    # Display summary
    print("\\nSummary:")
    print(result_df.to_string(index=False))

if __name__ == "__main__":
    main()
'''
    return script

if __name__ == "__main__":
    print("Regex Pattern for Transaction Extraction:")
    print("=" * 50)
    print(get_transaction_regex_pattern())
    print()
    
    test_regex_pattern()
    
    print("\\nImproved Converter Script:")
    print("=" * 50)
    print(get_improved_converter_script()) 