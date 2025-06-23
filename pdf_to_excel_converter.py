import pandas as pd
import pdfplumber
import re
from collections import defaultdict
import os

def extract_transactions_from_pdf(pdf_path):
    """
    Extract transaction data from PDF account statement
    """
    transactions = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    # Split text into lines
                    lines = text.split('\n')
                    
                    for line in lines:
                        # Look for transaction patterns
                        # This pattern may need adjustment based on actual PDF format
                        transaction_pattern = r'^(\d{2}/\d{2}/\d{2})\s+([A-Za-z0-9\-]+)\s+(\d{16})\s+(\d{2}/\d{2}/\d{2})\s+([0-9,]+\.?\d*)\s+([0-9,]+\.?\d*)$'
                        #transaction_pattern = r'^(\d{2}/\d{2}/\d{2})\s+([A-Za-z0-9\-\s\.]+?)\s+(\d{16})\s+(\d{2}/\d{2}/\d{2})\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*|-)?\s+([0-9,]+\.?\d*)$'

                        match = re.search(transaction_pattern, line)

                        if "From" in line:
                            print("test")
                        
                        if match and not "From" in line:
                            date, narration, withdrawal, deposit, balance, _ = match.groups()
                            
                            # Clean up the data
                            withdrawal = withdrawal.replace(',', '') if withdrawal != '-' else '0'
                            deposit = deposit.replace(',', '') if deposit != '-' else '0'
                            
                            transactions.append({
                                'Date': date,
                                'Narration': narration,
                                'Withdrawal': float(withdrawal) if withdrawal != '0' else 0,
                                'Deposit': float(deposit) if deposit != '0' else 0,
                                'Balance': float(balance.replace(',', ''))
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
    
    print("Grouping transactions by narration suffix...")
    grouped_data = group_transactions_by_narration_suffix(transactions)
    
    print("Creating Excel output...")
    result_df = create_excel_output(grouped_data, output_file)
    
    print(f"Excel file created: {output_file}")
    print(f"Total groups: {len(grouped_data)}")
    
    # Display summary
    print("\nSummary:")
    print(result_df.to_string(index=False))

if __name__ == "__main__":
    main() 