import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="Excel Transaction Grouper", layout="wide")

# --- Core processing functions adapted for Streamlit ---

def group_transactions_by_narration_suffix(df):
    """
    Group transactions by last 3 letters of narration, adapted for Streamlit.
    """
    grouped_data = defaultdict(lambda: {'withdrawals': [], 'deposits': [], 'transactions': []})
    
    # Find the narration, withdrawal, and deposit columns
    narration_col, withdrawal_col, deposit_col = None, None, None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'narration' in col_lower or 'description' in col_lower or 'particulars' in col_lower:
            narration_col = col
        elif 'withdrawal' in col_lower or 'debit' in col_lower:
            withdrawal_col = col
        elif 'deposit' in col_lower or 'credit' in col_lower:
            deposit_col = col
            
    st.info(f"Identified Columns: Narration='{narration_col}', Withdrawal='{withdrawal_col}', Deposit='{deposit_col}'")

    if not narration_col:
        st.error("Could not automatically find the 'Narration' column. Please ensure your Excel file has a column with a name like 'Narration', 'Description', or 'Particulars'.")
        return None

    # This is the first column, which the user's logic expects to be the 'Date' column
    # We get its name dynamically to avoid errors if it's not exactly 'Date'
    date_column_name = df.columns[0]

    # Process each transaction
    for index, row in df.iterrows():
        try:
            # User-added logic to skip summary rows at the end of the statement
            date_val = str(row[date_column_name])
            if '****' in date_val or "STATEMENT SUMMARY" in date_val:
                continue
        except KeyError:
            # Handle case where the column doesn't exist in a row, though unlikely with dataframes
            continue

        narration = str(row[narration_col]) if pd.notna(row[narration_col]) else ""
        
        if len(narration) >= 3:
            suffix = narration[-3:].upper()
            
            withdrawal, deposit = 0.0, 0.0
            
            if withdrawal_col and pd.notna(row[withdrawal_col]):
                try:
                    withdrawal = float(str(row[withdrawal_col]).replace(',', ''))
                except (ValueError, TypeError):
                    pass # Keep as 0 if conversion fails
            
            if deposit_col and pd.notna(row[deposit_col]):
                try:
                    deposit = float(str(row[deposit_col]).replace(',', ''))
                except (ValueError, TypeError):
                    pass # Keep as 0
            
            if withdrawal > 0:
                grouped_data[suffix]['withdrawals'].append(withdrawal)
            if deposit > 0:
                grouped_data[suffix]['deposits'].append(deposit)
    
    return grouped_data

def create_excel_output_bytes(grouped_data):
    """
    Creates the Excel file in memory and returns it as bytes.
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
    
    df_summary = pd.DataFrame(excel_data).sort_values('Narration_Suffix').reset_index(drop=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Grouped_Transactions', index=False)
        worksheet = writer.sheets['Grouped_Transactions']
        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column if cell.value)
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
    return output.getvalue(), df_summary

# --- Streamlit App UI ---

st.title("📂 Excel Account Statement Grouper")
st.write("Upload your account statement in Excel format. The app will group transactions by the last 3 letters of the narration and generate a summary file for you to download.")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' uploaded successfully!")

    # Using skiprows=20 as specified by the user's last change
    try:
        df = pd.read_excel(uploaded_file, skiprows=20)
        
        st.write("### Data Preview (first 5 transaction rows)")
        st.dataframe(df.head())

        if st.button("Process Transactions", type="primary"):
            with st.spinner("Analyzing and grouping transactions..."):
                grouped_data = group_transactions_by_narration_suffix(df)

                if grouped_data:
                    excel_bytes, summary_df = create_excel_output_bytes(grouped_data)
                    
                    st.write("### Grouped Transactions Summary")
                    st.dataframe(summary_df)
                    
                    st.download_button(
                        label="📥 Download Processed Excel File",
                        data=excel_bytes,
                        file_name=f"Grouped_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Could not group transactions. Please check the file format and ensure the columns are named correctly.")

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        st.exception(e) 