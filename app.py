import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO
import json

st.set_page_config(page_title="Excel Transaction Grouper", layout="wide")

# --- Core processing functions adapted for Streamlit ---

abbreviation = "{\"TIF Rent\":{\"Description\":\"Tiffin\",\"Category\":\"Tiffin\"},\"Ext LB\":{\"Description\":\"External Labour\",\"Category\":\"External Labour\"},\"Petrol\":{\"Description\":\"Petrol\",\"Category\":\"Transport\"},\"ptr\":{\"Description\":\"Petrol\",\"Category\":\"Transport\"},\"Tif Ptr\":{\"Description\":\"Tiffin\",\"Category\":\"Tiffin\"},\"Adv\":{\"Description\":\"Pinu\",\"Category\":\"Transport\"},\"Pinu\":{\"Description\":\"Pinu\",\"Category\":\"Transport\"},\"Bike\":{\"Description\":\"Bike\",\"Category\":\"Transport\"},\"Bharat\":{\"Description\":\"Bharat\",\"Category\":\"Bharat\"},\"Weed ptr\":{\"Description\":\"Weed Petrol\",\"Category\":\"Weed\"},\"Weed\":{\"Description\":\"Weed\",\"Category\":\"Weed\"},\"wd\":{\"Description\":\"Weed\",\"Category\":\"Weed\"},\"Tif\":{\"Description\":\"Tiffin\",\"Category\":\"Tiffin\"},\"Gas\":{\"Description\":\"Gas\",\"Category\":\"Transport\"},\"Plants\":{\"Description\":\"Plants\",\"Category\":\"Plants\"},\"Seeds\":{\"Description\":\"Seeds\",\"Category\":\"Seeds\"},\"Help\":{\"Description\":\"Helper\",\"Category\":\"Helper\"},\"Helper\":{\"Description\":\"Helper\",\"Category\":\"Helper\"},\"Nanu\":{\"Description\":\"Nanu\",\"Category\":\"Nanu\"},\"Suresh\":{\"Description\":\"Suresh\",\"Category\":\"Suresh\"},\"Jeev\":{\"Description\":\"Jeevamrut\",\"Category\":\"Fertilizer\"},\"Tempo Ptr\":{\"Description\":\"Tempo Petrol\",\"Category\":\"Transport\"}}"

abbreviation_map = json.loads(abbreviation)

def group_transactions_by_narration_suffix(df):
    """
    Group transactions by last 3 letters of narration, adapted for Streamlit.
    """
    grouped_data = defaultdict(lambda: {'withdrawals': [], 'deposits': [], 'transactions': []})
    
    # Find the narration, withdrawal, and deposit columns
    narration_col, withdrawal_col, deposit_col = None, None, None
    
    for col in df.columns:
        col_lower = str(col).lower()
        print(col_lower)
        if 'narration' in col_lower or 'description' in col_lower or 'particulars' in col_lower:
            narration_col = col
        elif 'withdrawal' in col_lower or 'debit' in col_lower:
            withdrawal_col = col
        elif 'deposit' in col_lower or 'credit' in col_lower:
            deposit_col = col
        elif 'date' in col_lower:
            date_col = col
            
    st.info(f"Identified Columns: Narration='{narration_col}', Withdrawal='{withdrawal_col}', Deposit='{deposit_col}', Date='{date_col}")

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
            if '****' in date_val:
                continue
            if "STATEMENT SUMMARY" in date_val:
                break
        except KeyError:
            # Handle case where the column doesn't exist in a row, though unlikely with dataframes
            continue

        narration = str(row[narration_col]) if pd.notna(row[narration_col]) else ""
        
        matched_key = None
        for key in abbreviation_map.keys():
            if key.lower() in narration.lower():
                matched_key = key
                break

        if matched_key:
            group_key = matched_key
        else:
            group_key = "Other"
        
        withdrawal, deposit = 0.0, 0.0
        date = 'NA'
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
        
        if date_col and pd.notna(row[date_col]):
            try:
                date = str(row[date_col])
            except (ValueError, TypeError):
                pass # Keep as 0

        narration_key = narration.replace('-', '_')
        data_key = f'{narration_key} - {date} - {group_key}'
        if withdrawal > 0:
            grouped_data[data_key]['withdrawals'] = withdrawal

    return grouped_data

def create_excel_output_bytes(grouped_data):
    """
    Creates the Excel file in memory and returns it as bytes.
    """
    excel_data = []
    for data_key, data in grouped_data.items():
        total_withdrawal = data['withdrawals']
        group_key = data_key.split('-')[-1].strip()

        if group_key in abbreviation_map:
            description = abbreviation_map[group_key]['Description']
            category = abbreviation_map[group_key]['Category']
        else:
            description = 'NA'
            category = 'NA'
        
        excel_data.append({
            'Date': data_key.split('-')[1].strip(),
            'Narration': data_key.split('-')[0].strip(),
            'Tag': data_key.split('-')[-1].strip(),
            'Description': description,
            "Category": category,
            'Total_Withdrawal': total_withdrawal
        })
    
    df_summary = pd.DataFrame(excel_data).sort_values('Date').reset_index(drop=True)
    
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

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

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