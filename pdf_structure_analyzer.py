import pdfplumber
import re

def analyze_pdf_structure(pdf_path):
    """
    Analyze the PDF structure to understand the exact format of transaction lines
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"PDF has {len(pdf.pages)} pages")
            
            for page_num, page in enumerate(pdf.pages):
                print(f"\n=== Page {page_num + 1} ===")
                text = page.extract_text()
                
                if text:
                    lines = text.split('\n')
                    print(f"Total lines: {len(lines)}")
                    
                    # Show all lines that contain dates (potential transaction lines)
                    print("\nLines containing dates (potential transactions):")
                    for i, line in enumerate(lines):
                        if re.search(r'\d{2}/\d{2}/\d{4}', line):
                            print(f"Line {i+1}: '{line}'")
                    
                    # Show lines that might be headers
                    print("\nLines that might be headers:")
                    for i, line in enumerate(lines):
                        if any(keyword in line.lower() for keyword in ['from :', 'to :', 'statement', 'account', 'balance brought forward', 'opening balance']):
                            print(f"Line {i+1}: '{line}'")
                    
                    # Show lines with specific patterns that might be transactions
                    print("\nLines with transaction-like patterns:")
                    for i, line in enumerate(lines):
                        # Look for lines with date + text + numbers pattern
                        if re.search(r'\d{2}/\d{2}/\d{4}.*\d+', line):
                            print(f"Line {i+1}: '{line}'")
                            
    except Exception as e:
        print(f"Error analyzing PDF: {e}")

if __name__ == "__main__":
    analyze_pdf_structure("Acct Statement_XX1020_19062025.pdf") 