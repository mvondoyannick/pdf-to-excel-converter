import tabula
import pandas as pd
import os

def pdf_to_excel(pdf_file_path, excel_file_path=None):
    if not os.path.exists(pdf_file_path):
        print(f"The file {pdf_file_path} does not exist.")
        return
    # If no excel_file_path provided, use the same base name as the PDF
    if excel_file_path is None:
        base = os.path.splitext(pdf_file_path)[0]
        excel_file_path = base + '.xls'
    # Read PDF tables into a list of DataFrames
    # 'pages="all"' extracts tables from all pages
    # 'multiple_tables=True' handles multiple tables per page
    
    try:
        tables = tabula.read_pdf(pdf_file_path, pages="all", multiple_tables=True, stream=True)
        # Try writing as .xls (older Excel format). If that fails, fall back to .xlsx.
        try:
            with pd.ExcelWriter(excel_file_path) as writer:
                for i, df in enumerate(tables):
                    df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
        except Exception as write_exc:
            # If .xls failed and we were using .xls, try .xlsx as fallback
            if excel_file_path.lower().endswith('.xls'):
                fallback_path = os.path.splitext(excel_file_path)[0] + '.xlsx'
                try:
                    with pd.ExcelWriter(fallback_path) as writer:
                        for i, df in enumerate(tables):
                            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
                    excel_file_path = fallback_path
                except Exception as fallback_exc:
                    print(f"Failed to write both .xls and .xlsx: {fallback_exc}")
                    return
            else:
                print(f"Failed to write Excel file: {write_exc}")
                return

        print(f"Successfully converted {pdf_file_path} to {excel_file_path}")

    except Exception as e:
        print(f"An error occurred while reading the PDF: {e}")
        return

if __name__ == '__main__':
    # Example usage when running this module directly
    pdf_file = 'KRIBI.pdf'
    # Call without excel_file to produce 'KRIBI.xls' (or fallback to .xlsx if needed)
    pdf_to_excel(pdf_file)
