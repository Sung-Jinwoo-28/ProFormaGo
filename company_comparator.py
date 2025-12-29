import sys
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from screener_extractor import ScreenerExtractor
import time
import io



def get_company_data(ticker, years=5):
    print(f"Fetch info for {ticker}...")
    try:
        # Use absolute path to ensure we land in the right place regardless of CWD of runner
        # But relative 'Companies' is fine if running from Scapy root
        companies_dir = os.path.join(os.getcwd(), 'Companies')
        if not os.path.exists(companies_dir):
            os.makedirs(companies_dir)
            
        extractor = ScreenerExtractor(ticker, years, output_folder=companies_dir)
        extractor.run(in_memory=True)
        
        dfs_to_merge = []
        if 'Profit & Loss' in extractor.data: dfs_to_merge.append(extractor.data['Profit & Loss'])
        if 'Balance Sheet' in extractor.data: dfs_to_merge.append(extractor.data['Balance Sheet'])
        if 'Investing Ratios' in extractor.data: dfs_to_merge.append(extractor.data['Investing Ratios'])
        
        if not dfs_to_merge:
            return None
            
        # Merge all metrics into one big list
        # Filter duplicates if any
        master_df = pd.concat(dfs_to_merge, ignore_index=True)
        master_df = master_df.drop_duplicates(subset=['Metric'])
        
        return master_df
        
    except Exception as e:
        print(f"Error extracting {ticker}: {e}")
        return None

def create_comparison_sheet(companies, in_memory=False):
    if len(companies) > 3:
        print("Error: Maximum 3 companies allowed.")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"

    # ... (rest of function logic remains same until saving) ...

    # --- Styles ---
    # Metric Header (Yellow)
    style_metric_header = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    font_bold = Font(bold=True)
    
    # Company Headers
    header_colors = ["00B0F0", "7030A0", "FF0000"] # Blue, Purple, Red
    font_white_bold = Font(color="FFFFFF", bold=True)
    
    align_center = Alignment(horizontal='center', vertical='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # --- Header Setup ---
    # A1: Metric
    c1 = ws.cell(row=1, column=1, value="Metrics")
    c1.fill = style_metric_header
    c1.font = font_bold
    c1.alignment = align_center
    c1.border = thin_border
    
    # Fetch Data
    all_data = [] # List of (Ticker, DF)
    for ticker in companies:
        df = get_company_data(ticker)
        if df is not None:
            all_data.append((ticker, df))
        time.sleep(1)

    if not all_data:
        print("No data fetched.")
        return None

    # Determine Global Metric List (Categorized)
    # Defined lists based on standard Screener format
    categories = {
        'Profit & Loss': [
            'Sales', 'Expenses', 'Operating Profit', 'OPM %', 'Other Income', 
            'Interest', 'Depreciation', 'Profit before tax', 'Tax %', 'Net Profit', 
            'EPS in Rs', 'Dividend Payout %'
        ],
        'Balance Sheet': [
            'Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities', 'Total Liabilities', 
            'Fixed Assets', 'CWIP', 'Investments', 'Other Assets', 'Total Assets'
        ],
        'Financial Ratios': [
            'Net Profit Margin %', 'Operating Profit Margin %', 'Return on Equity (ROE) %', 'ROCE %',
            'Debt to Equity Ratio', 'Interest Coverage Ratio', 'Financial Leverage', 
            'Asset Turnover Ratio', 'Current Ratio'
        ]
    }
    
    current_col = 2
    data_map = [] # List of (df, start_col, valid_years)
    
    for i, (ticker, df) in enumerate(all_data):
        # Identify Year Columns
        # Filter out 'Metric', 'TTM' and likely interim columns like 'Sep' to focus on Annual
        all_cols = [c for c in df.columns if c != 'Metric' and c != 'TTM' and not str(c).startswith('Sep ')]
        
        # Take last 3 columns max (User request: Max 3 years)
        valid_cols = all_cols[-3:]
        
        # Merge Header
        start_col = current_col
        end_col = current_col + len(valid_cols) - 1
        
        cell = ws.cell(row=1, column=start_col, value=ticker)
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        
        # Style
        color = header_colors[i] if i < len(header_colors) else "808080"
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        
        # Apply style to merged range
        for c_idx in range(start_col, end_col + 1):
             c = ws.cell(row=1, column=c_idx)
             c.fill = fill
             c.font = font_white_bold
             c.alignment = align_center
             c.border = thin_border
        
        # Sub-headers (Years)
        for j, col_name in enumerate(valid_cols):
            c = ws.cell(row=2, column=start_col + j, value=col_name)
            c.font = font_bold
            c.alignment = align_center
            c.border = thin_border
            
        data_map.append({'df': df, 'cols': valid_cols, 'col_idx': start_col})
        current_col = end_col + 1

    # Write Data
    row_idx = 3
    for cat_name, metrics_list in categories.items():
        # Write Category Heading
        c = ws.cell(row=row_idx, column=1, value=cat_name)
        c.font = Font(bold=True, size=14) 
        
        row_idx += 1
        
        has_data_in_cat = False
        
        for metric in metrics_list:
            # Write Metric Name (Candidate)
            c = ws.cell(row=row_idx, column=1, value=metric)
            c.font = font_bold
            c.border = thin_border
            
            is_empty_row = True
            
            for item in data_map:
                df = item['df']
                target_cols = item['cols']
                start_c = item['col_idx']
                
                # Find metric in DF
                # Use startswith for 'Sales' matching 'Sales +'
                # Ensure string type
                match = df[df['Metric'].astype(str).str.startswith(metric)]
                
                if match.empty:
                     # Try exact match sanitized
                     match = df[df['Metric'].astype(str).apply(lambda x: x.replace('+', '').strip()) == metric]
                
                # Fill borders for all target cells first (default)
                for k in range(len(target_cols)):
                    ws.cell(row=row_idx, column=start_c + k).border = thin_border

                if not match.empty:
                    vals = match.iloc[0]
                    for k, col_name in enumerate(target_cols):
                        val = vals.get(col_name)
                        cell = ws.cell(row=row_idx, column=start_c + k)
                        
                        if pd.notna(val):
                            cell.value = val
                            is_empty_row = False
                            
                            # Apply % if needed
                            if 'Margin' in metric or 'Yield' in metric or 'RO' in metric or '%' in metric: 
                                 if isinstance(val, (int, float)) and val < 10: 
                                      cell.number_format = '0.00%'
                            else:
                                 cell.number_format = '#,##0.00'
                        
                        cell.border = thin_border

            if not is_empty_row:
                 row_idx += 1
                 has_data_in_cat = True
                 
        if has_data_in_cat:
             row_idx += 2 # Leave 2 lines after category
             
    # Auto Width
    ws.column_dimensions['A'].width = 30
    for col in range(2, current_col):
        col_letter = ws.cell(row=2, column=col).column_letter
        ws.column_dimensions[col_letter].width = 15

    # Create filename with company names
    companies_str = "_".join(companies)
    filename = f"Comparison_{companies_str}.xlsx"
    
    if in_memory:
        out_buffer = io.BytesIO()
        wb.save(out_buffer)
        print("Generated comparison in memory.")
        out_buffer.seek(0)
        return out_buffer
    else:
        comparison_dir = os.path.join(os.getcwd(), 'Comparison Data')
        if not os.path.exists(comparison_dir):
            os.makedirs(comparison_dir)
            
        out_file = os.path.join(comparison_dir, filename)
        wb.save(out_file)
        print(f"Saved comparison to {out_file}")
        return out_file

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python company_comparator.py TICKER1 [TICKER2] [TICKER3]")
        sys.exit(1)
        
    tickers = sys.argv[1:]
    if len(tickers) > 3:
        tickers = tickers[:3]
        print("Truncating to first 3 tickers.")
        
    create_comparison_sheet(tickers)
