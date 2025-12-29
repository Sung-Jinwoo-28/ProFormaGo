import os
import pandas as pd
from screener_extractor import ScreenerExtractor
import time

# Companies to analyze
COMPANIES = ['RELIANCE', 'TCS', 'INFY', 'ITC', 'LT']
YEARS = 5

def generate_data():
    print(f"Starting analysis for: {', '.join(COMPANIES)}\n")
    generated_files = []
    
    for ticker in COMPANIES:
        print(f"--- Processing {ticker} ---")
        try:
            # Initialize and run extraction
            extractor = ScreenerExtractor(ticker, YEARS)
            extractor.run()
            
            filename = f"{ticker}_Financial_Model.xlsx"
            if os.path.exists(filename):
                generated_files.append((ticker, filename))
                print(f"Successfully generated {filename}")
            else:
                print(f"Failed to generate output for {ticker}")
        except Exception as e:
            print(f"Error processing {ticker}: {e}")
        
        # Be polite to the server
        time.sleep(2)
        print("\n")
        
    return generated_files

def analyze_and_rank(files):
    print("Consolidating Data and Calculating Rankings...")
    summary_data = []
    
    # Metrics to extract
    # Operating Metrics
    # Revenue (Sales), OPM %, ROCE %
    # Financing Metrics
    # Debt/Equity, Interest Coverage, Current Ratio
    
    metrics_map = {
        'Sales': 'Revenue',
        'Operating Profit Margin %': 'OPM %',
        'ROCE %': 'ROCE %',
        'Debt to Equity Ratio': 'Debt/Equity',
        'Interest Coverage Ratio': 'Interest Cov.',
        'Current Ratio': 'Current Ratio'
    }
    
    for ticker, filepath in files:
        try:
            df = pd.read_excel(filepath, sheet_name='Investing Ratios')
            
            # We use the TTM column for the most recent comparison, or the last historical year if TTM is missing.
            # But TTM is best for "Current Performance".
            col_to_use = 'TTM' if 'TTM' in df.columns else df.columns[-1]
            print(f"Using column '{col_to_use}' for {ticker}")
            
            company_metrics = {'Company': ticker}
            
            for metric_name, short_name in metrics_map.items():
                row = df[df['Metric'] == metric_name]
                if not row.empty:
                    val = row.iloc[0][col_to_use]
                    company_metrics[short_name] = val
                else:
                    company_metrics[short_name] = None
            
            summary_data.append(company_metrics)
            
        except Exception as e:
            print(f"Error reading ratios for {ticker}: {e}")

    summary_df = pd.DataFrame(summary_data)
    
    # Calculate Averages
    averages = summary_df.mean(numeric_only=True)
    # Add Average Row properly
    avg_row = averages.to_dict()
    avg_row['Company'] = 'AVERAGE'
    summary_df = pd.concat([summary_df, pd.DataFrame([avg_row])], ignore_index=True)

    print("\n--- Summary Data ---")
    print(summary_df)

    # Ranking Logic
    # Rank Companies (exclude Average row)
    rank_df = summary_df[summary_df['Company'] != 'AVERAGE'].copy()
    
    # Define Ranking Direction
    # Higher is Better: Revenue, OPM, ROCE, Interest Cov, Current Ratio
    # Lower is Better: Debt/Equity
    
    score_cols = []
    
    # Operating Score
    rank_df['Rank_Revenue'] = rank_df['Revenue'].rank(ascending=True) # Higher rank score = Better? 
    # Rank returns 1..N. usually 1 is lowest. We want 5 to be best?
    # Let's standardize: Score 1 (Best) to 5 (Worst) or vice versa.
    # Assignment says "Rank companies... choose best".
    # I'll assign Score: Higher Value = Higher Score for "Higher is Better".
    
    for m in ['Revenue', 'OPM %', 'ROCE %', 'Interest Cov.', 'Current Ratio']:
        rank_df[f'Score_{m}'] = rank_df[m].rank(ascending=True) # Max value gets Rank 5 (if 5 items)
        score_cols.append(f'Score_{m}')
        
    # Debt/Equity: Lower is Better. So use ascending=False (Low val gets High Rank)
    rank_df['Score_Debt/Equity'] = rank_df['Debt/Equity'].rank(ascending=False)
    score_cols.append('Score_Debt/Equity')
    
    # Calculate Composite Scores
    # Operating Score = Avg of Rev, OPM, ROCE
    rank_df['Operating Score'] = rank_df[['Score_Revenue', 'Score_OPM %', 'Score_ROCE %']].mean(axis=1)
    
    # Financing Score = Avg of Debt/Eq, Int Cov, Current Ratio
    rank_df['Financing Score'] = rank_df[['Score_Debt/Equity', 'Score_Interest Cov.', 'Score_Current Ratio']].mean(axis=1)
    
    # Total Score
    rank_df['Total Score'] = rank_df[['Operating Score', 'Financing Score']].mean(axis=1)
    
    # Sort by Total Score Descending (Best First)
    rank_df = rank_df.sort_values(by='Total Score', ascending=False)
    
    best_company = rank_df.iloc[0]['Company']
    print(f"\n🏆 Best Performing Company: {best_company}")
    
    # Save Report
    output_file = 'Assignment_Report.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        summary_df.to_excel(writer, sheet_name='Summary Data', index=False)
        rank_df.to_excel(writer, sheet_name='Rankings', index=False)
        
    print(f"Report saved to {output_file}")


if __name__ == "__main__":
    files = generate_data()
    if files:
        analyze_and_rank(files)
