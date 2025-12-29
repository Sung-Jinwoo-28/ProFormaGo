import sys
import os
from screener_extractor import ScreenerExtractor
from company_comparator import create_comparison_sheet

def main():
    while True:
        print("\n--- Financial Data Bot ---")
        print("1. Data Extraction")
        print("2. Compare 3 companies 3 years data")
        print("3. Exit")
        
        choice = input("Select an option (1/2/3): ").strip()
        
        if choice == '1':
            ticker = input("Which ticker do you want to extract? ").strip().upper()
            if not ticker:
                print("Ticker cannot be empty.")
                continue
                
            years_str = input("How many years of data do you want? ").strip()
            try:
                years = int(years_str)
            except ValueError:
                print("Invalid number for years. Defaulting to 5.")
                years = 5
            
            print(f"\nExtracting {years} years of data for {ticker}...")
            try:
                companies_dir = os.path.join(os.getcwd(), 'Companies')
                if not os.path.exists(companies_dir):
                    os.makedirs(companies_dir)
                    
                extractor = ScreenerExtractor(ticker, years, output_folder=companies_dir)
                extractor.run()
                # Filename is dynamic with timestamp now, extractor prints it.
            except Exception as e:
                print(f"An error occurred: {e}")
                
        elif choice == '2':
            tickers_input = input("Give 3 tickers of the companies separated by comma: ").strip()
            if not tickers_input:
                print("No tickers provided.")
                continue
            
            # Split and clean
            tickers = [t.strip().upper() for t in tickers_input.split(',')]
            tickers = [t for t in tickers if t] # remove empty strings
            
            if len(tickers) > 3:
                print("More than 3 tickers provided. Taking the first 3.")
                tickers = tickers[:3]
            elif len(tickers) < 1:
                print("Please provide at least one ticker.")
                continue
                
            print(f"\nComparing {', '.join(tickers)}...")
            try:
                create_comparison_sheet(tickers)
            except Exception as e:
                print(f"An error occurred during comparison: {e}")
                
        elif choice == '3':
            print("Exiting...")
            break
        else:
            print("Invalid option. Please try again.")

if __name__ == "__main__":
    main()
