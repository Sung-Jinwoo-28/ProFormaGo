from flask import Flask, render_template, request, send_file, jsonify
import os
from screener_extractor import ScreenerExtractor
from company_comparator import create_comparison_sheet

app = Flask(__name__)

# Ensure the Companies and Comparison Data directories exist
COMPANIES_DIR = os.path.join(os.getcwd(), 'Companies')
COMPARISON_DIR = os.path.join(os.getcwd(), 'Comparison Data')

if not os.path.exists(COMPANIES_DIR):
    os.makedirs(COMPANIES_DIR)
if not os.path.exists(COMPARISON_DIR):
    os.makedirs(COMPARISON_DIR)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        ticker = data.get('ticker')
        years = data.get('years')

        if not ticker:
            return jsonify({'error': 'Ticker is required'}), 400
        
        try:
            years = int(years) if years else None
        except ValueError:
             return jsonify({'error': 'Years must be a number'}), 400

        # Generate in memory
        extractor = ScreenerExtractor(ticker, years)
        file_buffer = extractor.run(in_memory=True)

        if file_buffer:
            filename = f"{ticker}_Financial_Model.xlsx"
            return send_file(
                file_buffer, 
                as_attachment=True, 
                download_name=filename, 
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
             return jsonify({'error': 'Failed to generate file. Check ticker or try again.'}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/compare', methods=['POST'])
def compare():
    try:
        data = request.get_json()
        tickers = data.get('tickers') 

        if not tickers:
            return jsonify({'error': 'Tickers are required'}), 400
        
        if isinstance(tickers, str):
            tickers = [t.strip() for t in tickers.split(',') if t.strip()]

        if len(tickers) < 1:
             return jsonify({'error': 'At least one ticker is required'}), 400
        
        if len(tickers) > 3:
             return jsonify({'error': 'Maximum 3 tickers allowed'}), 400

        file_buffer = create_comparison_sheet(tickers, in_memory=True)

        if file_buffer:
            filename = f"Comparison_{'_'.join(tickers)}.xlsx"
            return send_file(
                file_buffer, 
                as_attachment=True, 
                download_name=filename, 
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
             return jsonify({'error': 'Failed to generate comparison file.'}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=8000)
