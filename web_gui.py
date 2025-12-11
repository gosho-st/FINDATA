from flask import Flask, render_template, jsonify, request
from flask_cors import CORS
import pandas as pd
import threading
import os
import time
from datetime import datetime as dt
import random
from collections import defaultdict
import webbrowser

# Import the scraping logic from the original file
from financial_data_gui import (
    EXCHANGE_MAPPING, EXCHANGE_DISPLAY_NAMES, EXCHANGE_SORT_ORDER,
    CURRENCY_SYMBOLS, get_currency_symbol, normalize_company_name,
    normalize_ticker_for_alphaspread
)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Get the directory where this script is located (for proper path resolution)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, 
            static_folder=os.path.join(SCRIPT_DIR, 'static'), 
            template_folder=os.path.join(SCRIPT_DIR, 'templates'))
CORS(app)

# Global state
scraper_state = {
    'status': 'Ready',
    'progress': 0,
    'is_running': False,
    'current_ticker': None,
    'output_file': None,
    'error': None
}

companies_df = None
company_tickers = defaultdict(list)
ticker_to_company = {}


def load_companies():
    """Load company data from GitHub CSV."""
    global companies_df, company_tickers, ticker_to_company
    
    import urllib.request
    import io
    
    url = "https://raw.githubusercontent.com/gosho-st/exchanges/refs/heads/main/all_exchanges_stocks_20251204_201504.csv"
    
    try:
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            csv_data = response.read().decode('utf-8')
        
        df = pd.read_csv(io.StringIO(csv_data))
        df = df.rename(columns={
            'symbol': 'Symbol',
            'name': 'Name',
            'Exchange': 'ExchangeCode',
            'Exchange Name': 'ExchangeName'
        })
        
        df['Symbol'] = df['Symbol'].astype(str).str.strip()
        df['Name'] = df['Name'].fillna('')
        df['Exchange'] = df['ExchangeCode'].fillna('')
        df['AlphaSpreadExchange'] = df['ExchangeCode'].map(
            lambda x: EXCHANGE_MAPPING.get(x, x.lower() if x else 'nyse')
        )
        
        # Build ticker mappings
        for _, row in df.iterrows():
            symbol = row['Symbol']
            name = row['Name']
            exchange = row['AlphaSpreadExchange']
            normalized_name = normalize_company_name(name)
            
            ticker_info = {
                'symbol': symbol,
                'exchange': exchange,
                'original_name': name
            }
            
            if normalized_name:
                company_tickers[normalized_name].append(ticker_info)
            
            ticker_to_company[symbol.upper()] = {
                'name': name,
                'normalized_name': normalized_name,
                'exchange': exchange
            }
        
        companies_df = df
        print(f"Loaded {len(df)} companies")
        return True
    except Exception as e:
        print(f"Error loading companies: {e}")
        return False


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/support')
def support():
    return render_template('support.html')


@app.route('/api/exchanges')
def get_exchanges():
    """Get list of exchanges sorted by listing count."""
    if companies_df is None:
        return jsonify([])
    
    raw_exchanges = companies_df['Exchange'].dropna().unique().tolist()
    
    def sort_key(ex):
        try:
            return EXCHANGE_SORT_ORDER.index(ex)
        except ValueError:
            return len(EXCHANGE_SORT_ORDER)
    
    sorted_exchanges = sorted(raw_exchanges, key=sort_key)
    
    result = [{'code': 'ALL', 'name': 'All Exchanges'}]
    for ex in sorted_exchanges:
        display = EXCHANGE_DISPLAY_NAMES.get(ex, ex)
        result.append({'code': ex, 'name': display})
    
    return jsonify(result)


@app.route('/api/search')
def search_companies():
    """Search for companies by ticker or name."""
    query = request.args.get('q', '').upper().strip()
    exchange = request.args.get('exchange', 'ALL')
    
    if companies_df is None:
        return jsonify([])
    
    # Filter by exchange
    if exchange and exchange != 'ALL':
        filtered = companies_df[companies_df['Exchange'] == exchange]
    else:
        filtered = companies_df
    
    if query:
        # Search by ticker and name
        exact = filtered[filtered['Symbol'].str.upper() == query]
        starts = filtered[
            (filtered['Symbol'].str.upper().str.startswith(query, na=False)) & 
            (filtered['Symbol'].str.upper() != query)
        ]
        name_match = filtered[
            filtered['Name'].str.upper().str.contains(query, na=False) & 
            ~filtered['Symbol'].str.upper().str.startswith(query, na=False)
        ]
        results = pd.concat([exact, starts, name_match]).head(50)
    else:
        # Random 10 samples
        if len(filtered) > 10:
            results = filtered.sample(n=10)
        else:
            results = filtered
    
    return jsonify([
        {
            'symbol': row['Symbol'],
            'name': row['Name'][:60],
            'exchange': row['Exchange']
        }
        for _, row in results.iterrows()
    ])


@app.route('/api/status')
def get_status():
    """Get current scraper status."""
    return jsonify(scraper_state)


@app.route('/api/download', methods=['POST'])
def start_download():
    """Start downloading data for a ticker."""
    global scraper_state
    
    if scraper_state['is_running']:
        return jsonify({'error': 'Already running'}), 400
    
    data = request.json
    ticker = data.get('ticker')
    company_name = data.get('name', ticker)
    
    if not ticker:
        return jsonify({'error': 'No ticker provided'}), 400
    
    # Start download in background thread
    thread = threading.Thread(target=run_scraper, args=(ticker, company_name))
    thread.daemon = True
    thread.start()
    
    return jsonify({'status': 'started'})


def run_scraper(ticker, company_name):
    """Run the scraper in background - standalone without Tkinter."""
    global scraper_state
    
    scraper_state = {
        'status': f'Starting download for {ticker}...',
        'progress': 0,
        'is_running': True,
        'current_ticker': ticker,
        'output_file': None,
        'error': None
    }
    
    driver = None
    
    def update_status(msg, progress=None):
        scraper_state['status'] = msg
        if progress is not None:
            scraper_state['progress'] = progress
    
    try:
        # Set up output file
        downloads_folder = os.path.expanduser("~/Downloads")
        timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "".join(c for c in company_name if c.isalnum() or c in (' ', '-', '_')).strip()[:30]
        output_file = os.path.join(downloads_folder, f"{ticker}_{safe_name}_{timestamp}.xlsx")
        
        update_status('Setting up browser...', 5)
        
        # Set up Chrome options
        options = Options()
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        options.add_argument('--blink-settings=imagesEnabled=false')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.page_load_strategy = 'eager'
        prefs = {
            'profile.managed_default_content_settings.images': 2,
            'profile.managed_default_content_settings.stylesheets': 2,
            'profile.default_content_setting_values.notifications': 2,
            'disk-cache-size': 4096
        }
        options.add_experimental_option('prefs', prefs)
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        # Get alternative tickers to try
        alternatives = get_alternative_tickers_standalone(ticker, company_name)
        
        update_status(f"Searching across {len(alternatives)} exchange listings...", 10)
        
        # Try each alternative until we find one that works
        found_ticker = None
        found_exchange = None
        
        for idx, (alt_ticker, alt_exchange) in enumerate(alternatives[:10]):
            base_url = f"https://www.alphaspread.com/security/{alt_exchange}/{alt_ticker}/financials"
            update_status(f"Trying {alt_ticker.upper()} on {alt_exchange.upper()}... ({idx+1}/{min(len(alternatives), 10)})", 10 + idx * 2)
            
            try:
                driver.get(f"{base_url}/income-statement")
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.income-statement.statement'))
                )
                found_ticker = alt_ticker
                found_exchange = alt_exchange
                print(f"Found data for {ticker} as {alt_ticker} on {alt_exchange}")
                break
            except:
                continue
        
        if not found_ticker:
            raise Exception(f"Could not find financial data for {ticker} ({company_name}) on any exchange.")
        
        base_url = f"https://www.alphaspread.com/security/{found_exchange}/{found_ticker}/financials"
        update_status(f"Found on {found_exchange.upper()} - Fetching data...", 30)
        
        # Import the GUI class just for its methods (won't create window)
        from financial_data_gui import SimpleFinanceGUI
        
        # Create a minimal instance without initializing Tkinter
        gui = object.__new__(SimpleFinanceGUI)
        gui.companies_df = companies_df
        gui.company_tickers = company_tickers
        gui.ticker_to_company = ticker_to_company
        gui.output_file = output_file
        gui.company_info = {'name': company_name, 'currency': 'USD'}
        gui.selected_ticker = found_ticker
        gui.selected_company_name = company_name
        gui._workbook = None
        gui._formats = {}
        
        all_data = {}
        
        # Define statements
        statements = [
            ('income-statement', 'Income Statement'),
            ('balance-sheet', 'Balance Sheet'),
            ('cash-flow-statement', 'Cash Flow Statement')
        ]
        
        # Open tabs for parallel fetching
        update_status("Opening parallel tabs...", 35)
        tab_handles = []
        
        for idx, (key, name) in enumerate(statements):
            if idx == 0:
                driver.get(f"{base_url}/{key}")
                tab_handles.append(driver.current_window_handle)
            else:
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(f"{base_url}/{key}")
                tab_handles.append(driver.current_window_handle)
            time.sleep(0.2)
        
        # Wait for pages to load
        for idx, (key, name) in enumerate(statements):
            driver.switch_to.window(tab_handles[idx])
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, f'.{key}.statement'))
                )
            except:
                pass
        
        # Extract company info
        driver.switch_to.window(tab_handles[0])
        company_name_extracted, currency = gui.extract_company_info(driver)
        gui.company_info = {'name': company_name_extracted or company_name, 'currency': currency or 'USD'}
        
        update_status("Fetching financial statements...", 40)
        
        # Extract default data from all tabs
        for idx, (key, name) in enumerate(statements):
            driver.switch_to.window(tab_handles[idx])
            dates, fields, selected = gui.extract_data(driver, key)
            if dates and fields:
                df = gui.parse_data(dates, fields, selected)
                if df is not None:
                    all_data[f"{name} ({selected})"] = df
        
        # Fetch different periods
        periods_config = {
            'income-statement': ['Annual', 'Quarterly', 'TTM'],
            'balance-sheet': ['Annual', 'Quarterly'],
            'cash-flow-statement': ['Annual', 'Quarterly', 'TTM']
        }
        
        all_periods = ['Annual', 'Quarterly', 'TTM']
        
        for period_idx, period in enumerate(all_periods):
            update_status(f"Fetching {period} data...", 45 + period_idx * 15)
            
            # Click period buttons on all applicable tabs
            for idx, (key, name) in enumerate(statements):
                if period in periods_config[key]:
                    sheet_key = f"{name} ({period})"
                    if sheet_key not in all_data:
                        driver.switch_to.window(tab_handles[idx])
                        gui.click_period_fast(driver, period, key)
            
            time.sleep(2.0)
            
            for idx, (key, name) in enumerate(statements):
                if period in periods_config[key]:
                    sheet_key = f"{name} ({period})"
                    if sheet_key not in all_data:
                        driver.switch_to.window(tab_handles[idx])
                        dates, fields, selected = gui.extract_data_livewire(driver, key)
                        if dates and fields:
                            df = gui.parse_data(dates, fields, selected)
                            if df is not None:
                                actual_key = f"{name} ({selected})"
                                if actual_key not in all_data:
                                    all_data[actual_key] = df
                        else:
                            dates, fields, selected = gui.extract_data(driver, key)
                            if dates and fields:
                                df = gui.parse_data(dates, fields, selected)
                                if df is not None:
                                    actual_key = f"{name} ({selected})"
                                    if actual_key not in all_data:
                                        all_data[actual_key] = df
        
        # Fetch Revenue Breakdown
        update_status("Fetching Revenue Breakdown...", 85)
        driver.switch_to.window(tab_handles[0])
        breakdown_data = gui.scrape_revenue_breakdown_fast(driver, base_url)
        
        if all_data:
            update_status("Saving Excel file...", 90)
            
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                sheet_order = [
                    'Income Statement (Annual)',
                    'Balance Sheet (Annual)',
                    'Cash Flow Statement (Annual)',
                    'Income Statement (Quarterly)',
                    'Balance Sheet (Quarterly)',
                    'Cash Flow Statement (Quarterly)',
                    'Income Statement (TTM)',
                    'Cash Flow Statement (TTM)',
                ]
                
                # Write sheets in the specified order
                for sheet_name in sheet_order:
                    if sheet_name in all_data:
                        gui.format_excel_sheet_optimized(writer, all_data[sheet_name], sheet_name[:31])
                
                # Write any remaining sheets that weren't in the order list
                for sheet_name, df in all_data.items():
                    if sheet_name not in sheet_order:
                        gui.format_excel_sheet_optimized(writer, df, sheet_name[:31])
                
                # Add Revenue Breakdown
                if breakdown_data:
                    gui.format_revenue_breakdown_sheet(writer, breakdown_data)
            
            scraper_state['status'] = 'Complete!'
            scraper_state['progress'] = 100
            scraper_state['output_file'] = output_file
        else:
            raise Exception("No financial data could be extracted")
        
    except Exception as e:
        scraper_state['status'] = f'Error: {str(e)}'
        scraper_state['error'] = str(e)
        import traceback
        traceback.print_exc()
    
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
        scraper_state['is_running'] = False


def get_alternative_tickers_standalone(ticker, company_name):
    """Get all alternative tickers for a company to try."""
    alternatives = []
    seen_tickers = set()
    
    def add_ticker(t, exc='nyse'):
        t_lower = t.lower()
        if t_lower not in seen_tickers and t_lower:
            seen_tickers.add(t_lower)
            alternatives.append((t_lower, exc))
    
    # First, try the selected ticker as-is
    add_ticker(ticker)
    
    # Try with leading zeros stripped (for HK tickers like 0700 -> 700)
    stripped = ticker.lstrip('0')
    if stripped and stripped != ticker:
        add_ticker(stripped)
    
    # Try to find alternatives by company name
    normalized_name = normalize_company_name(company_name)
    if normalized_name and company_tickers:
        company_alts = company_tickers.get(normalized_name, [])
        
        # Sort by exchange priority
        exchange_priority = {'nyse': 1, 'nasdaq': 2, 'lse': 3, 'hkex': 4, 'tse': 5, 'xetra': 6, 'otc': 99}
        company_alts = sorted(company_alts, key=lambda x: exchange_priority.get(x.get('exchange', ''), 50))
        
        for alt in company_alts:
            alt_symbol = alt['symbol']
            add_ticker(alt_symbol)
            stripped_alt = alt_symbol.lstrip('0')
            if stripped_alt and stripped_alt != alt_symbol:
                add_ticker(stripped_alt)
    
    return alternatives


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    """Open the downloads folder."""
    downloads_folder = os.path.expanduser("~/Downloads")
    
    import subprocess
    import sys
    
    if sys.platform == 'darwin':
        subprocess.run(['open', downloads_folder])
    elif sys.platform == 'win32':
        subprocess.run(['explorer', downloads_folder])
    else:
        subprocess.run(['xdg-open', downloads_folder])
    
    return jsonify({'status': 'opened'})


PORT = 5050  # Using 5050 because macOS AirPlay uses 5000


def open_browser():
    """Open the browser after a short delay."""
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{PORT}')


if __name__ == '__main__':
    print("Loading company data...")
    load_companies()
    
    print("\n" + "="*50)
    print("  FINDATA — Financial Fundamental Data")
    print(f"  Open http://127.0.0.1:{PORT} in your browser")
    print("="*50 + "\n")
    
    # Open browser automatically
    threading.Thread(target=open_browser, daemon=True).start()
    
    app.run(debug=False, port=PORT, threaded=True)
