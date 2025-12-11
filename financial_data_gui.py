import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import re
import json
import time
from datetime import datetime as dt
import random
from collections import defaultdict

# Try to import PIL for background image
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# Exchange mapping for AlphaSpread URLs
EXCHANGE_MAPPING = {
    'NYSE': 'nyse',
    'NASDAQ': 'nasdaq',
    'OTC': 'otc',
    'US OTC': 'otc',
    'FRA': 'xetra',  # Frankfurt uses xetra on AlphaSpread
    'Frankfurt Stock Exchange': 'xetra',
    'TYO': 'tse',  # Tokyo Stock Exchange
    'Tokyo Stock Exchange': 'tse',
    'HKEX': 'hkex',
    'Hong Kong Stock Exchange': 'hkex',
    'BOM': 'bse',  # Bombay Stock Exchange
    'Bombay Stock Exchange': 'bse',
    'LSE': 'lse',  # London Stock Exchange
    'London Stock Exchange': 'lse',
    'TSX': 'tsx',  # Toronto Stock Exchange
    'ASX': 'asx',  # Australian Stock Exchange
}

# Exchange code to full name mapping (sorted by number of listings)
EXCHANGE_DISPLAY_NAMES = {
    'OTC': 'OTC - US OTC Markets',
    'US OTC': 'OTC - US OTC Markets',
    'FRA': 'FRA - Frankfurt Stock Exchange',
    'Frankfurt Stock Exchange': 'FRA - Frankfurt Stock Exchange',
    'BOM': 'BOM - Bombay Stock Exchange',
    'Bombay Stock Exchange': 'BOM - Bombay Stock Exchange',
    'TYO': 'TYO - Tokyo Stock Exchange',
    'Tokyo Stock Exchange': 'TYO - Tokyo Stock Exchange',
    'LON': 'LON - London Stock Exchange',
    'LSE': 'LON - London Stock Exchange',
    'London Stock Exchange': 'LON - London Stock Exchange',
    'NASDAQ': 'NASDAQ - Nasdaq Stock Market',
    'SHE': 'SHE - Shenzhen Stock Exchange',
    'NSE': 'NSE - National Stock Exchange of India',
    'HKG': 'HKG - Hong Kong Stock Exchange',
    'HKEX': 'HKG - Hong Kong Stock Exchange',
    'Hong Kong Stock Exchange': 'HKG - Hong Kong Stock Exchange',
    'SHA': 'SHA - Shanghai Stock Exchange',
    'NYSE': 'NYSE - New York Stock Exchange',
    'ASX': 'ASX - Australian Securities Exchange',
    'KOSDAQ': 'KOSDAQ - KOSDAQ',
    'TSXV': 'TSXV - TSX Venture Exchange',
    'TPEX': 'TPEX - Taipei Exchange',
    'BIT': 'BIT - Borsa Italiana',
    'BKK': 'BKK - Stock Exchange of Thailand',
    'TPE': 'TPE - Taiwan Stock Exchange',
    'KLSE': 'KLSE - Bursa Malaysia',
    'ETR': 'ETR - Deutsche Börse Xetra',
    'BVMF': 'BVMF - Brazil Stock Exchange',
    'KRX': 'KRX - Korea Stock Exchange',
    'IDX': 'IDX - Indonesia Stock Exchange',
    'VIE': 'VIE - Vienna Stock Exchange',
    'STO': 'STO - Nasdaq Stockholm',
    'EPA': 'EPA - Euronext Paris',
    'WSE': 'WSE - Warsaw Stock Exchange',
    'TSX': 'TSX - Toronto Stock Exchange',
    'BMV': 'BMV - Mexican Stock Exchange',
    'CSE': 'CSE - Canadian Securities Exchange',
    'AIM': 'AIM - London Stock Exchange AIM',
    'IST': 'IST - Istanbul Stock Exchange',
    'TLV': 'TLV - Tel Aviv Stock Exchange',
    'PSX': 'PSX - Pakistan Stock Exchange',
    'SGX': 'SGX - Singapore Exchange',
    'HOSE': 'HOSE - Ho Chi Minh Stock Exchange',
    'TADAWUL': 'TADAWUL - Saudi Stock Exchange',
    'DSE': 'DSE - Dhaka Stock Exchange',
    'COSE': 'COSE - Colombo Stock Exchange',
    'OSL': 'OSL - Oslo Børs',
    'HNX': 'HNX - Hanoi Stock Exchange',
    'PSE': 'PSE - Philippine Stock Exchange',
    'JSE': 'JSE - Johannesburg Stock Exchange',
    'SWX': 'SWX - SIX Swiss Exchange',
    'BME': 'BME - Madrid Stock Exchange',
    'AMEX': 'AMEX - NYSE American',
    'EGX': 'EGX - Egyptian Stock Exchange',
    'SNSE': 'SNSE - Santiago Stock Exchange',
    'BVL': 'BVL - Lima Stock Exchange',
    'BCBA': 'BCBA - Buenos Aires Stock Exchange',
    'HEL': 'HEL - Nasdaq Helsinki',
    'BVB': 'BVB - Bucharest Stock Exchange',
    'MOEX': 'MOEX - Moscow Stock Exchange',
    'SGXC': 'SGXC - Singapore Exchange Catalist',
    'CPH': 'CPH - Copenhagen Stock Exchange',
    'ATH': 'ATH - Athens Stock Exchange',
    'ASE': 'ASE - Amman Stock Exchange',
    'NGX': 'NGX - Nigerian Stock Exchange',
    'KWSE': 'KWSE - Kuwait Stock Exchange',
    'EBR': 'EBR - Euronext Brussels',
    'MUN': 'MUN - Munich Stock Exchange',
    'NEO': 'NEO - Cboe Canada',
    'NZE': 'NZE - New Zealand Stock Exchange',
    'AMS': 'AMS - Euronext Amsterdam',
    'XKON': 'XKON - Korea New Exchange',
    'BUL': 'BUL - Bulgarian Stock Exchange',
    'BST': 'BST - Stuttgart Stock Exchange',
    'NGM': 'NGM - Nordic Growth Market',
    'MSM': 'MSM - Muscat Securities Market',
    'JMSE': 'JMSE - Jamaica Stock Exchange',
    'ADX': 'ADX - Abu Dhabi Securities Exchange',
    'LUX': 'LUX - Luxembourg Stock Exchange',
    'MUSE': 'MUSE - Mauritius Stock Exchange',
    'BVC': 'BVC - Colombia Stock Exchange',
    'HAM': 'HAM - Hamburg Stock Exchange',
    'CBSE': 'CBSE - Casablanca Stock Exchange',
    'BVMT': 'BVMT - Tunis Stock Exchange',
    'AQU': 'AQU - Aquis Exchange',
    'BUD': 'BUD - Budapest Stock Exchange',
    'DFM': 'DFM - Dubai Financial Market',
    'PRA': 'PRA - Prague Stock Exchange',
    'NASE': 'NASE - Nairobi Stock Exchange',
    'XNGO': 'XNGO - Nagoya Stock Exchange',
    'ZSE': 'ZSE - Zagreb Stock Exchange',
    'QSE': 'QSE - Qatar Stock Exchange',
    'BRVM': 'BRVM - Ivory Coast Stock Exchange',
    'ELI': 'ELI - Euronext Lisbon',
    'CYS': 'CYS - Cyprus Stock Exchange',
    'DUSE': 'DUSE - Dusseldorf Stock Exchange',
    'PEX': 'PEX - Palestine Stock Exchange',
    'XSAT': 'XSAT - Spotlight Stock Market',
    'ZMSE': 'ZMSE - Zimbabwe Stock Exchange',
    'BAX': 'BAX - Bahrain Stock Exchange',
    'NMSE': 'NMSE - Namibian Stock Exchange',
    'ICE': 'ICE - Nasdaq Iceland',
    'MSE': 'MSE - Malta Stock Exchange',
    'TAL': 'TAL - Nasdaq Tallinn',
    'GHSE': 'GHSE - Ghana Stock Exchange',
    'DAR': 'DAR - Tanzania Stock Exchange',
    'VSE': 'VSE - Nasdaq Vilnius',
    'FKSE': 'FKSE - Fukuoka Stock Exchange',
    'CCSE': 'CCSE - Caracas Stock Exchange',
    'ISE': 'ISE - Euronext Dublin',
    'BELEX': 'BELEX - Belgrade Stock Exchange',
    'KASE': 'KASE - Kazakhstan Stock Exchange',
    'LUSE': 'LUSE - Lusaka Stock Exchange',
    'BSM': 'BSM - Botswana Stock Exchange',
    'SPSE': 'SPSE - Sapporo Stock Exchange',
    'UGSE': 'UGSE - Uganda Securities Exchange',
    'LJSE': 'LJSE - Ljubljana Stock Exchange',
    'MAL': 'MAL - Malawi Stock Exchange',
    'RSE': 'RSE - Nasdaq Riga',
    'BDB': 'BDB - Beirut Stock Exchange',
    'CHIX': 'CHIX - CBOE Europe',
    'BSSE': 'BSSE - Bratislava Stock Exchange',
    'UKR': 'UKR - PFTS Stock Exchange',
}

# Sorted exchange order (by number of listings, descending)
EXCHANGE_SORT_ORDER = [
    'OTC', 'US OTC', 'FRA', 'Frankfurt Stock Exchange', 'BOM', 'Bombay Stock Exchange',
    'TYO', 'Tokyo Stock Exchange', 'LON', 'LSE', 'London Stock Exchange', 'NASDAQ',
    'SHE', 'NSE', 'HKG', 'HKEX', 'Hong Kong Stock Exchange', 'SHA', 'NYSE', 'ASX',
    'KOSDAQ', 'TSXV', 'TPEX', 'BIT', 'BKK', 'TPE', 'KLSE', 'ETR', 'BVMF', 'KRX',
    'IDX', 'VIE', 'STO', 'EPA', 'WSE', 'TSX', 'BMV', 'CSE', 'AIM', 'IST', 'TLV',
    'PSX', 'SGX', 'HOSE', 'TADAWUL', 'DSE', 'COSE', 'OSL', 'HNX', 'PSE', 'JSE',
    'SWX', 'BME', 'AMEX', 'EGX', 'SNSE', 'BVL', 'BCBA', 'HEL', 'BVB', 'MOEX',
    'SGXC', 'CPH', 'ATH', 'ASE', 'NGX', 'KWSE', 'EBR', 'MUN', 'NEO', 'NZE', 'AMS',
    'XKON', 'BUL', 'BST', 'NGM', 'MSM', 'JMSE', 'ADX', 'LUX', 'MUSE', 'BVC', 'HAM',
    'CBSE', 'BVMT', 'AQU', 'BUD', 'DFM', 'PRA', 'NASE', 'XNGO', 'ZSE', 'QSE',
    'BRVM', 'ELI', 'CYS', 'DUSE', 'PEX', 'XSAT', 'ZMSE', 'BAX', 'NMSE', 'ICE',
    'MSE', 'TAL', 'GHSE', 'DAR', 'VSE', 'FKSE', 'CCSE', 'ISE', 'BELEX', 'KASE',
    'LUSE', 'BSM', 'SPSE', 'UGSE', 'LJSE', 'MAL', 'RSE', 'BDB', 'CHIX', 'BSSE', 'UKR'
]

# Currency code to symbol mapping
CURRENCY_SYMBOLS = {
    'USD': '$',
    'EUR': '€',
    'GBP': '£',
    'JPY': '¥',
    'CNY': '¥',
    'CNH': '¥',
    'RMB': '¥',
    'HKD': 'HK$',
    'CHF': 'CHF ',
    'CAD': 'C$',
    'AUD': 'A$',
    'INR': '₹',
    'KRW': '₩',
    'SGD': 'S$',
    'TWD': 'NT$',
    'BRL': 'R$',
    'MXN': 'MX$',
    'SEK': 'kr ',
    'NOK': 'kr ',
    'DKK': 'kr ',
    'PLN': 'zł ',
    'THB': '฿',
    'IDR': 'Rp ',
    'MYR': 'RM ',
    'PHP': '₱',
    'ZAR': 'R ',
    'RUB': '₽',
    'TRY': '₺',
    'ILS': '₪',
    'AED': 'AED ',
    'SAR': 'SAR ',
}


def get_currency_symbol(currency_code):
    """Get currency symbol from currency code."""
    if not currency_code:
        return '$'
    code = currency_code.upper().strip()
    return CURRENCY_SYMBOLS.get(code, f'{code} ')


def normalize_ticker_for_alphaspread(ticker, exchange):
    """
    Normalize ticker for AlphaSpread URL.
    - Strip leading zeros for Hong Kong tickers (0700 -> 700)
    - Handle other exchange-specific quirks
    """
    ticker = str(ticker).strip()
    
    # Hong Kong Exchange: strip leading zeros
    if exchange and exchange.lower() in ['hkex', 'hong kong']:
        ticker = ticker.lstrip('0') or '0'  # Keep at least one digit
    
    return ticker.lower()


def normalize_company_name(name):
    """Normalize company name for matching."""
    if not name:
        return ''
    name = str(name).lower().strip()
    # Remove common suffixes
    suffixes = [
        'inc.', 'inc', 'corp.', 'corp', 'corporation', 'company', 'co.',
        'ltd.', 'ltd', 'limited', 'plc', 'llc', 's.a.', 'sa', 'ag', 'se',
        'n.v.', 'nv', 'holdings', 'holding', 'group', 'the', '&'
    ]
    for suffix in suffixes:
        name = name.replace(suffix, ' ')
    # Remove special characters and extra spaces
    name = re.sub(r'[^a-z0-9\s]', '', name)
    name = ' '.join(name.split())
    return name


class PerformanceTimer:
    """Track timing for different operations."""
    def __init__(self):
        self.timings = defaultdict(list)
        self.start_times = {}
    
    def start(self, operation):
        self.start_times[operation] = time.perf_counter()
    
    def stop(self, operation):
        if operation in self.start_times:
            elapsed = time.perf_counter() - self.start_times[operation]
            self.timings[operation].append(elapsed)
            del self.start_times[operation]
            return elapsed
        return 0
    
    def get_summary(self):
        summary = []
        total = 0
        for op, times in self.timings.items():
            op_total = sum(times)
            total += op_total
            summary.append(f"  {op}: {op_total:.2f}s ({len(times)} calls)")
        summary.insert(0, f"Total time: {total:.2f}s")
        return "\n".join(summary)


# Formula definitions for calculated fields
INCOME_STATEMENT_FORMULAS = {
    'Gross Profit': {'sources': ['Revenue', 'Cost of Revenue'], 'signs': ['+', '+']},
    'Operating Income': {'sources': ['Gross Profit', 'Operating Expenses'], 'signs': ['+', '+']},
    'Pre-Tax Income': {'sources': ['Operating Income', 'Interest Income Expense', 'Non-Reccuring Items', 'Total Other Income'], 'signs': ['+', '+', '+', '+']},
    'Income from Continuing Operations': {'sources': ['Pre-Tax Income', 'Tax Provision'], 'signs': ['+', '+']},
    'Net Income (Common)': {'sources': ['Income from Continuing Operations', 'Income to Minority Interest', 'Equity Earnings Affiliates'], 'signs': ['+', '+', '+']},
}

BALANCE_SHEET_FORMULAS = {
    'Total Current Assets': {'sources': ['Cash & Cash Equivalents', 'Short-Term Investments', 'Total Receivables', 'Inventory', 'Other Current Assets'], 'signs': ['+', '+', '+', '+', '+']},
    'Total Assets': {'sources': ['Total Current Assets', 'PP&E Net', 'Intangible Assets', 'Goodwill', 'Long-Term Investments', 'Other Long-Term Assets'], 'signs': ['+', '+', '+', '+', '+', '+']},
    'Total Current Liabilities': {'sources': ['Accounts Payable', 'Accrued Liabilities', 'Short-Term Debt', 'Current Portion of Long-Term Debt', 'Other Current Liabilities'], 'signs': ['+', '+', '+', '+', '+']},
    'Total Liabilities': {'sources': ['Total Current Liabilities', 'Long-Term Debt', 'Deferred Income Tax', 'Minority Interest', 'Other Liabilities'], 'signs': ['+', '+', '+', '+', '+']},
    'Total Equity': {'sources': ['Common Stock', 'Retained Earnings', 'Additional Paid In Capital', 'Unrealized Security Profit/Loss', 'Treasury Stock', 'Other Equity'], 'signs': ['+', '+', '+', '+', '+', '+']},
    'Total Liabilities & Equity': {'sources': ['Total Liabilities', 'Total Equity'], 'signs': ['+', '+']},
}

CASH_FLOW_FORMULAS = {
    'Cash from Operating Activities': {'sources': ['Net Income', 'Depreciation & Amortization', 'Change in Deffered Taxes', 'Other Non-Cash Items', 'Change in Working Capital'], 'signs': ['+', '+', '+', '+', '+']},
    'Cash from Investing Activities': {'sources': ['Capital Expenditures', 'Other Items'], 'signs': ['+', '+']},
    'Cash from Financing Activities': {'sources': ['Net Issuance of Common Stock', 'Net Issuance of Debt', 'Cash Paid for Dividends', 'Other'], 'signs': ['+', '+', '+', '+']},
    'Net Change in Cash': {'sources': ['Cash from Operating Activities', 'Cash from Investing Activities', 'Cash from Financing Activities', 'Effect of Foreign Exchange Rates'], 'signs': ['+', '+', '+', '+']},
    'Free Cash Flow': {'sources': ['Cash from Operating Activities', 'Capital Expenditures'], 'signs': ['+', '+']},
}


def get_column_letter(col_num):
    """Convert column number to Excel letter (0=A, 1=B, etc.)"""
    result = ""
    while col_num >= 0:
        result = chr(col_num % 26 + ord('A')) + result
        col_num = col_num // 26 - 1
    return result


class SimpleFinanceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AlphaSpread Financial Data")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Modern color scheme
        self.colors = {
            'bg_dark': '#1a1a2e',
            'bg_medium': '#16213e',
            'bg_light': '#0f3460',
            'accent': '#e94560',
            'accent_hover': '#ff6b6b',
            'text_white': '#ffffff',
            'text_light': '#a0a0a0',
            'text_muted': '#6c757d',
            'input_bg': '#ffffff',
            'card_bg': 'rgba(255, 255, 255, 0.95)',
        }
        
        # Configure root background
        self.root.configure(bg=self.colors['bg_dark'])
        
        # Load companies from GitHub (combined NYSE + NASDAQ list)
        self.companies_df = self.load_companies_from_github()
        
        self.selected_ticker = None
        self.selected_exchange = None
        self.output_file = None
        self.bg_image = None
        self.bg_photo = None
        self.create_ui()
    
    def load_companies_from_github(self):
        """Load company data from GitHub CSV with all exchanges."""
        import urllib.request
        import io
        
        url = "https://raw.githubusercontent.com/gosho-st/exchanges/refs/heads/main/all_exchanges_stocks_20251204_201504.csv"
        
        try:
            # Fetch CSV from GitHub
            req = urllib.request.Request(
                url, 
                headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                csv_data = response.read().decode('utf-8')
            
            # Parse CSV - columns: symbol, name, Exchange, Exchange Name
            df = pd.read_csv(io.StringIO(csv_data))
            
            # Rename columns for consistency
            df = df.rename(columns={
                'symbol': 'Symbol',
                'name': 'Name',
                'Exchange': 'ExchangeCode',
                'Exchange Name': 'ExchangeName'
            })
            
            # Clean up - ensure Symbol is a string
            df['Symbol'] = df['Symbol'].astype(str).str.strip()
            df['Name'] = df['Name'].fillna('')
            
            # Keep original exchange code for display, map to AlphaSpread format separately
            df['Exchange'] = df['ExchangeCode'].fillna('')  # Use original code for display
            df['AlphaSpreadExchange'] = df['ExchangeCode'].map(lambda x: EXCHANGE_MAPPING.get(x, x.lower() if x else 'nyse'))
            
            # Build alternative ticker mapping by normalized company name
            self.company_tickers = defaultdict(list)
            self.ticker_to_company = {}
            
            for _, row in df.iterrows():
                symbol = row['Symbol']
                name = row['Name']
                exchange = row['Exchange']  # Original code for display
                alphaspread_exchange = row['AlphaSpreadExchange']  # For URL building
                normalized_name = normalize_company_name(name)
                
                # Store ticker info (use AlphaSpread format for URL building)
                ticker_info = {
                    'symbol': symbol,
                    'exchange': alphaspread_exchange,
                    'original_name': name
                }
                
                # Group by normalized company name
                if normalized_name:
                    self.company_tickers[normalized_name].append(ticker_info)
                
                # Also allow lookup by ticker
                self.ticker_to_company[symbol.upper()] = {
                    'name': name,
                    'normalized_name': normalized_name,
                    'exchange': alphaspread_exchange
                }
            
            # Create display DataFrame - prefer major exchanges
            # Priority: NYSE > NASDAQ > LSE > HKEX > TSE > other OTC
            exchange_priority = {'NYSE': 1, 'NASDAQ': 2, 'LON': 3, 'HKG': 4, 'TYO': 5, 'FRA': 6, 'ASX': 7, 'TSX': 8, 'OTC': 99}
            df['ExchangePriority'] = df['Exchange'].map(lambda x: exchange_priority.get(x, 50))
            
            # Sort by priority and drop duplicates by name (keep best exchange)
            df = df.sort_values('ExchangePriority')
            display_df = df.drop_duplicates(subset='Name', keep='first')
            
            # Also keep unique symbols that might be missed
            all_symbols = df.drop_duplicates(subset='Symbol', keep='first')
            
            # Combine - prefer by name, but also include unique symbols
            display_df = pd.concat([display_df, all_symbols]).drop_duplicates(subset='Symbol', keep='first')
            
            print(f"Loaded {len(df)} total listings, {len(display_df)} unique companies")
            print(f"Built alternative ticker mappings for {len(self.company_tickers)} company groups")
            
            return display_df[['Symbol', 'Name', 'Exchange']]
            
        except Exception as e:
            print(f"Failed to load from GitHub: {e}")
            import traceback
            traceback.print_exc()
            print("Falling back to local files...")
            self.company_tickers = defaultdict(list)
            self.ticker_to_company = {}
            return self.load_companies_from_local_files()
    
    def load_companies_from_local_files(self):
        """Fallback: Load from local CSV files."""
        try:
            # Load NASDAQ companies
            nasdaq_path = os.path.join(os.path.dirname(__file__), "nasdaq.csv")
            nasdaq_df = pd.read_csv(nasdaq_path)
            nasdaq_df = nasdaq_df[nasdaq_df['Symbol'].str.match(r'^[A-Z]+$', na=False)]
            nasdaq_df['Exchange'] = 'nasdaq'
            nasdaq_df = nasdaq_df.rename(columns={'Symbol': 'Symbol', 'Security Name': 'Name'})
            
            # Load NYSE companies
            nyse_path = os.path.join(os.path.dirname(__file__), "nyse.csv")
            nyse_df = pd.read_csv(nyse_path)
            nyse_df = nyse_df.rename(columns={'ACT Symbol': 'Symbol', 'Company Name': 'Name'})
            nyse_df = nyse_df[nyse_df['Symbol'].str.match(r'^[A-Z]+$', na=False)]
            nyse_df['Exchange'] = 'nyse'
            
            # Combine both
            combined = pd.concat([nasdaq_df[['Symbol', 'Name', 'Exchange']], 
                                  nyse_df[['Symbol', 'Name', 'Exchange']]], ignore_index=True)
            combined = combined.drop_duplicates(subset='Symbol', keep='first')
            return combined
        except Exception as e:
            print(f"Failed to load local files: {e}")
            return pd.DataFrame(columns=['Symbol', 'Name', 'Exchange'])
    
    def create_ui(self):
        """Create a beautiful modern UI with globe background."""
        
        # Get unique exchanges from the data and sort by listing count order
        raw_exchanges = self.companies_df['Exchange'].dropna().unique().tolist()
        
        # Sort exchanges by EXCHANGE_SORT_ORDER (most listings first)
        def sort_key(ex):
            try:
                return EXCHANGE_SORT_ORDER.index(ex)
            except ValueError:
                return len(EXCHANGE_SORT_ORDER)
        
        sorted_exchanges = sorted(raw_exchanges, key=sort_key)
        
        # Build display names list with "CODE - Full Name" format
        self.exchange_code_map = {}
        display_names = ['All Exchanges']
        for ex in sorted_exchanges:
            if ex in EXCHANGE_DISPLAY_NAMES:
                display_name = EXCHANGE_DISPLAY_NAMES[ex]
            else:
                display_name = ex
            display_names.append(display_name)
            self.exchange_code_map[display_name] = ex
        
        self.available_exchanges = display_names
        
        # Create main canvas for background
        self.canvas = tk.Canvas(self.root, highlightthickness=0, bg='#dce4ec')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Load and set background image
        self.setup_background()
        
        # Bind resize event
        self.canvas.bind('<Configure>', self.on_resize)
        
        # Create styled ttk theme
        self.setup_styles()
        
        # ===== HEADER SECTION (Top) =====
        # Create rounded rectangle for header
        self.header_frame = tk.Frame(self.canvas, bg='#1e2a4a')
        self.header_frame.place(relx=0.5, y=15, anchor='n', relwidth=0.96, height=75)
        
        # Add rounded corner effect with canvas overlay
        header_inner = tk.Frame(self.header_frame, bg='#1e2a4a', padx=20, pady=10)
        header_inner.pack(fill=tk.BOTH, expand=True)
        
        # App title
        title_label = tk.Label(header_inner, 
                               text="AlphaSpread Financial Data",
                               font=('Helvetica Neue', 26, 'bold'),
                               fg='#ffffff',
                               bg='#1e2a4a')
        title_label.pack()
        
        # Tagline
        tagline_label = tk.Label(header_inner,
                                 text="Fundamental data for 50,000+ tickers across 50+ global exchanges",
                                 font=('Helvetica Neue', 11),
                                 fg='#7eb8da',
                                 bg='#1e2a4a')
        tagline_label.pack(pady=(2, 0))
        
        # ===== LEFT PANEL (Search & List) =====
        self.left_panel = tk.Frame(self.canvas, bg='#ffffff', 
                                    highlightbackground='#d0d0d0', 
                                    highlightthickness=1)
        self.left_panel.place(x=20, y=105, width=300, relheight=0.78, rely=0)
        
        inner_panel = tk.Frame(self.left_panel, bg='#ffffff', padx=15, pady=12)
        inner_panel.pack(fill=tk.BOTH, expand=True)
        
        # Panel title
        panel_title = tk.Label(inner_panel,
                               text="🔍 Search Companies",
                               font=('Helvetica Neue', 13, 'bold'),
                               fg='#1e2a4a',
                               bg='#ffffff',
                               anchor='w')
        panel_title.pack(fill=tk.X, pady=(0, 8))
        
        # Exchange dropdown
        exchange_frame = tk.Frame(inner_panel, bg='#ffffff')
        exchange_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(exchange_frame, 
                 text="Exchange",
                 font=('Helvetica Neue', 9),
                 fg='#6c757d',
                 bg='#ffffff').pack(anchor='w')
        
        self.exchange_var = tk.StringVar(value='All Exchanges')
        self.exchange_dropdown = ttk.Combobox(exchange_frame, 
                                               textvariable=self.exchange_var,
                                               values=self.available_exchanges, 
                                               state='readonly',
                                               font=('Helvetica Neue', 10),
                                               width=26)
        self.exchange_dropdown.pack(fill=tk.X, pady=(2, 0))
        self.exchange_dropdown.bind('<<ComboboxSelected>>', self.on_exchange_change)
        
        # Search entry
        search_frame = tk.Frame(inner_panel, bg='#ffffff')
        search_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(search_frame,
                 text="Search ticker or company",
                 font=('Helvetica Neue', 9),
                 fg='#6c757d',
                 bg='#ffffff').pack(anchor='w')
        
        self.search_entry = tk.Entry(search_frame,
                                      font=('Helvetica Neue', 11),
                                      relief='solid',
                                      bd=1,
                                      fg='#333333',
                                      bg='#fafafa',
                                      insertbackground='#1e2a4a')
        self.search_entry.pack(fill=tk.X, pady=(2, 0), ipady=4)
        self.search_entry.bind('<KeyRelease>', self.on_search)
        
        # Results list
        list_label = tk.Label(inner_panel,
                              text="Results (showing 10 random samples)",
                              font=('Helvetica Neue', 9),
                              fg='#6c757d',
                              bg='#ffffff')
        list_label.pack(anchor='w', pady=(4, 2))
        self.list_hint_label = list_label
        
        list_frame = tk.Frame(inner_panel, bg='#ffffff')
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.listbox = tk.Listbox(list_frame,
                                   font=('Helvetica Neue', 10),
                                   relief='solid',
                                   bd=1,
                                   bg='#fafafa',
                                   selectmode=tk.SINGLE,
                                   selectbackground='#4a90d9',
                                   selectforeground='#ffffff',
                                   activestyle='none',
                                   highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        
        self.update_list()
        
        # ===== RIGHT PANEL (Selected & Actions) =====
        self.right_panel = tk.Frame(self.canvas, bg='#ffffff',
                                     highlightbackground='#d0d0d0',
                                     highlightthickness=1)
        self.right_panel.place(relx=1.0, x=-20, y=105, width=300, height=280, anchor='ne')
        
        inner_panel_r = tk.Frame(self.right_panel, bg='#ffffff', padx=15, pady=12)
        inner_panel_r.pack(fill=tk.BOTH, expand=True)
        
        # Panel title
        panel_title_r = tk.Label(inner_panel_r,
                                  text="📊 Download Data",
                                  font=('Helvetica Neue', 13, 'bold'),
                                  fg='#1e2a4a',
                                  bg='#ffffff',
                                  anchor='w')
        panel_title_r.pack(fill=tk.X, pady=(0, 12))
        
        # Selected company display
        self.selected_frame = tk.Frame(inner_panel_r, bg='#f0f4f8')
        self.selected_frame.pack(fill=tk.X, pady=(0, 12), ipady=8, ipadx=8)
        
        self.selected_ticker_label = tk.Label(self.selected_frame,
                                               text="No ticker selected",
                                               font=('Helvetica Neue', 18, 'bold'),
                                               fg='#1e2a4a',
                                               bg='#f0f4f8')
        self.selected_ticker_label.pack(anchor='w', padx=10, pady=(8, 0))
        
        self.selected_name_label = tk.Label(self.selected_frame,
                                             text="Select a company from the list",
                                             font=('Helvetica Neue', 9),
                                             fg='#6c757d',
                                             bg='#f0f4f8',
                                             wraplength=240)
        self.selected_name_label.pack(anchor='w', padx=10, pady=(0, 8))
        
        # Download button - FIXED COLORS
        self.fetch_btn = tk.Button(inner_panel_r,
                                    text="⬇  Download Financial Data",
                                    font=('Helvetica Neue', 11, 'bold'),
                                    fg='#ffffff',
                                    bg='#3b7dd8',
                                    activebackground='#2d6bc4',
                                    activeforeground='#ffffff',
                                    disabledforeground='#a0a0a0',
                                    relief='flat',
                                    bd=0,
                                    padx=15,
                                    pady=10,
                                    cursor='hand2',
                                    state=tk.DISABLED,
                                    command=self.start_fetch)
        self.fetch_btn.pack(fill=tk.X, pady=(0, 8))
        
        # Open folder button
        self.folder_btn = tk.Button(inner_panel_r,
                                     text="📁  Open Downloads Folder",
                                     font=('Helvetica Neue', 10),
                                     fg='#333333',
                                     bg='#e8ecf0',
                                     activebackground='#d8dce0',
                                     relief='flat',
                                     bd=0,
                                     padx=12,
                                     pady=8,
                                     cursor='hand2',
                                     command=self.open_folder)
        self.folder_btn.pack(fill=tk.X)
        
        # ===== STATUS PANEL (Bottom Right) =====
        self.status_panel = tk.Frame(self.canvas, bg='#ffffff',
                                      highlightbackground='#d0d0d0',
                                      highlightthickness=1)
        self.status_panel.place(relx=1.0, x=-20, rely=1.0, y=-20, width=300, height=120, anchor='se')
        
        inner_panel_s = tk.Frame(self.status_panel, bg='#ffffff', padx=15, pady=10)
        inner_panel_s.pack(fill=tk.BOTH, expand=True)
        
        # Status title
        tk.Label(inner_panel_s,
                 text="⚡ Status",
                 font=('Helvetica Neue', 11, 'bold'),
                 fg='#1e2a4a',
                 bg='#ffffff',
                 anchor='w').pack(fill=tk.X, pady=(0, 8))
        
        # Progress bar
        self.progress = ttk.Progressbar(inner_panel_s, length=260, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(0, 6))
        
        # Status label
        self.status_label = tk.Label(inner_panel_s,
                                      text="Ready",
                                      font=('Helvetica Neue', 10),
                                      fg='#28a745',
                                      bg='#ffffff',
                                      anchor='w',
                                      wraplength=250)
        self.status_label.pack(fill=tk.X)
        
        # Keep reference to selected_label for compatibility
        self.selected_label = self.selected_ticker_label
    
    def setup_background(self):
        """Load and display the globe background image or create gradient."""
        if not HAS_PIL:
            # Fallback gradient if PIL not available
            self.create_gradient_background()
            return
        
        try:
            # Get the directory where this script is located
            script_dir = os.path.dirname(os.path.abspath(__file__))
            bg_path = os.path.join(script_dir, 'globe.png')
            
            if os.path.exists(bg_path):
                self.bg_image = Image.open(bg_path)
                self.update_background()
            else:
                # Create beautiful gradient fallback
                self.create_gradient_background()
        except Exception as e:
            print(f"Could not load background: {e}")
            self.create_gradient_background()
    
    def create_gradient_background(self):
        """Create a beautiful gradient background matching the globe theme."""
        if not HAS_PIL:
            self.canvas.configure(bg='#1a1a2e')
            return
        
        try:
            width = max(self.canvas.winfo_width(), 900)
            height = max(self.canvas.winfo_height(), 700)
            
            # Create gradient image
            gradient = Image.new('RGB', (width, height))
            
            # Define colors for gradient (matching globe image colors)
            top_color = (200, 210, 225)      # Light grayish-blue
            mid_color = (180, 200, 220)       # Soft blue
            bottom_color = (190, 195, 215)    # Light purple-gray
            
            for y in range(height):
                # Calculate blend ratios
                if y < height // 2:
                    ratio = y / (height // 2)
                    r = int(top_color[0] + (mid_color[0] - top_color[0]) * ratio)
                    g = int(top_color[1] + (mid_color[1] - top_color[1]) * ratio)
                    b = int(top_color[2] + (mid_color[2] - top_color[2]) * ratio)
                else:
                    ratio = (y - height // 2) / (height // 2)
                    r = int(mid_color[0] + (bottom_color[0] - mid_color[0]) * ratio)
                    g = int(mid_color[1] + (bottom_color[1] - mid_color[1]) * ratio)
                    b = int(mid_color[2] + (bottom_color[2] - mid_color[2]) * ratio)
                
                for x in range(width):
                    gradient.putpixel((x, y), (r, g, b))
            
            self.bg_photo = ImageTk.PhotoImage(gradient)
            self.canvas.delete('bg')
            self.canvas.create_image(0, 0, anchor='nw', image=self.bg_photo, tags='bg')
            self.canvas.tag_lower('bg')
        except Exception as e:
            print(f"Error creating gradient: {e}")
            self.canvas.configure(bg='#c8d4e3')
    
    def update_background(self):
        """Update background image to fit window - center the globe nicely."""
        if not HAS_PIL or self.bg_image is None:
            return
        
        try:
            # Get canvas dimensions
            width = self.canvas.winfo_width()
            height = self.canvas.winfo_height()
            
            if width > 1 and height > 1:
                # Calculate scale to fit the image (contain, not cover)
                img_ratio = self.bg_image.width / self.bg_image.height
                canvas_ratio = width / height
                
                if canvas_ratio > img_ratio:
                    # Canvas is wider - fit to height
                    new_height = height
                    new_width = int(height * img_ratio)
                else:
                    # Canvas is taller - fit to width
                    new_width = width
                    new_height = int(width / img_ratio)
                
                # Resize image
                resized = self.bg_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                # Create a new image with background color and paste the globe centered
                bg_color = (220, 228, 236)  # Light grayish-blue to match globe
                final_img = Image.new('RGB', (width, height), bg_color)
                
                # Center the globe image
                x_offset = (width - new_width) // 2
                y_offset = (height - new_height) // 2
                
                # Handle transparency if present
                if resized.mode == 'RGBA':
                    final_img.paste(resized, (x_offset, y_offset), resized)
                else:
                    final_img.paste(resized, (x_offset, y_offset))
                
                self.bg_photo = ImageTk.PhotoImage(final_img)
                self.canvas.delete('bg')
                self.canvas.create_image(0, 0, anchor='nw', image=self.bg_photo, tags='bg')
                self.canvas.tag_lower('bg')
        except Exception as e:
            print(f"Error updating background: {e}")
    
    def on_resize(self, event):
        """Handle window resize."""
        if self.bg_image is not None:
            self.update_background()
        else:
            self.create_gradient_background()
    
    def setup_styles(self):
        """Configure ttk styles for modern look."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Combobox style
        style.configure('TCombobox',
                        fieldbackground='#ffffff',
                        background='#ffffff',
                        foreground='#333333',
                        arrowcolor='#1a1a2e',
                        borderwidth=1,
                        relief='solid')
        
        # Progress bar style
        style.configure('TProgressbar',
                        background='#e94560',
                        troughcolor='#f0f0f0',
                        borderwidth=0,
                        lightcolor='#e94560',
                        darkcolor='#e94560')

    def on_exchange_change(self, event=None):
        """Handle exchange dropdown change."""
        self.search_entry.delete(0, tk.END)  # Clear search when changing exchange
        self.update_list()
    
    def on_search(self, event=None):
        self.update_list(self.search_entry.get().upper().strip())
    
    def update_list(self, search=""):
        self.listbox.delete(0, tk.END)
        
        # Get selected exchange filter (convert display name back to code)
        selected_display = self.exchange_var.get() if hasattr(self, 'exchange_var') else 'All Exchanges'
        
        # Convert display name to actual exchange code
        if selected_display and selected_display != 'All Exchanges':
            selected_exchange = self.exchange_code_map.get(selected_display, selected_display)
            exchange_filtered = self.companies_df[self.companies_df['Exchange'] == selected_exchange]
        else:
            exchange_filtered = self.companies_df
        
        if search:
            # When searching, show all matches (up to 100)
            exact = exchange_filtered[exchange_filtered['Symbol'].str.upper() == search]
            starts = exchange_filtered[(exchange_filtered['Symbol'].str.upper().str.startswith(search, na=False)) & (exchange_filtered['Symbol'].str.upper() != search)]
            name_match = exchange_filtered[exchange_filtered['Name'].str.upper().str.contains(search, na=False) & ~exchange_filtered['Symbol'].str.upper().str.startswith(search, na=False)]
            filtered = pd.concat([exact, starts, name_match]).head(100)
        else:
            # When not searching, show 10 random samples from the selected exchange
            if len(exchange_filtered) > 10:
                filtered = exchange_filtered.sample(n=10, random_state=None)  # Random each time
            else:
                filtered = exchange_filtered
        
        for _, row in filtered.iterrows():
            name = str(row['Name'])[:50] if pd.notna(row['Name']) else ''
            self.listbox.insert(tk.END, f"{row['Symbol']} - {name}")
    
    def on_select(self, event=None):
        sel = self.listbox.curselection()
        if sel:
            text = self.listbox.get(sel[0])
            # Parse "AAPL - Apple Inc."
            parts = text.split(' - ', 1)
            self.selected_ticker = parts[0]
            self.selected_company_name = parts[1] if len(parts) > 1 else self.selected_ticker
            # Get exchange from dataframe
            match = self.companies_df[self.companies_df['Symbol'] == self.selected_ticker]
            self.selected_exchange = match.iloc[0]['Exchange'] if len(match) > 0 else 'nyse'
            
            # Update the new styled labels
            if hasattr(self, 'selected_ticker_label'):
                self.selected_ticker_label.config(text=self.selected_ticker)
                self.selected_name_label.config(text=self.selected_company_name[:60] + ('...' if len(self.selected_company_name) > 60 else ''))
            else:
                self.selected_label.config(text=f"Selected: {text}")
            
            # Enable and highlight the download button
            self.fetch_btn.config(state=tk.NORMAL, bg='#3b7dd8', fg='#ffffff')
    
    def get_alternative_tickers(self, ticker, company_name):
        """
        Get all alternative tickers for a company to try.
        Since AlphaSpread auto-corrects exchanges, we mainly need to try different ticker formats.
        Returns list of (normalized_ticker, exchange) tuples to try.
        """
        alternatives = []
        seen_tickers = set()
        
        def add_ticker(t, exc='nyse'):
            """Add ticker if not already added (case-insensitive)."""
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
        if normalized_name and hasattr(self, 'company_tickers'):
            company_alts = self.company_tickers.get(normalized_name, [])
            
            # Sort by exchange priority (prefer major exchanges)
            exchange_priority = {'nyse': 1, 'nasdaq': 2, 'lse': 3, 'hkex': 4, 'tse': 5, 'xetra': 6, 'otc': 99}
            company_alts = sorted(company_alts, key=lambda x: exchange_priority.get(x.get('exchange', ''), 50))
            
            for alt in company_alts:
                alt_symbol = alt['symbol']
                # Add original symbol
                add_ticker(alt_symbol)
                # Also try stripped version
                stripped_alt = alt_symbol.lstrip('0')
                if stripped_alt and stripped_alt != alt_symbol:
                    add_ticker(stripped_alt)
        
        return alternatives
    
    def start_fetch(self):
        if not self.selected_ticker:
            return
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{self.selected_ticker}_financials.xlsx"
        )
        if not self.output_file:
            return
        self.fetch_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        thread = threading.Thread(target=self.fetch_data)
        thread.daemon = True
        thread.start()
    
    def status(self, msg, prog=None):
        self.status_label.config(text=msg)
        if prog is not None:
            self.progress['value'] = prog
        self.root.update_idletasks()
    
    def fetch_data(self):
        ticker = self.selected_ticker
        company_name = self.selected_company_name
        self.status(f"Fetching {self.selected_ticker}...", 5)
        
        driver = None
        timer = PerformanceTimer()
        timer.start('total')
        timer.start('browser_init')
        try:
            options = Options()
            options.add_argument('--headless=new')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36')
            # Performance optimizations
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-infobars')
            options.add_argument('--disable-logging')
            options.add_argument('--log-level=3')
            options.add_argument('--blink-settings=imagesEnabled=false')  # Disable images
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.page_load_strategy = 'eager'  # Don't wait for all resources
            prefs = {
                'profile.managed_default_content_settings.images': 2,  # Block images
                'profile.managed_default_content_settings.stylesheets': 2,  # Block CSS
                'profile.default_content_setting_values.notifications': 2,
                'disk-cache-size': 4096
            }
            options.add_experimental_option('prefs', prefs)
            
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            timer.stop('browser_init')
            
            all_data = {}
            
            # Get all alternative tickers to try
            alternatives = self.get_alternative_tickers(ticker, company_name)
            
            # Try each alternative until we find one that works
            timer.start('find_valid_ticker')
            found_ticker = None
            found_exchange = None
            
            self.status(f"Searching across {len(alternatives)} exchange listings...", 6)
            
            for idx, (alt_ticker, alt_exchange) in enumerate(alternatives[:10]):  # Limit to first 10 alternatives
                base_url = f"https://www.alphaspread.com/security/{alt_exchange}/{alt_ticker}/financials"
                self.status(f"Trying {alt_ticker.upper()} on {alt_exchange.upper()}... ({idx+1}/{min(len(alternatives), 10)})", 6 + idx)
                
                try:
                    driver.get(f"{base_url}/income-statement")
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '.income-statement.statement'))
                    )
                    # Found it!
                    found_ticker = alt_ticker
                    found_exchange = alt_exchange
                    print(f"Found data for {ticker} as {alt_ticker} on {alt_exchange}")
                    break
                except:
                    continue
            
            timer.stop('find_valid_ticker')
            
            if not found_ticker:
                raise Exception(f"Could not find financial data for {ticker} ({company_name}) on any exchange. Tried: {[f'{t}@{e}' for t, e in alternatives[:10]]}")
            
            # Update base_url with the found ticker/exchange
            base_url = f"https://www.alphaspread.com/security/{found_exchange}/{found_ticker}/financials"
            self.status(f"Found on {found_exchange.upper()} as {found_ticker.upper()}", 15)
            timer.start('initial_page_load')
            
            # Define all the URLs we need to fetch
            statements = [
                ('income-statement', 'Income Statement'),
                ('balance-sheet', 'Balance Sheet'),
                ('cash-flow-statement', 'Cash Flow Statement')
            ]
            
            # Open multiple tabs for parallel fetching (3 tabs for 3 statements)
            self.status("Opening parallel tabs...", 8)
            timer.start('open_tabs')
            tab_handles = []
            original_window = driver.current_window_handle
            
            # Open tabs for each statement type
            for idx, (key, name) in enumerate(statements):
                if idx == 0:
                    driver.get(f"{base_url}/{key}")
                    tab_handles.append(driver.current_window_handle)
                else:
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.get(f"{base_url}/{key}")
                    tab_handles.append(driver.current_window_handle)
                time.sleep(0.2)  # Reduced delay
            
            # Wait for pages with dynamic check
            for idx, (key, name) in enumerate(statements):
                driver.switch_to.window(tab_handles[idx])
                try:
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, f'.{key}.statement'))
                    )
                except:
                    pass  # Continue even if timeout
            timer.stop('open_tabs')
            
            # Fetch default data from all tabs in parallel (round-robin)
            self.status("Fetching default periods from all statements...", 15)
            timer.start('extract_default_data')
            
            # Extract company info from first tab
            driver.switch_to.window(tab_handles[0])
            company_name, currency = self.extract_company_info(driver)
            self.company_info = {'name': company_name, 'currency': currency}
            
            for idx, (key, name) in enumerate(statements):
                driver.switch_to.window(tab_handles[idx])
                dates, fields, selected = self.extract_data(driver, key)
                if dates and fields:
                    df = self.parse_data(dates, fields, selected)
                    if df is not None:
                        all_data[f"{name} ({selected})"] = df
            timer.stop('extract_default_data')
            
            # Now fetch Annual, Quarterly, TTM for each - parallel across tabs
            periods_config = {
                'income-statement': ['Annual', 'Quarterly', 'TTM'],
                'balance-sheet': ['Annual', 'Quarterly'],  # No TTM for balance sheet
                'cash-flow-statement': ['Annual', 'Quarterly', 'TTM']
            }
            
            # Fetch periods in waves - all Annual first, then all Quarterly, etc.
            all_periods = ['Annual', 'Quarterly', 'TTM']
            
            for period_idx, period in enumerate(all_periods):
                self.status(f"Fetching {period} data from all statements...", 25 + period_idx * 20)
                timer.start(f'fetch_{period.lower()}')
                
                # Click period buttons on all applicable tabs
                for idx, (key, name) in enumerate(statements):
                    if period in periods_config[key]:
                        sheet_key = f"{name} ({period})"
                        if sheet_key not in all_data:
                            driver.switch_to.window(tab_handles[idx])
                            self.click_period_fast(driver, period, key)
                
                # Wait for Livewire updates - need enough time for data to load
                time.sleep(2.0)  # Increased - Livewire needs time to fetch and render
                
                for idx, (key, name) in enumerate(statements):
                    if period in periods_config[key]:
                        sheet_key = f"{name} ({period})"
                        if sheet_key not in all_data:
                            driver.switch_to.window(tab_handles[idx])
                            dates, fields, selected = self.extract_data_livewire(driver, key)
                            if dates and fields:
                                df = self.parse_data(dates, fields, selected)
                                if df is not None:
                                    actual_key = f"{name} ({selected})"
                                    if actual_key not in all_data:
                                        all_data[actual_key] = df
                            else:
                                # Fallback to extract_data
                                dates, fields, selected = self.extract_data(driver, key)
                                if dates and fields:
                                    df = self.parse_data(dates, fields, selected)
                                    if df is not None:
                                        actual_key = f"{name} ({selected})"
                                        if actual_key not in all_data:
                                            all_data[actual_key] = df
                timer.stop(f'fetch_{period.lower()}')
                
                # Log what we've collected so far
                print(f"After {period}: {len(all_data)} sheets collected: {list(all_data.keys())}")
            
            # Fetch Revenue Breakdown (use first tab)
            self.status("Fetching Revenue Breakdown...", 75)
            timer.start('revenue_breakdown')
            driver.switch_to.window(tab_handles[0])
            breakdown_data = self.scrape_revenue_breakdown_fast(driver, base_url)
            timer.stop('revenue_breakdown')
            
            if all_data:
                self.status("Saving with formatting...", 85)
                timer.start('excel_save')
                
                # Pre-create format objects once for reuse
                self._workbook = None
                self._formats = {}
                
                with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                    # Define the desired sheet order (Balance Sheet TTM not available)
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
                            self.format_excel_sheet_optimized(writer, all_data[sheet_name], sheet_name[:31])
                    
                    # Write any remaining sheets that weren't in the order list
                    for sheet_name, df in all_data.items():
                        if sheet_name not in sheet_order:
                            self.format_excel_sheet_optimized(writer, df, sheet_name[:31])
                    
                    if breakdown_data:
                        self.format_revenue_breakdown_sheet(writer, breakdown_data)
                
                timer.stop('excel_save')
                timer.stop('total')
                
                total_sheets = len(all_data) + (1 if breakdown_data else 0)
                
                # Print performance summary
                print("\n=== Performance Summary ===")
                print(timer.get_summary())
                print("===========================\n")
                
                self.status(f"Done! Saved {total_sheets} sheets.", 100)
                self.root.after(0, lambda: messagebox.showinfo("Success", f"Saved to {self.output_file}\n\n{timer.get_summary()}"))
            else:
                self.status("No data found", 0)
                
        except Exception as e:
            self.status(f"Error: {e}", 0)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            import traceback
            traceback.print_exc()
        finally:
            if driver:
                driver.quit()
            self.root.after(0, lambda: self.fetch_btn.config(state=tk.NORMAL))
    
    def extract_company_info(self, driver):
        """Extract company name and currency from the page."""
        company_name = self.selected_ticker.upper()
        currency = 'USD'
        
        try:
            # Try to get company name from the page header
            header = driver.find_element(By.CSS_SELECTOR, '.security-header h1, .company-name, h1.title')
            if header:
                company_name = header.text.strip()
        except:
            pass
        
        try:
            # Try to get currency from the page
            currency_elem = driver.find_element(By.XPATH, "//*[contains(text(), 'Currency:')]") 
            if currency_elem:
                text = currency_elem.text
                if 'USD' in text:
                    currency = 'USD'
                elif 'EUR' in text:
                    currency = 'EUR'
                elif 'GBP' in text:
                    currency = 'GBP'
                else:
                    # Extract currency code
                    import re
                    match = re.search(r'Currency:\s*([A-Z]{3})', text)
                    if match:
                        currency = match.group(1)
        except:
            pass
        
        return company_name, currency
    
    def extract_data(self, driver, statement_type):
        try:
            elem = driver.find_element(By.CSS_SELECTOR, f'.{statement_type}.statement')
            wire = elem.get_attribute('wire:initial-data')
            if wire:
                data = json.loads(wire).get('serverMemo', {}).get('data', {})
                return data.get('dates', []), data.get('fieldsData', {}), data.get('selectedPeriod', 'Unknown')
        except:
            pass
        return None, None, None
    
    def extract_data_livewire(self, driver, statement_type):
        try:
            elem = driver.find_element(By.CSS_SELECTOR, f'.{statement_type}.statement')
            wire_id = elem.get_attribute('wire:id')
            if wire_id:
                script = f"""
                    var component = window.Livewire.find('{wire_id}');
                    if (component) {{
                        return JSON.stringify({{
                            dates: component.get('dates'),
                            fieldsData: component.get('fieldsData'),
                            selectedPeriod: component.get('selectedPeriod')
                        }});
                    }}
                    return null;
                """
                result = driver.execute_script(script)
                if result:
                    data = json.loads(result)
                    return data.get('dates', []), data.get('fieldsData', {}), data.get('selectedPeriod', 'Unknown')
        except:
            pass
        return None, None, None
    
    def click_period_fast(self, driver, period, statement_type):
        """Optimized version of click_period with minimal waits."""
        try:
            dropdown = driver.find_element(By.CSS_SELECTOR, f'.{statement_type}.statement .vperiod.dropdown')
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
            driver.execute_script("arguments[0].click();", dropdown)
            
            # Use WebDriverWait instead of fixed sleep
            try:
                WebDriverWait(driver, 2).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, f'.{statement_type}.statement .vperiod.dropdown .menu'))
                )
            except:
                time.sleep(0.5)  # Fallback
            
            menu = driver.find_element(By.CSS_SELECTOR, f'.{statement_type}.statement .vperiod.dropdown .menu')
            items = menu.find_elements(By.CSS_SELECTOR, '.item')
            
            # Try exact match first, then partial match
            for item in items:
                item_text = item.text.strip().lower()
                if period.lower() == item_text or period.lower() in item_text:
                    driver.execute_script("arguments[0].click();", item)
                    time.sleep(0.5)  # Brief wait after click for Livewire to start updating
                    return True
            
            print(f"Available periods for {statement_type}: {[i.text for i in items]}")
        except Exception as e:
            print(f"Error clicking period {period} for {statement_type}: {e}")
        return False
    
    def click_period(self, driver, period, statement_type):
        """Original click_period kept for backward compatibility."""
        return self.click_period_fast(driver, period, statement_type)
    
    def parse_data(self, dates, fields_data, period_type):
        date_labels, date_keys = [], []
        for d in dates:
            date_str = d.get('date', d)[:10] if isinstance(d, dict) else str(d)[:10]
            try:
                parsed = dt.strptime(date_str, '%Y-%m-%d')
                label = f"FY{parsed.year}" if period_type == 'Annual' else f"Q{(parsed.month-1)//3+1}/{parsed.year}"
                date_labels.append(label)
                date_keys.append(parsed)
            except:
                date_labels.append(date_str)
                date_keys.append(None)
        
        rows = []
        for group, items in fields_data.items():
            for item in items:
                row = {'Field': item.get('name', ''), '_Type': item.get('ingroupType', '')}
                unit = item.get('unit', 'usd')
                for i, val in enumerate(item.get('values', [])):
                    if i < len(date_labels):
                        v = val.get('value', 0)
                        if v and unit != 'usd_per_share' and 'EPS' not in item.get('name', ''):
                            v = v / 1_000_000
                        row[date_labels[i]] = v or 0
                rows.append(row)
        
        df = pd.DataFrame(rows)
        fixed = ['Field', '_Type']
        date_cols = [c for c in df.columns if c not in fixed]
        col_map = {date_labels[i]: date_keys[i] for i in range(len(date_labels)) if date_keys[i]}
        sorted_cols = sorted(date_cols, key=lambda x: col_map.get(x, dt.min))
        return df[fixed + sorted_cols]
    
    def scrape_revenue_breakdown_fast(self, driver, base_url):
        """Optimized Revenue Breakdown scraper with reduced waits."""
        try:
            driver.get(f"{base_url}/revenue-breakdown")
            
            # Use explicit wait instead of fixed sleep
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Breakdown by')]"))
                )
            except:
                time.sleep(2.5)  # Fallback - need enough time for page to load
            
            # Click Show More buttons - faster approach
            for _ in range(3):  # Reduced iterations
                try:
                    buttons = driver.find_elements(By.XPATH, "//*[contains(text(), 'Show More')]")
                    if not buttons:
                        break
                    for btn in buttons:
                        try:
                            if 'Show Less' not in btn.text:
                                driver.execute_script("arguments[0].click();", btn)
                        except:
                            pass
                    time.sleep(0.5)  # Allow content to expand
                except:
                    break
            
            page_text = driver.find_element(By.TAG_NAME, 'body').text
            
            breakdown_data = {}
            lines = page_text.split('\n')
            current_section = None
            current_items = []
            total_revenue = 0
            
            for line in lines:
                line = line.strip()
                
                if 'Breakdown by Geography' in line:
                    if current_section and current_items:
                        breakdown_data[current_section] = {'total': total_revenue, 'items': current_items}
                    current_section = 'Geography'
                    current_items = []
                    total_revenue = 0
                elif 'Breakdown by Segments' in line:
                    if current_section and current_items:
                        breakdown_data[current_section] = {'total': total_revenue, 'items': current_items}
                    current_section = 'Segments'
                    current_items = []
                    total_revenue = 0
                
                if not current_section:
                    continue
                if 'SEE ALSO' in line or 'Summary' in line:
                    break
                
                if 'Total Revenue:' in line:
                    match = re.search(r'([\d.]+)([BMK]?)\s*USD', line, re.I)
                    if match:
                        val = float(match.group(1))
                        mult = match.group(2).upper()
                        total_revenue = val * 1000 if mult == 'B' else val if mult == 'M' else val / 1000 if mult == 'K' else val * 1000
                    continue
                
                match = re.search(r'^(.+?):\s*([\d.]+)([BMK]?)\s*USD', line, re.I)
                if match:
                    name = match.group(1).strip()
                    val = float(match.group(2))
                    mult = match.group(3).upper()
                    if mult == 'B':
                        val *= 1000
                    elif mult == 'K':
                        val /= 1000
                    current_items.append({'name': name, 'value': val})
            
            if current_section and current_items:
                breakdown_data[current_section] = {'total': total_revenue, 'items': current_items}
            
            return breakdown_data
        except Exception as e:
            print(f"Revenue breakdown error: {e}")
            return {}
    
    def scrape_revenue_breakdown(self, driver, base_url):
        """Original method - redirects to fast version."""
        return self.scrape_revenue_breakdown_fast(driver, base_url)
    
    def get_formula_definitions(self, sheet_name):
        if 'Income Statement' in sheet_name:
            return INCOME_STATEMENT_FORMULAS
        elif 'Balance Sheet' in sheet_name:
            return BALANCE_SHEET_FORMULAS
        elif 'Cash Flow' in sheet_name:
            return CASH_FLOW_FORMULAS
        return {}
    
    def _get_formats(self, workbook, currency_symbol='$'):
        """Get or create cached format objects for the workbook."""
        # Cache key includes currency symbol to regenerate if currency changes
        cache_key = (id(workbook), currency_symbol)
        if not hasattr(self, '_format_cache_key') or self._format_cache_key != cache_key:
            self._format_cache_key = cache_key
            self._workbook = workbook
            
            # Build currency format string with the appropriate symbol
            curr_fmt = f'{currency_symbol}#,##0'
            
            self._formats = {
                'title': workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#2E75B6'}),
                'subtitle': workbook.add_format({'font_size': 11, 'font_color': '#666666'}),
                'header': workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1}),
                'group': workbook.add_format({'bold': True, 'bg_color': '#D9E2F3', 'border': 1}),
                'level1': workbook.add_format({'indent': 1, 'border': 1}),
                'currency': workbook.add_format({'num_format': curr_fmt, 'border': 1, 'align': 'right'}),
                'currency_bold': workbook.add_format({'num_format': curr_fmt, 'bold': True, 'bg_color': '#D9E2F3', 'border': 1, 'align': 'right'}),
                'eps': workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'align': 'right'}),
                'eps_bold': workbook.add_format({'num_format': '#,##0.00', 'bold': True, 'bg_color': '#D9E2F3', 'border': 1, 'align': 'right'}),
                'ratio_header': workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#2E75B6', 'font_color': 'white', 'border': 1}),
                'ratio_label': workbook.add_format({'bold': False, 'bg_color': '#DEEBF7', 'border': 1, 'indent': 1}),
                'ratio_pct': workbook.add_format({'num_format': '0.0%', 'bg_color': '#DEEBF7', 'border': 1, 'align': 'right'}),
            }
        return self._formats
    
    def format_excel_sheet_optimized(self, writer, df, sheet_name):
        """Optimized Excel sheet formatter with cached format objects."""
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet
        
        # Get company info
        company_info = getattr(self, 'company_info', {'name': self.selected_ticker, 'currency': 'USD'})
        currency_code = company_info.get('currency', 'USD')
        currency_symbol = get_currency_symbol(currency_code)
        
        # Get cached formats with proper currency symbol
        fmts = self._get_formats(workbook, currency_symbol)
        
        # Build full company display name: "TICKER - Company Name"
        ticker = self.selected_ticker.upper() if self.selected_ticker else ''
        full_name = getattr(self, 'selected_company_name', ticker)
        company_display = f"{ticker} - {full_name}" if full_name and full_name != ticker else ticker
        
        # Determine statement type from sheet name
        if 'Income Statement' in sheet_name:
            statement_type = 'Income Statement'
        elif 'Balance Sheet' in sheet_name:
            statement_type = 'Balance Sheet'
        elif 'Cash Flow' in sheet_name:
            statement_type = 'Cash Flow Statement'
        else:
            statement_type = sheet_name
        
        # Write company header section (rows 0-2)
        worksheet.write(0, 0, statement_type, fmts['title'])
        worksheet.write(1, 0, company_display, fmts['subtitle'])
        worksheet.write(2, 0, f"Currency: {currency_code}", fmts['subtitle'])
        
        # Data starts at row 4 (leaving row 3 as spacer)
        header_row = 4
        
        formula_defs = self.get_formula_definitions(sheet_name)
        
        # Build field to row map (adjusted for header offset)
        field_to_row = {}
        for idx, row in df.iterrows():
            field_to_row[row.get('Field', '')] = idx + header_row + 2  # +2 for header row and 1-indexing
        
        date_cols = [c for c in df.columns if c not in ['Field', '_Type']]
        output_cols = ['Field'] + date_cols
        
        # Write column headers
        for col_num, col_name in enumerate(output_cols):
            worksheet.write(header_row, col_num, col_name, fmts['header'])
        
        # Pre-compute row data to reduce DataFrame access overhead
        rows_data = df.to_dict('records')
        
        # Write data
        for row_num, row_dict in enumerate(rows_data):
            excel_row = row_num + header_row + 1  # Adjusted for header offset
            row_type = row_dict.get('_Type', '')
            field_name = row_dict.get('Field', '')
            is_group = row_type in ['group-total', 'important']
            is_eps = 'EPS' in field_name
            has_formula = field_name in formula_defs
            
            for col_num, col_name in enumerate(output_cols):
                value = row_dict.get(col_name)
                
                if col_name == 'Field':
                    fmt = fmts['group'] if is_group else (fmts['level1'] if row_type == 'level-1' else None)
                    worksheet.write(excel_row, col_num, value, fmt)
                else:
                    if is_eps:
                        fmt = fmts['eps_bold'] if is_group else fmts['eps']
                    else:
                        fmt = fmts['currency_bold'] if is_group else fmts['currency']
                    
                    if has_formula:
                        # Write Excel formula
                        formula_def = formula_defs[field_name]
                        col_letter = get_column_letter(col_num)
                        terms = []
                        for src in formula_def['sources']:
                            src_row = field_to_row.get(src)
                            if src_row:
                                terms.append(f"{col_letter}{src_row}")
                        if terms:
                            formula = "=SUM(" + ",".join(terms) + ")"
                            worksheet.write_formula(excel_row, col_num, formula, fmt)
                        else:
                            worksheet.write(excel_row, col_num, value if pd.notna(value) else 0, fmt)
                    else:
                        worksheet.write(excel_row, col_num, value if pd.notna(value) else 0, fmt)
        
        # Add ratios for Income Statement (not TTM)
        if 'Income Statement' in sheet_name and 'TTM' not in sheet_name:
            self.add_ratios_optimized(worksheet, workbook, df, field_to_row, output_cols, header_row)
        
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:ZZ', 15)
        worksheet.freeze_panes(header_row + 1, 1)  # Freeze below data headers
    
    def format_excel_sheet(self, writer, df, sheet_name):
        """Original method - redirects to optimized version."""
        return self.format_excel_sheet_optimized(writer, df, sheet_name)
    
    def add_ratios_optimized(self, worksheet, workbook, df, field_to_row, output_cols, header_row=4):
        """Optimized ratios section using cached formats."""
        last_row = len(df) + header_row + 1
        ratio_row = last_row + 2
        
        fmts = self._get_formats(workbook)
        
        # Header row
        worksheet.write(ratio_row, 0, 'Ratios', fmts['ratio_header'])
        for col in range(1, len(output_cols)):
            worksheet.write(ratio_row, col, '', fmts['ratio_header'])
        
        ratios = [
            ('Gross Profit Margin', 'Gross Profit', 'Revenue'),
            ('Operating Profit Margin', 'Operating Income', 'Revenue'),
            ('Net Profit Margin', 'Net Income (Common)', 'Revenue'),
            ('R&D as % of Revenue', 'Research & Development', 'Revenue'),
            ('SG&A as % of Revenue', 'Selling, General & Administrative', 'Revenue'),
        ]
        
        current_row = ratio_row + 1
        for ratio_name, num_field, denom_field in ratios:
            worksheet.write(current_row, 0, ratio_name, fmts['ratio_label'])
            num_row = field_to_row.get(num_field)
            denom_row = field_to_row.get(denom_field)
            
            if num_row and denom_row:
                for col in range(1, len(output_cols)):
                    col_letter = get_column_letter(col)
                    if 'R&D' in ratio_name or 'SG&A' in ratio_name:
                        formula = f"=ABS({col_letter}{num_row})/{col_letter}{denom_row}"
                    else:
                        formula = f"={col_letter}{num_row}/{col_letter}{denom_row}"
                    worksheet.write_formula(current_row, col, formula, fmts['ratio_pct'])
            else:
                for col in range(1, len(output_cols)):
                    worksheet.write(current_row, col, 'N/A', fmts['ratio_label'])
            current_row += 1
        
        # Revenue Y/Y Growth
        worksheet.write(current_row, 0, 'Revenue Y/Y Growth', fmts['ratio_label'])
        rev_row = field_to_row.get('Revenue')
        if rev_row:
            worksheet.write(current_row, 1, 'N/A', fmts['ratio_label'])
            for col in range(2, len(output_cols)):
                col_letter = get_column_letter(col)
                prev_letter = get_column_letter(col - 1)
                formula = f"=({col_letter}{rev_row}-{prev_letter}{rev_row})/{prev_letter}{rev_row}"
                worksheet.write_formula(current_row, col, formula, fmts['ratio_pct'])
    
    def add_ratios(self, worksheet, workbook, df, field_to_row, output_cols, header_row=4):
        """Original method - redirects to optimized version."""
        return self.add_ratios_optimized(worksheet, workbook, df, field_to_row, output_cols, header_row)
    
    def format_revenue_breakdown_sheet(self, writer, breakdown_data):
        """Format Revenue Breakdown sheet."""
        workbook = writer.book
        worksheet = workbook.add_worksheet('Revenue Breakdown')
        writer.sheets['Revenue Breakdown'] = worksheet
        
        # Get company info
        company_info = getattr(self, 'company_info', {'name': self.selected_ticker, 'currency': 'USD'})
        currency_code = company_info.get('currency', 'USD')
        currency_symbol = get_currency_symbol(currency_code)
        
        # Build full company display name: "TICKER - Company Name"
        ticker = self.selected_ticker.upper() if self.selected_ticker else ''
        full_name = getattr(self, 'selected_company_name', ticker)
        company_display = f"{ticker} - {full_name}" if full_name and full_name != ticker else ticker
        
        # Header formats
        company_title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#2E75B6'})
        subtitle_fmt = workbook.add_format({'font_size': 11, 'font_color': '#666666'})
        
        # Write company header section (rows 0-2)
        worksheet.write(0, 0, 'Revenue Breakdown', company_title_fmt)
        worksheet.write(1, 0, company_display, subtitle_fmt)
        worksheet.write(2, 0, f"Currency: {currency_code}", subtitle_fmt)
        
        # Data formats with proper currency symbol
        curr_fmt_str = f'{currency_symbol}#,##0'
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#2E75B6', 'font_color': 'white', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
        item_fmt = workbook.add_format({'border': 1, 'indent': 1})
        currency_fmt = workbook.add_format({'num_format': curr_fmt_str, 'border': 1, 'align': 'right'})
        pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'right'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E2F3', 'border': 1})
        total_curr_fmt = workbook.add_format({'bold': True, 'num_format': curr_fmt_str, 'bg_color': '#D9E2F3', 'border': 1, 'align': 'right'})
        total_pct_fmt = workbook.add_format({'bold': True, 'num_format': '0.0%', 'bg_color': '#D9E2F3', 'border': 1, 'align': 'right'})
        
        row = 4  # Start after header section
        for section_name, section_data in breakdown_data.items():
            worksheet.write(row, 0, f'Breakdown by {section_name}', title_fmt)
            worksheet.write(row, 1, '', title_fmt)
            worksheet.write(row, 2, '', title_fmt)
            row += 1
            
            worksheet.write(row, 0, 'Segment', header_fmt)
            worksheet.write(row, 1, 'Revenue (M)', header_fmt)
            worksheet.write(row, 2, '% of Total', header_fmt)
            row += 1
            
            total = section_data.get('total', 0)
            if total == 0:
                total = sum(item.get('value', 0) for item in section_data.get('items', []))
            
            for item in section_data.get('items', []):
                val = item.get('value', 0)
                pct = val / total if total > 0 else 0
                worksheet.write(row, 0, item.get('name', ''), item_fmt)
                worksheet.write(row, 1, val, currency_fmt)
                worksheet.write(row, 2, pct, pct_fmt)
                row += 1
            
            worksheet.write(row, 0, 'Total', total_fmt)
            worksheet.write(row, 1, total, total_curr_fmt)
            worksheet.write(row, 2, 1.0, total_pct_fmt)
            row += 2
        
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 12)
        worksheet.freeze_panes(5, 0)  # Freeze below first data header
    
    def open_folder(self):
        os.system(f'open "{os.path.dirname(__file__)}"')


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleFinanceGUI(root)
    root.mainloop()
