#!/usr/bin/env python3
"""
Fetch OI data for all stocks in FNO stock list.
For each stock, calculate OI metrics for 7 strikes (ATM ± 3 strikes).
Export results to Excel file with date in filename.
"""

import pandas as pd
import numpy as np
import requests
from datetime import datetime
from pathlib import Path
import time
import json
import pytz

# Script directory
SCRIPT_DIR = Path(__file__).resolve().parent

# NSE Option Chain API configuration
NSE_OC_PAGE = "https://www.nseindia.com/option-chain"
NSE_CONTRACT_INFO_API = "https://www.nseindia.com/api/option-chain-contract-info"
NSE_OC_V3_API = "https://www.nseindia.com/api/option-chain-v3"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Referer": "https://www.nseindia.com/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "X-Requested-With": "XMLHttpRequest",
}

# Helper function to load environment values from .env file
def load_env_values(path: Path, keys: List[str]) -> Dict[str, str]:
    """Load environment values from .env file."""
    if not path.exists():
        return {}
    result: Dict[str, str] = {}
    try:
        for line in path.read_text().splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "=" not in stripped:
                continue
            key, val = stripped.split("=", 1)
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key in keys:
                result[key] = val
    except Exception as exc:
        print(f"Warning: failed to load {path}: {exc}")
    return result


# Telegram Configuration
import os
from typing import Dict, List

# Load from .env file first (for local execution)
_env_values = load_env_values(SCRIPT_DIR / ".env", ["TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID"])

TELEGRAM_BOT_TOKEN = ""
TELEGRAM_CHAT_ID = ""

if _env_values.get("TELEGRAM_BOT_TOKEN"):
    TELEGRAM_BOT_TOKEN = _env_values["TELEGRAM_BOT_TOKEN"]
if _env_values.get("TELEGRAM_CHAT_ID"):
    TELEGRAM_CHAT_ID = _env_values["TELEGRAM_CHAT_ID"]

# Re-check environment variables (important for Vercel deployment)
if not TELEGRAM_BOT_TOKEN and os.environ.get("TELEGRAM_BOT_TOKEN"):
    TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
if not TELEGRAM_CHAT_ID and os.environ.get("TELEGRAM_CHAT_ID"):
    TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

TELEGRAM_ENABLED = bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)

if not TELEGRAM_ENABLED:
    print("Telegram notifications disabled: missing bot token or chat ID")
else:
    masked_chat = TELEGRAM_CHAT_ID if len(TELEGRAM_CHAT_ID) <= 4 else f"***{TELEGRAM_CHAT_ID[-4:]}"
    print(f"Telegram notifications enabled (chat {masked_chat})")

# Input/Output file paths
FNO_STOCK_LIST_FILE = SCRIPT_DIR / "fno_stock_list.xlsx"


def read_fno_stocks():
    """Read stock symbols from the FNO stock list Excel file."""
    try:
        df = pd.read_excel(FNO_STOCK_LIST_FILE)
        
        # Look for a column named 'Symbol' (case-insensitive)
        symbol_col = None
        for col in df.columns:
            if col.strip().lower() == 'symbol':
                symbol_col = col
                break
        
        if symbol_col is None:
            print(f"Error: No 'Symbol' column found in {FNO_STOCK_LIST_FILE}")
            print(f"Available columns: {df.columns.tolist()}")
            return []
        
        symbols = df[symbol_col].dropna().tolist()
        print(f"Loaded {len(symbols)} stock symbols from {FNO_STOCK_LIST_FILE}")
        return symbols
    
    except Exception as e:
        print(f"Error reading FNO stock list: {e}")
        return []


def determine_strike_interval(strikes_list):
    """
    Auto-detect strike price interval from a list of strikes.
    Returns the most common difference between consecutive strikes.
    """
    if len(strikes_list) < 2:
        return 50  # Default fallback
    
    strikes_sorted = sorted(strikes_list)
    differences = []
    
    for i in range(1, len(strikes_sorted)):
        diff = strikes_sorted[i] - strikes_sorted[i-1]
        if diff > 0:
            differences.append(diff)
    
    if not differences:
        return 50  # Default fallback
    
    # Return the most common difference
    from collections import Counter
    counter = Counter(differences)
    most_common_interval = counter.most_common(1)[0][0]
    
    return most_common_interval


def calculate_atm_strike(price, interval):
    """Round price to nearest strike based on interval."""
    return round(price / interval) * interval


def get_valid_strikes(atm_strike, interval, count=3):
    """
    Generate list of strikes: ATM ± count strikes.
    Returns a list of 2*count + 1 strikes (e.g., 7 strikes for count=3).
    """
    strikes = []
    for i in range(-count, count + 1):
        strikes.append(atm_strike + (i * interval))
    return strikes


def fetch_stock_oi_data(symbol, session=None):
    """
    Fetch OI data for a single stock from NSE option chain API using new v3 endpoints.
    Returns a dictionary with symbol, underlying value, and OI sums.
    """
    try:
        print(f"Fetching OI data for {symbol}...")
        
        # Create session if not provided
        if session is None:
            session = requests.Session()
            session.headers.update(HEADERS)
        
        # Step 1: Get contract info to find available expiries
        print(f"  Getting contract info for {symbol}...")
        contract_response = session.get(
            NSE_CONTRACT_INFO_API, 
            params={"symbol": symbol}, 
            timeout=15
        )
        contract_response.raise_for_status()
        contract_data = contract_response.json()
        
        print(f"  Contract response status: {contract_response.status_code}")
        print(f"  Contract data keys: {list(contract_data.keys()) if contract_data else 'Empty contract data'}")
        
        # Extract expiry dates from contract info
        expiry_dates = contract_data.get("expiryDates", [])
        if not expiry_dates:
            print(f"  [ERROR] No expiry dates found in contract info for {symbol}")
            print(f"  Contract response: {contract_response.text[:500]}...")
            return None
        
        print(f"  Available expiries: {len(expiry_dates)} dates")
        print(f"  First few expiries: {expiry_dates[:3]}")
        
        # Get nearest expiry
        parsed_expiries = []
        for expiry in expiry_dates:
            try:
                expiry_dt = datetime.strptime(expiry, "%d-%b-%Y")
                parsed_expiries.append((expiry_dt, expiry))
            except ValueError as e:
                print(f"  Warning: Could not parse expiry date '{expiry}': {e}")
                continue
        
        if not parsed_expiries:
            print(f"  [ERROR] No valid expiry dates could be parsed for {symbol}")
            return None
        
        
        parsed_expiries.sort()
        # Use IST timezone for correct date comparison on Vercel (which runs in UTC)
        ist_tz = pytz.timezone('Asia/Kolkata')
        today = datetime.now(ist_tz).date()
        selected_expiry = None
        
        for expiry_dt, expiry_str in parsed_expiries:
            if expiry_dt.date() > today:
                selected_expiry = expiry_str
                break
        
        if selected_expiry is None:
            selected_expiry = parsed_expiries[0][1]
        
        print(f"  Selected expiry: {selected_expiry}")
        
        # Step 2: Get option chain data using v3 API
        print(f"  Getting option chain data for {symbol} with expiry {selected_expiry}...")
        params = {
            "type": "Equity",
            "symbol": symbol,
            "expiry": selected_expiry
        }
        
        api_response = session.get(NSE_OC_V3_API, params=params, timeout=15)
        api_response.raise_for_status()
        payload = api_response.json()
        
        # Debug: Print the actual payload structure
        print(f"  API Response status: {api_response.status_code}")
        print(f"  Payload keys: {list(payload.keys()) if payload else 'Empty payload'}")
        
        # The v3 API structure might be different, let's handle both possible structures
        records_section = payload.get("records", {})
        if not records_section:
            # Try direct data access if records section is missing
            raw_records = payload.get("data", [])
            underlying_value = payload.get("underlyingValue")
        else:
            raw_records = records_section.get("data", [])
            underlying_value = records_section.get("underlyingValue")
        
        print(f"  Records section keys: {list(records_section.keys()) if records_section else 'Empty records'}")
        print(f"  Raw records count: {len(raw_records)}")
        
        if not raw_records:
            print(f"  [ERROR] No option chain data available for {symbol}")
            print(f"  Full API response for debugging:")
            print(f"  Response text: {api_response.text[:500]}...")
            return None
        
        # Get underlying value
        if underlying_value is None:
            # Fallback: try to get from first record
            underlying_value = float(raw_records[0].get("CE", {}).get("underlyingValue", 0))
            if underlying_value == 0:
                underlying_value = float(raw_records[0].get("PE", {}).get("underlyingValue", 0))
        
        underlying_value = float(underlying_value)
        
        # Determine strike interval from available strikes
        all_strikes = [record.get("strikePrice") for record in raw_records if record.get("strikePrice")]
        strike_interval = determine_strike_interval(all_strikes)
        
        # Calculate ATM strike
        atm_strike = calculate_atm_strike(underlying_value, strike_interval)
        
        # Get 7 valid strikes (ATM ± 3)
        strikes_to_sum = get_valid_strikes(atm_strike, strike_interval, count=3)
        
        print(f"  Underlying: {underlying_value:.2f}, ATM: {atm_strike}, Interval: {strike_interval}")
        print(f"  Strikes to sum: {strikes_to_sum}")
        
        # Sum OI for the selected strikes
        ce_oi_sum = 0
        pe_oi_sum = 0
        ce_chg_oi_sum = 0
        pe_chg_oi_sum = 0
        
        for record in raw_records:
            strike_price = record.get("strikePrice")
            if strike_price not in strikes_to_sum:
                continue
            
            ce_leg = record.get("CE")
            pe_leg = record.get("PE")
            
            if ce_leg:
                ce_oi_sum += float(ce_leg.get("openInterest", 0))
                ce_chg_oi_sum += float(ce_leg.get("changeinOpenInterest", 0))
            
            if pe_leg:
                pe_oi_sum += float(pe_leg.get("openInterest", 0))
                pe_chg_oi_sum += float(pe_leg.get("changeinOpenInterest", 0))
        
        print(f"  [OK] CE OI: {ce_oi_sum:,.0f}, PE OI: {pe_oi_sum:,.0f}")
        
        return {
            "Symbol": symbol,
            "Underlying_Value": underlying_value,
            "Sum_CE_OI": int(ce_oi_sum),
            "Sum_PE_OI": int(pe_oi_sum),
            "Sum_CE_Change_OI": int(ce_chg_oi_sum),
            "Sum_PE_Change_OI": int(pe_chg_oi_sum),
        }
        
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            print(f"  [ERROR] {symbol}: No option chain data available (404)")
        else:
            print(f"  [ERROR] {symbol}: HTTP error {e.response.status_code}")
        return None
    except Exception as e:
        print(f"  [ERROR] Error fetching OI data for {symbol}: {e}")
        return None


def send_telegram_message(message: str) -> bool:
    """Send message to Telegram."""
    if not TELEGRAM_ENABLED:
        print("Telegram not enabled, skipping message send")
        return False
    
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message,
            "parse_mode": "HTML"
        }
        
        print("Attempting to send Telegram message...")
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        result = response.json()
        ok_flag = result.get("ok", True)
        print(f"Telegram message sent successfully (ok={ok_flag})")
        return True
    except Exception as e:
        error_text = getattr(e, 'response', None)
        extra = ''
        if error_text is not None and hasattr(error_text, 'text'):
            extra = f" | response={error_text.text}"
        print(f"Failed to send Telegram message: {e}{extra}")
        return False


def format_telegram_message(top_10_df: pd.DataFrame, total_count: int, date_str: str) -> str:
    """Format top 10 stocks for Telegram message."""
    lines = []
    lines.append("📊 <b>Stock OI Analysis - Top 10</b>")
    lines.append(f"Date: {date_str}")
    lines.append(f"Total Stocks: {total_count}")
    lines.append("")
    lines.append("<b>Top 10 by Combined CH_OI:</b>")
    lines.append("")
    
    for idx, row in top_10_df.iterrows():
        symbol = row['Symbol']
        combined_oi = row['Combined_OI']
        combined_ch_oi = row['Combined_CH_OI']
        
        # Format with sign
        oi_str = f"{combined_oi:+.4f}"
        ch_oi_str = f"{combined_ch_oi:+.4f}"
        
        lines.append(f"<b>{symbol}</b>")
        lines.append(f"  OI: {oi_str} | CH_OI: {ch_oi_str}")
    
    lines.append("")
    lines.append("Powered by NSE Stock OI Tracker")
    
    return "\n".join(lines)


def vercel_handler():
    """Vercel API handler function."""
    try:
        print("=== Vercel Handler Started ===")
        
        # Read stock symbols
        symbols = read_fno_stocks()
        if not symbols:
            return {
                "statusCode": 500,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": json.dumps({
                    "success": False,
                    "error": "No symbols to process",
                    "timestamp": datetime.now(pytz.timezone('Asia/Kolkata')).isoformat()
                }, indent=2)
            }
        
        print(f"Processing {len(symbols)} stocks...")
        
        # Create a persistent session for all requests
        session = requests.Session()
        session.headers.update(HEADERS)
        
        # Warm-up request to NSE
        try:
            session.get(NSE_OC_PAGE, timeout=10)
        except:
            pass
        
        # Fetch OI data for each stock
        results = []
        failed_symbols = []
        
        for i, symbol in enumerate(symbols, 1):
            print(f"[{i}/{len(symbols)}] {symbol}")
            oi_data = fetch_stock_oi_data(symbol, session=session)
            
            if oi_data:
                results.append(oi_data)
            else:
                failed_symbols.append(symbol)
        
        # Create DataFrame
        if not results:
            return {
                "statusCode": 500,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": json.dumps({
                    "success": False,
                    "error": "No data collected",
                    "failed_symbols": failed_symbols,
                    "timestamp": datetime.now(pytz.timezone('Asia/Kolkata')).isoformat()
                }, indent=2)
            }
        
        df = pd.DataFrame(results)
        
        # Calculate Combined_OI = (Sum_PE_OI - Sum_CE_OI) / MIN(Sum_PE_OI, Sum_CE_OI)
        min_oi = np.minimum(df['Sum_PE_OI'], df['Sum_CE_OI'])
        min_oi = min_oi.replace(0, np.nan)
        df['Combined_OI'] = ((df['Sum_PE_OI'] - df['Sum_CE_OI']) / min_oi).round(4)
        
        # Calculate Combined_CH_OI = (Sum_PE_Change_OI - Sum_CE_Change_OI) / (Sum_PE_OI + Sum_CE_OI)
        total_oi = df['Sum_PE_OI'] + df['Sum_CE_OI']
        total_oi = total_oi.replace(0, np.nan)
        df['Combined_CH_OI'] = ((df['Sum_PE_Change_OI'] - df['Sum_CE_Change_OI']) / total_oi).round(4)
        
        # Sort by Combined_CH_OI descending
        df_sorted = df.sort_values('Combined_CH_OI', ascending=False).reset_index(drop=True)
        
        # Get top 10 for Telegram
        top_10 = df_sorted.head(10)[['Symbol', 'Combined_OI', 'Combined_CH_OI']].copy()
        
        # Get current date (IST)
        ist_tz = pytz.timezone('Asia/Kolkata')
        today_str = datetime.now(ist_tz).strftime("%Y-%m-%d")
        
        # Send Telegram message with top 10
        telegram_sent = False
        if TELEGRAM_ENABLED:
            telegram_message = format_telegram_message(top_10, len(df_sorted), today_str)
            telegram_sent = send_telegram_message(telegram_message)
        
        # Prepare full JSON response
        response_data = {
            "success": True,
            "telegram_sent": telegram_sent,
            "date": today_str,
            "summary": {
                "total_stocks": len(symbols),
                "successful": len(results),
                "failed": len(failed_symbols),
                "failed_symbols": failed_symbols
            },
            "data": df_sorted.to_dict('records'),
            "timestamp": datetime.now(ist_tz).isoformat()
        }
        
        return {
            "statusCode": 200,
            "headers": {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                "Access-Control-Allow-Headers": "Content-Type"
            },
            "body": json.dumps(response_data, indent=2)
        }
        
    except Exception as e:
        print(f"Error in vercel_handler: {e}")
        import traceback
        traceback.print_exc()
        
        error_response = {
            "success": False,
            "error": str(e),
            "timestamp": datetime.now(pytz.timezone('Asia/Kolkata')).isoformat()
        }
        
        return {
            "statusCode": 500,
            "headers": {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*"
            },
            "body": json.dumps(error_response, indent=2)
        }


def main():
    """Main execution function."""
    # Record start time
    start_time = time.time()
    
    print("=" * 60)
    print("Stock OI Data Fetcher")
    print("=" * 60)
    print()
    
    # Read stock symbols
    symbols = read_fno_stocks()
    if not symbols:
        print("No symbols to process. Exiting.")
        return
    
    print(f"\nProcessing {len(symbols)} stocks...")
    print()
    
    # Create a persistent session for all requests
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # Fetch OI data for each stock
    results = []
    failed_symbols = []
    
    for i, symbol in enumerate(symbols, 1):
        print(f"[{i}/{len(symbols)}] {symbol}")
        
        oi_data = fetch_stock_oi_data(symbol, session=session)
        
        if oi_data:
            results.append(oi_data)
        else:
            failed_symbols.append(symbol)
        
        print()
    
    # Create DataFrame
    if results:
        df = pd.DataFrame(results)
        
        # Calculate Combined_OI = (Sum_PE_OI - Sum_CE_OI) / MIN(Sum_PE_OI, Sum_CE_OI)
        min_oi = np.minimum(df['Sum_PE_OI'], df['Sum_CE_OI'])
        min_oi = min_oi.replace(0, np.nan)  # Avoid division by zero
        df['Combined_OI'] = ((df['Sum_PE_OI'] - df['Sum_CE_OI']) / min_oi).round(4)
        
        # Calculate Combined_CH_OI = (Sum_PE_Change_OI - Sum_CE_Change_OI) / (Sum_PE_OI + Sum_CE_OI)
        total_oi = df['Sum_PE_OI'] + df['Sum_CE_OI']
        total_oi = total_oi.replace(0, np.nan)  # Avoid division by zero
        df['Combined_CH_OI'] = ((df['Sum_PE_Change_OI'] - df['Sum_CE_Change_OI']) / total_oi).round(4)
        
        # Sort by Combined_CH_OI descending for display
        df_sorted = df.sort_values('Combined_CH_OI', ascending=False).reset_index(drop=True)
        
        # Generate output filename with today's date (IST)
        ist_tz = pytz.timezone('Asia/Kolkata')
        today_str = datetime.now(ist_tz).strftime("%Y-%m-%d")
        
        # Only save Excel file when running locally (not on Vercel)
        is_vercel = os.environ.get('VERCEL', '') == '1'
        output_file = None
        
        if not is_vercel:
            output_file = SCRIPT_DIR / f"stock_OI_data_{today_str}.xlsx"
            # Export to Excel
            df_sorted.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Excel file saved: {output_file}")
        else:
            print("Running on Vercel - Excel file output skipped")
        
        # Send Telegram message with top 10 (if enabled)
        telegram_sent = False
        if TELEGRAM_ENABLED:
            top_10 = df_sorted.head(10)[['Symbol', 'Combined_OI', 'Combined_CH_OI']].copy()
            telegram_message = format_telegram_message(top_10, len(df_sorted), today_str)
            telegram_sent = send_telegram_message(telegram_message)
        
        print("=" * 60)
        print("Summary")
        print("=" * 60)
        print(f"Successfully processed: {len(results)} stocks")
        print(f"Failed: {len(failed_symbols)} stocks")
        if failed_symbols:
            print(f"Failed symbols: {', '.join(failed_symbols)}")
        print()
        if output_file:
            print(f"Output file: {output_file}")
        if telegram_sent:
            print("Telegram message sent successfully")
        print()
        print("Sample data (Top 10 by Combined_CH_OI):")
        print(df_sorted.head(10).to_string(index=False))
        print()
        
        # Calculate and print total execution time
        end_time = time.time()
        total_time = end_time - start_time
        print(f"Total execution time: {total_time:.2f} seconds")
        print("[DONE] Script completed successfully!")
    else:
        print("=" * 60)
        print("No data collected. No output file created.")
        print("=" * 60)
        
        # Calculate and print total execution time even for failed case
        end_time = time.time()
        total_time = end_time - start_time
        print(f"Total execution time: {total_time:.2f} seconds")


if __name__ == "__main__":
    main()
