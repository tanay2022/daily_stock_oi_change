#!/usr/bin/env python3
"""
Fetch historical OI data for a specific stock using nselib.
Configure the stock symbol at the top of the script.
Automatically fetches last 2 months of data.
Output will be saved as Excel file with stock name and date in filename.

NOTE: This script fetches data DATE-BY-DATE (like the working NIFTY script)
and uses the NEAREST expiry to ensure we get actual OI values from nselib.
"""

import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
from pathlib import Path

# ============================================================================
# CONFIGURATION - EDIT THESE VALUES
# ============================================================================
STOCK_SYMBOL = "MARICO"  # Change this to the stock you want to analyze

# Automatically calculate last 2 months
END_DATE = datetime.now().strftime("%Y-%m-%d")
START_DATE = (datetime.now() - timedelta(days=60)).strftime("%Y-%m-%d")
# ============================================================================

# Script directory
SCRIPT_DIR = Path(__file__).resolve().parent

# Try to import nselib
try:
    from nselib import derivatives
    NSELIB_AVAILABLE = True
except ImportError:
    print("ERROR: nselib not available. Please install it using: pip install nselib")
    NSELIB_AVAILABLE = False
    exit(1)


def get_stock_price_history(symbol, start_date, end_date):
    """Fetch stock price history from Yahoo Finance."""
    try:
        # Try NSE symbol format first
        ticker_symbol = f"{symbol}.NS"
        ticker = yf.Ticker(ticker_symbol)
        hist = ticker.history(start=start_date, end=(datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d"))
        
        if hist.empty:
            # Try BSE symbol format
            ticker_symbol = f"{symbol}.BO"
            ticker = yf.Ticker(ticker_symbol)
            hist = ticker.history(start=start_date, end=(datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d"))
        
        if hist.empty:
            print(f"Warning: No price data found for {symbol} on Yahoo Finance")
            return pd.DataFrame()
        
        hist = hist.reset_index()
        print(f"Fetched price history for {symbol}: {len(hist)} days")
        return hist
    
    except Exception as e:
        print(f"Error fetching price history for {symbol}: {e}")
        return pd.DataFrame()


def determine_strike_interval_from_data(df):
    """Auto-detect strike price interval from option chain data."""
    if df.empty:
        return 50  # Default fallback
    
    strikes = sorted(df['STRIKE_PRICE'].dropna().unique())
    if len(strikes) < 2:
        return 50
    
    differences = []
    for i in range(1, min(len(strikes), 20)):  # Check first 20 strikes
        diff = strikes[i] - strikes[i-1]
        if diff > 0:
            differences.append(diff)
    
    if not differences:
        return 50
    
    # Return the most common difference
    from collections import Counter
    counter = Counter(differences)
    most_common_interval = counter.most_common(1)[0][0]
    
    return most_common_interval


def fetch_historical_oi_data_stock(symbol, start_date, end_date):
    """
    Fetch historical OI data for a stock using nselib.
    Fetches data DATE-BY-DATE (like NIFTY script) to get actual OI values.
    Returns DataFrame with date, CE_OI, PE_OI, CE_CHG_OI, PE_CHG_OI, OI_%, CH_OI_%
    """
    print(f"\nFetching historical OI data for {symbol} from {start_date} to {end_date}...")
    print("Fetching data date-by-date (this may take a while)...")
    print()
    
    # Fetch stock price history for underlying values
    hist = get_stock_price_history(symbol, start_date, end_date)
    
    records = []
    current = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    
    strike_interval = None
    
    while current <= end_dt:
        # Skip weekends
        if current.weekday() >= 5:
            current += timedelta(days=1)
            continue
        
        date_str = current.strftime("%Y-%m-%d")
        from_date = current.strftime('%d-%m-%Y')
        to_date = (current + timedelta(days=1)).strftime('%d-%m-%Y')
        
        try:
            # Get stock close price for the day
            stock_close = None
            if not hist.empty:
                stock_row = hist[hist['Date'].dt.date == current.date()]
                if stock_row.empty:
                    print(f"  {date_str}: No price data (likely holiday)")
                    current += timedelta(days=1)
                    continue
                stock_close = stock_row['Close'].iloc[0]
            
            # Fetch options data from nselib FOR THIS SPECIFIC DATE
            print(f"  {date_str}: Fetching option chain...", end=" ")
            df = derivatives.option_price_volume_data(
                symbol=symbol,
                instrument="OPTSTK",
                from_date=from_date,
                to_date=to_date
            )
            
            if df.empty:
                print("No data")
                current += timedelta(days=1)
                continue
            
            # Filter for the specific date
            df = df[df['TIMESTAMP'] == current.strftime('%d-%b-%Y')]
            if df.empty:
                print("No data for this date")
                current += timedelta(days=1)
                continue
            
            # Parse data types
            df['EXPIRY_DT'] = pd.to_datetime(df['EXPIRY_DT'], format="%d-%b-%Y")
            df['STRIKE_PRICE'] = pd.to_numeric(df['STRIKE_PRICE'], errors='coerce').round(2)
            df['OPEN_INT'] = pd.to_numeric(df['OPEN_INT'], errors='coerce').fillna(0).round().astype(int)
            df['CHANGE_IN_OI'] = pd.to_numeric(df['CHANGE_IN_OI'], errors='coerce').fillna(0).round().astype(int)
            
            # Get underlying value from option chain if not available from Yahoo
            if stock_close is None and 'UNDERLYING_VALUE' in df.columns:
                underlying_values = pd.to_numeric(df['UNDERLYING_VALUE'], errors='coerce').dropna()
                if not underlying_values.empty:
                    stock_close = underlying_values.iloc[0]
            
            if stock_close is None:
                print("Cannot determine underlying price")
                current += timedelta(days=1)
                continue
            
            # Determine strike interval on first successful fetch
            if strike_interval is None:
                strike_interval = determine_strike_interval_from_data(df)
                print(f"\n  Detected strike interval: {strike_interval}")
            
            # Calculate ATM strike - round to 2 decimals
            atm_strike = round(round(stock_close / strike_interval) * strike_interval, 2)
            
            # Get NEAREST expiry (first valid expiry after current date) - matching NIFTY script
            unique_expiries = sorted(df['EXPIRY_DT'].unique())
            nearest_expiry = None
            for exp in unique_expiries:
                if exp.date() > current.date():
                    nearest_expiry = exp
                    break
            
            if nearest_expiry is None:
                print("No valid expiry")
                current += timedelta(days=1)
                continue
            
            # Filter for nearest expiry
            df_expiry = df[df['EXPIRY_DT'] == nearest_expiry]
            
            # Calculate strikes to sum (ATM ± 3 strikes) - round to 2 decimals
            strikes_to_sum = [
                round(atm_strike - (3 * strike_interval), 2),
                round(atm_strike - (2 * strike_interval), 2),
                round(atm_strike - strike_interval, 2),
                round(atm_strike, 2),
                round(atm_strike + strike_interval, 2),
                round(atm_strike + (2 * strike_interval), 2),
                round(atm_strike + (3 * strike_interval), 2)
            ]
            
            # Filter CE and PE data for the selected strikes
            df_ce = df_expiry[
                (df_expiry['OPTION_TYPE'] == 'CE') & 
                (df_expiry['STRIKE_PRICE'].isin(strikes_to_sum))
            ]
            df_pe = df_expiry[
                (df_expiry['OPTION_TYPE'] == 'PE') & 
                (df_expiry['STRIKE_PRICE'].isin(strikes_to_sum))
            ]
            
            # Sum OI values
            ce_oi_sum = df_ce['OPEN_INT'].sum()
            ce_chg_oi_sum = df_ce['CHANGE_IN_OI'].sum()
            pe_oi_sum = df_pe['OPEN_INT'].sum()
            pe_chg_oi_sum = df_pe['CHANGE_IN_OI'].sum()
            
            print(f"Price: {stock_close:.2f}, ATM: {atm_strike}, CE OI: {ce_oi_sum:,}, PE OI: {pe_oi_sum:,}")
            
            records.append({
                "Date": date_str,
                "Underlying_Price": round(stock_close, 2),
                "ATM_Strike": atm_strike,
                "CE_OI": ce_oi_sum,
                "PE_OI": pe_oi_sum,
                "CE_CHG_OI": ce_chg_oi_sum,
                "PE_CHG_OI": pe_chg_oi_sum,
            })
            
        except Exception as e:
            print(f"[ERROR] {date_str}: {e}")
        
        current += timedelta(days=1)
    
    if not records:
        return pd.DataFrame()
    
    # Create DataFrame
    df_result = pd.DataFrame(records)
    
    # Calculate OI metrics
    # OI_% = (PE_OI - CE_OI) / MIN(PE_OI, CE_OI) * 100
    min_oi = np.minimum(df_result['PE_OI'], df_result['CE_OI'])
    min_oi = min_oi.replace(0, np.nan)  # Avoid division by zero
    df_result['OI_%'] = ((df_result['PE_OI'] - df_result['CE_OI']) / min_oi * 100).round(2)
    
    # CH_OI_% = (PE_CHG_OI - CE_CHG_OI) / (CE_OI + PE_OI) * 100
    total_oi = df_result['CE_OI'] + df_result['PE_OI']
    total_oi = total_oi.replace(0, np.nan)  # Avoid division by zero
    df_result['CH_OI_%'] = ((df_result['PE_CHG_OI'] - df_result['CE_CHG_OI']) / total_oi * 100).round(2)
    
    return df_result


def main():
    """Main execution function."""
    print("=" * 70)
    print("Stock Historical OI Data Fetcher (Date-by-Date)")
    print("=" * 70)
    print()
    print(f"Stock Symbol: {STOCK_SYMBOL}")
    print(f"Date Range: {START_DATE} to {END_DATE}")
    print()
    
    if not NSELIB_AVAILABLE:
        print("ERROR: nselib is required but not available")
        return
    
    # Fetch historical OI data
    df = fetch_historical_oi_data_stock(STOCK_SYMBOL, START_DATE, END_DATE)
    
    if df.empty:
        print()
        print("=" * 70)
        print("No data collected. No output file created.")
        print("=" * 70)
        return
    
    # Generate output filename
    output_file = SCRIPT_DIR / f"{STOCK_SYMBOL}_OI_historical_{START_DATE}_to_{END_DATE}.xlsx"
    
    # Export to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print()
    print("=" * 70)
    print("Summary")
    print("=" * 70)
    print(f"Total days processed: {len(df)}")
    print(f"Output file: {output_file}")
    print()
    print("Sample data (first 10 rows):")
    print(df.head(10).to_string(index=False))
    print()
    print("[DONE] Script completed successfully!")


if __name__ == "__main__":
    main()
