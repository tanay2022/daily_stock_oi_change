# Stock OI Analysis (Vercel Ready)

Fetches option chain Open Interest (OI) data for all F&O stocks from NSE, calculates Combined OI and Combined Change in OI metrics, and provides both JSON API and Telegram notifications.

## Features

- Fetches OI data for all stocks in the F&O stock list
- Calculates OI metrics for 7 strikes around ATM (ATM ± 3 strikes)
- Auto-detects strike price intervals for each stock
- Computes Combined_OI and Combined_CH_OI metrics
- Returns complete data as JSON, sorted by Combined_CH_OI (descending)
- Sends top 10 stocks to Telegram (Symbol, Combined_OI, Combined_CH_OI)
- Timezone-aware (uses IST) for correct date handling on Vercel

## Project Layout

```
stock_OI_change/
├── api/
│   └── stock-oi.py          # Vercel handler
├── stock_daily_OI_change.py # Core OI fetching and analysis
├── fno_stock_list.xlsx      # List of F&O stock symbols
├── requirements.txt         # Python dependencies
├── vercel.json             # Vercel configuration
├── .env.example            # Environment variable template
├── .gitignore              # Git ignore rules
└── README.md
```

## Environment Variables

Create a `.env` file or configure in Vercel → *Settings → Environment Variables*:

```
TELEGRAM_BOT_TOKEN=123456789:ABC...
TELEGRAM_CHAT_ID=987654321
```

**Note**: Telegram is optional. If not configured, the API will still return JSON data without sending notifications.

## Local Usage

### 1. Install Dependencies

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # Linux/Mac

pip install -r requirements.txt
```

### 2. Run the Script Locally

```bash
python stock_daily_OI_change.py
```

This will:
- Fetch OI data for all stocks in `fno_stock_list.xlsx`
- Calculate Combined_OI and Combined_CH_OI metrics
- Save results to `stock_OI_data_{date}.xlsx`
- Print summary to console

### 3. Test the Vercel Handler

```python
from stock_daily_OI_change import vercel_handler
response = vercel_handler()
print(response["body"])
```

## Vercel Deployment

### 1. Prerequisites

- Install Vercel CLI: `npm install -g vercel`
- Have a Vercel account

### 2. Deploy to Vercel

```bash
cd stock_OI_change
vercel login
vercel deploy --prod
```

### 3. Configure Environment Variables

In Vercel Dashboard → Settings → Environment Variables, add:
- `TELEGRAM_BOT_TOKEN` (optional)
- `TELEGRAM_CHAT_ID` (optional)

### 4. API Endpoint

After deployment, your API will be available at:
```
https://<your-project>.vercel.app/api/stock-oi
```

Example call:
```bash
curl https://<your-project>.vercel.app/api/stock-oi
```

## API Response Format

```json
{
  "success": true,
  "telegram_sent": true,
  "date": "2025-12-09",
  "summary": {
    "total_stocks": 183,
    "successful": 180,
    "failed": 3,
    "failed_symbols": ["STOCK1", "STOCK2"]
  },
  "data": [
    {
      "Symbol": "RELIANCE",
      "Underlying_Value": 2850.50,
      "Sum_CE_OI": 1234567,
      "Sum_PE_OI": 2345678,
      "Sum_CE_Change_OI": 12345,
      "Sum_PE_Change_OI": 23456,
      "Combined_OI": 0.4567,
      "Combined_CH_OI": 0.0123
    },
    ...
  ],
  "timestamp": "2025-12-09T15:30:00+05:30"
}
```

**Data is sorted by `Combined_CH_OI` in descending order.**

## Telegram Message Format

The top 10 stocks (by Combined_CH_OI) are sent to Telegram in this format:

```
📊 Stock OI Analysis - Top 10
Date: 2025-12-09
Total Stocks: 180

Top 10 by Combined CH_OI:

RELIANCE
  OI: +0.4567 | CH_OI: +0.0123
TCS
  OI: -0.2345 | CH_OI: +0.0098
...
```

## Metrics Explanation

### Combined_OI
```
Combined_OI = (Sum_PE_OI - Sum_CE_OI) / MIN(Sum_PE_OI, Sum_CE_OI)
```
- Positive value: More Put OI (bearish sentiment)
- Negative value: More Call OI (bullish sentiment)

### Combined_CH_OI
```
Combined_CH_OI = (Sum_PE_Change_OI - Sum_CE_Change_OI) / (Sum_PE_OI + Sum_CE_OI)
```
- Positive value: More Put OI being added (bearish build-up)
- Negative value: More Call OI being added (bullish build-up)

## Timezone Handling

The script uses **IST (Indian Standard Time)** for all date calculations to ensure correct behavior when deployed on Vercel (which runs in UTC). This fixes the issue where dates would be off by one day when the script runs before 5:30 AM IST.

## Troubleshooting

| Issue | Resolution |
|-------|------------|
| No symbols loaded | Check that `fno_stock_list.xlsx` exists and has a 'Symbol' column |
| NSE API blocked | NSE may rate-limit requests. Try reducing the number of stocks or adding delays |
| Telegram not sending | Verify `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID` are set correctly |
| Wrong date in output | Timezone fix applied - should use IST correctly now |
| Vercel timeout | Vercel functions have a 10-second timeout on free tier. Consider upgrading or reducing stock count |

## File Requirements

### fno_stock_list.xlsx

Must contain a column named `Symbol` with stock symbols:

| Symbol |
|--------|
| RELIANCE |
| TCS |
| INFY |
| ... |

## Notes

- The script automatically detects strike price intervals for each stock
- Fetches data for 7 strikes: ATM ± 3 strikes
- On Vercel, Excel file output is disabled (read-only filesystem)
- All data is returned via JSON API
- Failed stocks are logged and included in the response

---

Prepared for Vercel deployment with timezone fixes and Telegram integration.
