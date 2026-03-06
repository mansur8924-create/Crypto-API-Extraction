"""
Crypto Market Tracker

Tracks the top 10 cryptocurrencies from CoinGecko, calculates RSI
to measure momentum, gives simple Buy/Hold/Sell recommendations,
and saves a styled Excel report with a market cap chart.

Author: Mansur Mohammed
"""

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import time
from datetime import datetime
import warnings

# Suppress pandas warnings for clean console output
warnings.simplefilter("ignore", category=FutureWarning)


class CryptoMarketPipeline:
    """
    A pipeline to fetch, process, analyze, and report crypto market data.
    """

    def __init__(self, report_file):
        # CoinGecko endpoints
        self.market_url = "https://api.coingecko.com/api/v3/coins/markets"
        self.global_url = "https://api.coingecko.com/api/v3/global"

        # Output paths
        self.report_file = report_file
        self.plot_file = "Market_Cap_Visual.png"

        # Parameters: top 10 USD coins with 7-day sparkline data
        self.params = {
            "vs_currency": "usd",
            "order": "market_cap_desc",
            "per_page": 10,
            "page": 1,
            "sparkline": "true",
        }

        # Create a session with retries for reliability
        self.session = self._create_session()

    def _create_session(self):
        """
        Set up a robust HTTP session with retry logic for network reliability.
        """
        session = requests.Session()
        retries = Retry(
            total=3,
            backoff_factor=2,
            status_forcelist=[429, 500, 502, 503, 504]
        )
        session.mount("https://", HTTPAdapter(max_retries=retries))
        return session

    def _calculate_rsi(self, prices, period=14):
        """
        Compute the Relative Strength Index (RSI) from price data.
        Returns 50 if there is insufficient data.
        """
        try:
            s = pd.Series(prices)
            delta = s.diff()
            gain = delta.clip(lower=0).rolling(period).mean()
            loss = -delta.clip(upper=0).rolling(period).mean()
            rs = gain / loss
            return 100 - (100 / (1 + rs)).iloc[-1]
        except Exception:
            return 50.0  # Neutral RSI if data is missing

    def _get_recommendation(self, rsi, change):
        """
        Generate a simple Buy/Hold/Sell recommendation based on RSI and 24h price change.
        """
        if rsi < 35:
            return "BUY"
        elif rsi > 65 and change > 5:
            return "SELL"
        else:
            return "HOLD"

    def _fetch_data(self):
        """
        Fetch market and global crypto data from CoinGecko.
        """
        try:
            global_data = self.session.get(self.global_url, timeout=15).json()
            total_mcap = global_data["data"]["total_market_cap"]["usd"]
            market_data = self.session.get(self.market_url, params=self.params, timeout=15).json()
            return market_data, total_mcap
        except requests.exceptions.RequestException as e:
            print(f"⚠️ Network error: {e}")
            return None, None

    def _process_data(self, raw_data, total_mcap):
        """
        Transform raw API data into a DataFrame with key metrics and recommendations.
        """
        df = pd.DataFrame(raw_data)
        df["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        df["Dominance %"] = (df["market_cap"] / total_mcap) * 100
        df["RSI"] = df["sparkline_in_7d"].apply(lambda x: self._calculate_rsi(x["price"]))
        df["Recommendation"] = df.apply(
            lambda row: self._get_recommendation(row["RSI"], row["price_change_percentage_24h"]), axis=1
        )
        df["Trend"] = df["price_change_percentage_24h"].apply(lambda x: "▲" if x > 0 else "▼" if x < 0 else "-")

        # Rename columns for a professional report
        df = df.rename(columns={
            "market_cap_rank": "Rank",
            "name": "Asset",
            "symbol": "Ticker",
            "current_price": "Price",
            "price_change_percentage_24h": "24h Change %",
            "total_volume": "24h Volume",
            "market_cap": "Market Cap"
        })

        cols = ["Timestamp", "Trend", "Recommendation", "Rank", "Asset", "Ticker", "Price",
                "24h Change %", "RSI", "Dominance %", "24h Volume", "Market Cap"]
        return df[cols]

    def _save_excel(self, df):
        """
        Save the processed DataFrame to a styled Excel report.
        """
        writer = pd.ExcelWriter(self.report_file, engine="xlsxwriter")
        df.to_excel(writer, index=False, sheet_name="Top 10 Crypto", startrow=4)
        workbook = writer.book
        worksheet = writer.sheets["Top 10 Crypto"]

        # Simple header styling
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E1F2"})
        worksheet.set_row(4, None, header_fmt)
        worksheet.set_column("A:L", 15)

        writer.close()

    def _generate_plot(self, df):
        """
        Create a bar chart of top 10 market caps.
        """
        plt.figure(figsize=(10, 5))
        sns.barplot(x="Asset", y="Market Cap", data=df, palette="viridis")
        plt.title(f"Top 10 Crypto Market Caps ({datetime.now().strftime('%H:%M')})")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(self.plot_file)
        plt.close()

    def run_cycle(self):
        """
        Run one end-to-end update: fetch, process, save Excel, and generate plot.
        """
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Updating crypto market data...")
        data, total_mcap = self._fetch_data()
        if data:
            df = self._process_data(data, total_mcap)
            try:
                self._save_excel(df)
                self._generate_plot(df)
                print("✅ Report updated successfully!\n")
                print(df[["Ticker", "Price", "24h Change %", "Recommendation"]].to_string(index=False))
            except Exception as e:
                print(f"⚠️ Error saving report: {e}")


if __name__ == "__main__":
    REPORT_FILE = "Crypto_Top10_Report.xlsx"
    pipeline = CryptoMarketPipeline(REPORT_FILE)

    print("🚀 Crypto Market Tracker started.")

    try:
        while True:
            pipeline.run_cycle()
            print("⏳ Waiting 30 minutes until the next update...\n")
            time.sleep(1800)  # 30-minute refresh
    except KeyboardInterrupt:
        print("🛑 Pipeline stopped by user.")
