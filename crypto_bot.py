"""
PROJECT: Crypto Executive Market Intelligence Pipeline
PURPOSE: Automated extraction, technical analysis, and reporting of Top 10 Crypto Assets.
KEY FEATURES:
    - Robust Session Management (Exponential Backoff Retries).
    - Technical Analysis: Relative Strength Index (RSI) calculation from Sparkline data.
    - Financial Logic: Automated 'Buy/Hold/Sell' recommendations based on momentum.
    - Executive Reporting: Styled Excel output with conditional formatting and automated Matplotlib visuals.
AUTHOR: Mansur Mohammed
DATE: 2026-03-04
"""

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
import os
import time
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings

# Suppress system-level warnings to ensure a professional, clean console interface.
warnings.simplefilter(action='ignore', category=FutureWarning)

class CryptoExecutivePipeline:
    def __init__(self, target_filename):
        """
        Initializes the intelligence engine, defining API endpoints and dynamic file paths.
        """
        # API Endpoints: CoinGecko Public V3
        self.market_url = "https://api.coingecko.com/api/v3/coins/markets"
        self.global_url = "https://api.coingecko.com/api/v3/global"
        
        # PORTABLE PATH LOGIC: Ensures the script runs regardless of the user's directory structure.
        self.target_file = target_filename
        self.plot_file = 'Market_Cap_Visual.png'
        
        # Extraction Parameters: Top 10 USD Assets, sorted by Market Cap.
        self.params = {
            'vs_currency': 'usd',
            'order': 'market_cap_desc',
            'per_page': 10,
            'page': 1,
            'sparkline': 'true' 
        }
        
        # Establishing a network connection with 'Shock Absorbers' (Retries).
        self.session = self._create_robust_session()

    def _create_robust_session(self):
        """
        Implements an HTTP session with an exponential backoff strategy for network reliability.
        """
        session = requests.Session()
        # Strategy: Retry 3 times on common server errors (500s) or Rate Limits (429).
        retries = Retry(total=3, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])
        session.mount('https://', HTTPAdapter(max_retries=retries))
        return session

    def _calculate_rsi(self, sparkline_data, period=14):
        """
        Calculates the Relative Strength Index (RSI) to measure price momentum.
        """
        try:
            series = pd.Series(sparkline_data)
            delta = series.diff()
            gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
            loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
            rs = gain / loss
            # Standard RSI Formula: 100 - (100 / (1 + RS))
            return 100 - (100 / (1 + rs)).iloc[-1]
        except Exception:
            return 50.0  # Return Neutral RSI in case of insufficient data.

    def _get_recommendation(self, rsi, change):
        """
        Business Logic: Quantitative recommendation engine based on RSI and price action.
        """
        if rsi < 35: return "STRONG BUY"      # Indicator of 'Oversold' conditions.
        elif rsi > 65 and change > 5: return "TAKE PROFIT" # Indicator of 'Overbought' conditions.
        else: return "HOLD"                  # Stable market conditions.

    def _fetch_intel(self):
        """
        Executes the data acquisition from the REST API.
        """
        try:
            # Acquiring Global Market Metrics for Dominance calculations.
            g_response = self.session.get(self.global_url, timeout=15).json()
            total_mcap = g_response['data']['total_market_cap']['usd']
            
            # Acquiring Asset-specific Market Data.
            c_data = self.session.get(self.market_url, params=self.params, timeout=15).json()
            return c_data, total_mcap
        except requests.exceptions.RequestException as e:
            logging.error(f"Network Latency/Error: {e}. Re-attempting in next cycle.")
            return None, None

    def _process(self, raw_data, total_mcap):
        """
        Data Transformation & Feature Engineering.
        """
        df = pd.DataFrame(raw_data)
        df['Timestamp'] = datetime.now().strftime("%m/%d/%Y %H:%M")
        
        # Feature Engineering: Market Dominance % and Momentum (RSI).
        df['Dominance %'] = (df['market_cap'] / total_mcap) * 100
        df['RSI (14)'] = df['sparkline_in_7d'].apply(lambda x: self._calculate_rsi(x['price']))
        
        # Applying Tactical Recommendations.
        df['Recommendation'] = df.apply(lambda row: self._get_recommendation(row['RSI (14)'], row['price_change_percentage_24h']), axis=1)
        df['Trend'] = df['price_change_percentage_24h'].apply(lambda x: '▲' if x > 0 else '▼' if x < 0 else '▬')
        
        # Column Normalization for professional reporting.
        df = df.rename(columns={
            'market_cap_rank': 'Rank',
            'name': 'Asset',
            'symbol': 'Ticker',
            'current_price': 'Live Price',
            'price_change_percentage_24h': '24h Change %',
            'total_volume': '24h Volume',
            'market_cap': 'Market Cap'
        })
        
        cols = ['Timestamp', 'Trend', 'Recommendation', 'Rank', 'Asset', 'Ticker', 'Live Price', 
                '24h Change %', 'RSI (14)', 'Dominance %', '24h Volume', 'Market Cap']
        return df[cols]

    def _generate_visuals(self, df):
        """
        Generates professional visualizations for executive review.
        """
        plt.figure(figsize=(12, 6))
        sns.set_style("whitegrid")
        sns.barplot(x='Asset', y='Market Cap', data=df, palette='viridis')
        
        plt.title(f'Market Capitalization Comparison ({datetime.now().strftime("%H:%M")})', fontsize=14, fontweight='bold')
        plt.xticks(rotation=45)
        plt.ylabel('Market Cap (USD)', fontsize=12)
        plt.tight_layout()
        plt.savefig(self.plot_file)
        plt.close()

    def _save_styled_excel(self, df):
        """
        Persistence Layer: Generates a 'Bank-Standard' Excel report with conditional styling.
        """
        writer = pd.ExcelWriter(self.target_file, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Market Intelligence', startrow=4)
        
        workbook  = writer.book
        worksheet = writer.sheets['Market Intelligence']

        # Formatting Profiles
        title_fmt  = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})
        label_fmt  = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        buy_fmt    = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})
        sell_fmt   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
        mcap_fmt   = workbook.add_format({'num_format': '$#,##0'})
        money_fmt  = workbook.add_format({'num_format': '$#,##0.00'})
        pct_fmt    = workbook.add_format({'num_format': '0.00"%"'})

        green_shd = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red_shd   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

        # Executive Summary Injection
        top_coin = df.loc[df['24h Change %'].idxmax()]
        vol_leader = df.loc[df['24h Volume'].idxmax()]
        
        worksheet.write('A1', 'EXECUTIVE MARKET REPORT: TOP 10 CRYPTO ASSETS', title_fmt)
        worksheet.write('A2', 'Top Performer (24h):', label_fmt)
        worksheet.write('B2', f"{top_coin['Asset']} (+{top_coin['24h Change %']:.2f}%)")
        worksheet.write('A3', 'Liquidity Leader:', label_fmt)
        worksheet.write('B3', f"{vol_leader['Asset']} (${vol_leader['24h Volume']:,.0f})")

        # Column Formatting
        worksheet.set_column('A:A', 18)
        worksheet.set_column('B:C', 18)
        worksheet.set_column('G:G', 15, money_fmt)
        worksheet.set_column('H:J', 12, pct_fmt)
        worksheet.set_column('K:L', 22, mcap_fmt)

        # Automated Conditional Formatting
        worksheet.conditional_format('H6:H15', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_shd})
        worksheet.conditional_format('H6:H15', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_shd})
        worksheet.conditional_format('C6:C15', {'type': 'cell', 'criteria': '==', 'value': '"STRONG BUY"', 'format': buy_fmt})
        worksheet.conditional_format('C6:C15', {'type': 'cell', 'criteria': '==', 'value': '"TAKE PROFIT"', 'format': sell_fmt})

        writer.close()

    def run_cycle(self):
        """
        Executes a single end-to-end data intelligence cycle.
        """
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 Executing Market Intelligence Sequence...")
        data, total_mcap = self._fetch_intel()
        
        if data:
            df = self._process(data, total_mcap)
            try:
                self._save_styled_excel(df)
                self._generate_visuals(df)
                
                print(f"✅ SUCCESS: Reports updated. Persistence confirmed.")
                print("-" * 60)
                print(df[['Ticker', 'Live Price', '24h Change %', 'Recommendation']].to_string(index=False))
                print("-" * 60)
                
            except PermissionError:
                print("❌ IO ERROR: File in use. Please close 'Professional_Crypto_Dashboard.xlsx'.")
            except Exception as e:
                print(f"❌ SYSTEM ERROR: {e}")

if __name__ == "__main__":
    # PORTABLE FILENAME: Saves directly into the project folder for GitHub compatibility.
    REPORT_NAME = 'Professional_Crypto_Dashboard.xlsx'
    
    bot = CryptoExecutivePipeline(REPORT_NAME)
    print("🤖 Crypto Executive Pipeline initialized. Continuous monitoring active.")

    try:
        while True:
            bot.run_cycle()
            # Professional cycle interval (30 Minutes)
            time.sleep(1800) 
    except KeyboardInterrupt:
        print("\n🛑 Manual Override: Pipeline safely deactivated.")
