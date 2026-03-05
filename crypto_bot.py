import requests # The tool that makes internet requests to get data from CoinGecko
from requests.adapters import HTTPAdapter # Helps us customize how the internet connection works
from urllib3.util.retry import Retry # Adds "persistence" so the robot retries if the internet glitches
import pandas as pd # The primary engine for organizing data into tables (DataFrames)
import os # The construction crew that handles folders and file paths on your computer
import time # Used to make the robot wait (hibernate) between cycles
import matplotlib.pyplot as plt # The core tool for drawing charts and graphs
import seaborn as sns # Enhances the look of charts with professional colors and styles
from datetime import datetime # Creates timestamps so we know exactly when data was pulled
import warnings # Used to hide minor, annoying system warnings

# Tells Python to ignore "FutureWarnings" to keep your terminal looking clean
warnings.simplefilter(action='ignore', category=FutureWarning)

class CryptoExecutivePipeline:
    def __init__(self, target_path):
        # The specific API addresses for market data and global market stats
        self.market_url = "https://api.coingecko.com/api/v3/coins/markets"
        self.global_url = "https://api.coingecko.com/api/v3/global"
        
        # Identifies the folder where the file will live
        self.base_path = os.path.dirname(target_path)
        # Ensures the file ends in .xlsx so Excel's professional styling works
        self.target_file = target_path.replace('.csv', '.xlsx')
        # Defines the name and location for the automated bar chart image
        self.plot_file = os.path.join(self.base_path, 'Market_Cap_Visual.png')
        
        # Settings for the API: Get Top 10 coins in USD, sorted by size, with price history
        self.params = {
            'vs_currency': 'usd',
            'order': 'market_cap_desc',
            'per_page': 10,
            'page': 1,
            'sparkline': 'true' 
        }
        # Creates a "Robust Session" (an internet connection with shock absorbers)
        self.session = self._create_robust_session()

    def _create_robust_session(self):
        """Creates an internet connection that automatically retries if CoinGecko is busy."""
        session = requests.Session()
        # Retry 3 times, wait longer between each try, and retry on common internet errors
        retries = Retry(total=3, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])
        # "Mounts" these rules to every 'https://' request we make
        session.mount('https://', HTTPAdapter(max_retries=retries))
        return session

    def _calculate_rsi(self, sparkline_data, period=14):
        """Mathematical formula to find the Relative Strength Index (Momentum)."""
        try:
            series = pd.Series(sparkline_data) # Turn price list into a math series
            delta = series.diff() # Find the price change between each point
            gain = (delta.where(delta > 0, 0)).rolling(window=period).mean() # Average of green moves
            loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean() # Average of red moves
            rs = gain / loss # Find the ratio
            return 100 - (100 / (1 + rs)).iloc[-1] # The final RSI score (0-100)
        except Exception:
            return 50.0 # Return a neutral score if the math fails

    def _get_recommendation(self, rsi, change):
        """Strategic logic that decides if a coin is a Buy or Sell based on RSI."""
        if rsi < 35: return "STRONG BUY" # RSI under 35 means the coin is "Oversold" (Cheap)
        elif rsi > 65 and change > 5: return "TAKE PROFIT" # Over 65 means "Overbought" (Expensive)
        else: return "HOLD" # Stay the course if it's in the middle

    def _fetch_intel(self):
        """Goes to the internet and retrieves the raw data packets."""
        try:
            # Get global market cap to calculate how much of the market each coin owns
            g_data = self.session.get(self.global_url, timeout=15).json()
            total_mcap = g_data['data']['total_market_cap']['usd']
            # Get the top 10 coins list
            c_data = self.session.get(self.market_url, params=self.params, timeout=15).json()
            return c_data, total_mcap
        except requests.exceptions.RequestException as e:
            # Log the error if the internet fails
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ⚠️ Network hiccup detected: {e}. Retrying next cycle.")
            return None, None

    def _process(self, raw_data, total_mcap):
        """The 'Chef' phase: Takes raw data and cooks it into a clean table."""
        df = pd.DataFrame(raw_data)
        # Adds the current date and time to every row
        df['Timestamp'] = datetime.now().strftime("%m/%d/%Y %H:%M")
        # Calculates Dominance: (Coin Size / Total Market Size) * 100
        df['Dominance %'] = (df['market_cap'] / total_mcap) * 100
        # Runs the RSI math on the 7-day price history (sparkline)
        df['RSI (14)'] = df['sparkline_in_7d'].apply(lambda x: self._calculate_rsi(x['price']))
        # Applies our Buy/Sell logic to every coin
        df['Recommendation'] = df.apply(lambda row: self._get_recommendation(row['RSI (14)'], row['price_change_percentage_24h']), axis=1)
        # Adds a visual Arrow based on price movement
        df['Trend'] = df['price_change_percentage_24h'].apply(lambda x: '▲' if x > 0 else '▼' if x < 0 else '▬')
        
        # Renames messy API names into professional labels for your report
        df = df.rename(columns={
            'market_cap_rank': 'Rank',
            'name': 'Asset',
            'symbol': 'Ticker',
            'current_price': 'Live Price',
            'price_change_percentage_24h': '24h Change %',
            'total_volume': '24h Volume',
            'market_cap': 'Market Cap'
        })
        
        # Organizes the columns in a specific order that makes sense for an analyst
        cols = ['Timestamp', 'Trend', 'Recommendation', 'Rank', 'Asset', 'Ticker', 'Live Price', 
                '24h Change %', 'RSI (14)', 'Dominance %', '24h Volume', 'Market Cap']
        return df[cols]

    def _generate_visuals(self, df):
        """Draws a professional bar chart comparing the size of the assets."""
        plt.figure(figsize=(12, 6)) # Sets the canvas size
        sns.set_style("whitegrid") # Adds a clean white background with gridlines
        # Draws the bars using the Asset names and their Market Caps
        sns.barplot(x='Asset', y='Market Cap', data=df, palette='viridis')
        # Adds titles and labels
        plt.title(f'Top 10 Market Capitalization ({datetime.now().strftime("%H:%M")})', fontsize=14, fontweight='bold')
        plt.xticks(rotation=45) # Tilts the names so they don't overlap
        plt.ylabel('Market Cap (USD)', fontsize=12)
        plt.tight_layout() # Ensures nothing gets cut off
        plt.savefig(self.plot_file) # Saves the chart as an image on your Desktop
        plt.close() # Clears the canvas for the next cycle

    def _save_styled_excel(self, df):
        """The 'Painter' phase: Saves data to Excel and adds professional colors."""
        writer = pd.ExcelWriter(self.target_file, engine='xlsxwriter')
        # Places the data in the sheet, starting at row 4 to leave room for the header
        df.to_excel(writer, index=False, sheet_name='Market Intelligence', startrow=4)
        
        workbook  = writer.book
        worksheet = writer.sheets['Market Intelligence']

        # Definitions for colors, fonts, and borders (Bank Standard)
        title_fmt  = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})
        label_fmt  = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        buy_fmt    = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})
        sell_fmt   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
        mcap_fmt   = workbook.add_format({'num_format': '$#,##0'}) # Commas and No Decimals
        money_fmt  = workbook.add_format({'num_format': '$#,##0.00'}) # Commas and Decimals
        pct_fmt    = workbook.add_format({'num_format': '0.00"%"'}) # Percent sign

        green_shd = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red_shd   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

        # Finds the specific coins with the best performance for the Summary section
        top_coin = df.loc[df['24h Change %'].idxmax()]
        vol_leader = df.loc[df['24h Volume'].idxmax()]
        
        # Writes the Executive Summary at the very top of the Excel sheet
        worksheet.write('A1', 'EXECUTIVE MARKET REPORT: TOP 10 ASSETS', title_fmt)
        worksheet.write('A2', 'Top Gainer (24h):', label_fmt)
        worksheet.write('B2', f"{top_coin['Asset']} ({top_coin['24h Change %']:.2f}%)")
        worksheet.write('A3', 'Highest Volume:', label_fmt)
        worksheet.write('B3', f"{vol_leader['Asset']} (${vol_leader['24h Volume']:,.0f})")

        # Sets the widths for columns so the text isn't squished
        worksheet.set_column('A:A', 18) # Timestamp
        worksheet.set_column('B:C', 18) # Trend & Recommendation
        worksheet.set_column('G:G', 15, money_fmt) # Price
        worksheet.set_column('H:J', 12, pct_fmt) # Percentages
        worksheet.set_column('K:L', 22, mcap_fmt) # Market Cap & Volume

        # Conditional Formatting: Automatically paints cells green if price is UP, and red if DOWN
        worksheet.conditional_format('H6:H15', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_shd})
        worksheet.conditional_format('H6:H15', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_shd})
        # Conditional Formatting for Recommendations: STRONG BUY is Green, TAKE PROFIT is Red
        worksheet.conditional_format('C6:C15', {'type': 'cell', 'criteria': '==', 'value': '"STRONG BUY"', 'format': buy_fmt})
        worksheet.conditional_format('C6:C15', {'type': 'cell', 'criteria': '==', 'value': '"TAKE PROFIT"', 'format': sell_fmt})

        writer.close() # Finalizes and saves the Excel file

    def run_cycle(self):
        """The master switch that runs one full cycle of the robot."""
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 Initiating Market Intelligence Scan...")
        data, total_mcap = self._fetch_intel()
        
        if data:
            df = self._process(data, total_mcap)
            try:
                self._save_styled_excel(df) # Save to Excel
                self._generate_visuals(df) # Draw the chart
                
                print(f"✅ SUCCESS: Data refreshed. Graph and Excel updated.")
                print("-" * 60)
                # Prints a mini-dashboard directly into your terminal for quick viewing
                display_cols = ['Ticker', 'Live Price', '24h Change %', 'Recommendation']
                print(df[display_cols].to_string(index=False))
                print("-" * 60)
                print("⏳ Sleeping for 30 minutes. Press Ctrl+C to stop the engine.")
                
            except PermissionError:
                # If you have the file open, the robot can't write to it!
                print("❌ ERROR: Permission Denied. Please CLOSE the Excel workbook so Python can update it.")
            except Exception as e:
                print(f"❌ ERROR: A system failure occurred: {e}")

if __name__ == "__main__":
    # Your official folder path on your Desktop
    DEST = r'C:\Users\mansu\OneDrive\Desktop\Data Analyst Boot Camp\API pulling(project)-python\Professional_Crypto_Dashboard.xlsx'
    
    # Creates the folder if it doesn't exist yet to prevent crashes
    os.makedirs(os.path.dirname(DEST), exist_ok=True)
    bot = CryptoExecutivePipeline(DEST)

    print("🤖 Crypto Executive Pipeline is ONLINE. Starting automated cycles...")
    try:
        while True: # Infinite loop
            bot.run_cycle() # Run the engine
            time.sleep(1800) # Hibernate for 30 minutes
    except KeyboardInterrupt:
        # The emergency stop button (Ctrl+C)
        print("\n🛑 Manual Override: Pipeline safely shut down.")
