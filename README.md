# Crypto-API-Extraction
 Project Overview
This project demonstrates the automated extraction of live financial data using the CoinMarketCap API. I developed a Python pipeline that requests real-time price, volume, and market cap data for the top cryptocurrencies. This data is then normalized and stored in a structured format for longitudinal price-trend analysis.

 Tools & Technologies
Python: Used as the primary engine for API communication.

Requests Library: Handles the HTTP GET requests and API key authentication.

JSON: The data format used for parsing the live server response.

Pandas: Used to transform "nested" JSON data into a clean, flat data frame.

 Project Structure
API Connection Engine: A script that manages authentication and pulls the latest "Latest Listings" from the server.

Data Normalization: A process that flattens complex JSON structures into a clean table.

Automated Storage: The script appends new data to a local CSV/Excel file, allowing for the creation of a custom historical database.

 Key Skills Demonstrated
API Authentication: Safely handling API keys and header requirements.

Data Normalization: Converting multi-layered JSON objects into a single-layer table ready for analysis.

Automation: Scheduling regular "hits" to the API to track market volatility over time.

 How to Use
Get an API Key: Sign up at CoinMarketCap for a free developer key.

Configure the Script: Replace the X-CMC_PRO_API_KEY placeholder with your actual key.

Run the Pipeline: Execute the script to begin collecting live market data.
