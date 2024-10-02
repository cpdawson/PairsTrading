import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

# Read the updated Excel file into a DataFrame
df = pd.read_excel('PairsTradingAnalysis.xlsx')

# Get the list of unique tickers
tickers = df['Ticker'].unique()

# Initialize dictionary to hold stock data
stock_data = {}

# Define end date as today and start date as one year ago
end_date = datetime.today()
start_date = end_date - timedelta(days=365)

# Download the stock price data for each ticker
for ticker in tickers:
    stock_data[ticker] = yf.download(ticker, start=start_date, end=end_date, progress=False)
    print(ticker)

# Compute correlation matrix
price_data = pd.DataFrame({ticker: data['Adj Close'] for ticker, data in stock_data.items()})
correlation_matrix = price_data.corr()

# Find pairs with correlation greater than 0.92
correlated_pairs = []
for i, ticker1 in enumerate(tickers):
    for j, ticker2 in enumerate(tickers):
        if i < j and correlation_matrix.at[ticker1, ticker2] > 0.92:
            print(f"Pair: {ticker1} - {ticker2}, Correlation: {correlation_matrix.at[ticker1, ticker2]}")
            correlated_pairs.append((ticker1, ticker2, correlation_matrix.at[ticker1, ticker2]))

# Convert correlated pairs to DataFrame
correlated_pairs_df = pd.DataFrame(correlated_pairs, columns=['Ticker 1', 'Ticker 2', 'Correlation'])

# Save to a new Excel file
correlated_pairs_df.to_excel('Correlated_Pairs.xlsx', index=False)

