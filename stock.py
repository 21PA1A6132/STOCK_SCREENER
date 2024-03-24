import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import requests
import json
import pandas as pd
from openpyxl import Workbook

def fetch_stock_data(capital, volume, sector, dividend, exchange):
    api_key = '846cad55fe3b2544461bfde238aa3d43'
    url = f'https://financialmodelingprep.com/api/v3/stock-screener?marketCapMoreThan={capital}&betaMoreThan=1&volumeMoreThan={volume}&sector={sector}&exchange={exchange}&dividendMoreThan={dividend}&limit=100&apikey={api_key}'
    response = requests.get(url)
    data = json.loads(response.text)
    return data

def run_stock_screener():
    try:\
        # Get user input from Entry widgets
        capital = float(capital_entry.get())
        volume = float(volume_entry.get())
        sector = sector_entry.get()
        exchange = exchange_entry.get()
        dividend = float(dividend_entry.get())

        # Fetch stock data
        stock_data = fetch_stock_data(capital, volume, sector, dividend, exchange)

        # Convert data to DataFrame
        df = pd.DataFrame(stock_data)

        # Filter DataFrame based on minimum dividend yield
        df_filtered = df[df['lastAnnualDividend'] > dividend / 100 if dividend != 0 else 0.01]

        df_filtered.to_excel('stock_data.xlsx', index=False)

        # Display filtered results in the text widget
        result_text.config(state=tk.NORMAL)
        result_text.delete("1.0", tk.END)
        result_text.insert(tk.END, df_filtered.to_string(index=False))
        result_text.config(state=tk.DISABLED)
    except ValueError:
        messagebox.showerror("Error", "Invalid input. Please enter numeric values for capital, volume, and dividend.")\

# Create the main window
root = tk.Tk()
root.title("Stock Screener")

# Create and pack widgets
capital_label = tk.Label(root, text="Minimum Market Capital:")
capital_label.pack(pady=5)
capital_entry = tk.Entry(root)
capital_entry.pack(pady=5)

volume_label = tk.Label(root, text="Minimum Stock Volume:")
volume_label.pack(pady=5)
volume_entry = tk.Entry(root)
volume_entry.pack(pady=5)

sector_label = tk.Label(root, text="Stock Sector:")
sector_label.pack(pady=5)
sector_entry = tk.Entry(root)
sector_entry.pack(pady=5)

exchange_label = tk.Label(root, text="Stock Exchange:")
exchange_label.pack(pady=5)
exchange_entry = tk.Entry(root)
exchange_entry.pack(pady=5)

dividend_label = tk.Label(root, text="Minimum Dividend Yield:")
dividend_label.pack(pady=5)
dividend_entry = tk.Entry(root)
dividend_entry.pack(pady=5)

run_button = tk.Button(root, text="Run Stock Screener", command=run_stock_screener)
run_button.pack(pady=10)

result_text = tk.Text(root, height=10, width=50)
result_text.pack(pady=10)
result_text.config(state=tk.DISABLED)

# Run the

# Run the Tkinter event loop
root.mainloop()
