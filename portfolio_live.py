import yfinance as yf
import requests
import pandas as pd
import matplotlib.pyplot as plt

# === Path ===
file_path = r"C:\Users\Leon\Desktop\程式語言資料\stock\投資組合損益.xlsx"

# === Tickers (Stocks & ETFs) ===
stock_symbols = {
    "TSLA": "TSLA",
    "NVDA": "NVDA",
    "AMD": "AMD",
    "TSM": "TSM",
    "AAPL": "AAPL",
    "NIO": "NIO",
    "VGT": "VGT",
    "FIG": "FIG",
    "VOO": "VOO",
    "VTI": "VTI",
    "IBM": "IBM",
    "CENN": "CENN",
}

# === Tickers (Crypto) ===
crypto_symbols = {
    "ETH": "ethereum",
    "ADA": "cardano",
    "FIL": "filecoin",
    "SOL": "solana",
    "DOGE": "dogecoin",
    "TAO": "bittensor",
    "ATH": "ath",
    "COMP": "compound",
    "IOTA": "iota",
    "VET": "vechain",
    "CELR": "celer-network",
    "XTZ": "tezos",
    "ZEC": "zcash",
    "LUNC": "terra-luna",
    "LOOKS": "looksrare",
    "TRUMP": "trumpcoin",
    "BNB": "binancecoin",
}

# === Get stock price ===
def get_stock_price(symbol):
    try:
        ticker = yf.Ticker(symbol)
        price = ticker.history(period="1d")["Close"].iloc[-1]
        return float(price)
    except Exception:
        return None

# === Get crypto price ===
def get_crypto_price(symbol):
    try:
        url = f"https://api.coingecko.com/api/v3/simple/price"
        params = {"ids": symbol, "vs_currencies": "usd"}
        resp = requests.get(url, params=params).json()
        return resp[symbol]["usd"]
    except Exception:
        return None

# === Load portfolio ===
df = pd.read_excel(file_path)

# Update prices
new_prices = {}
for asset in df["資產"]:
    if asset in stock_symbols:
        new_prices[asset] = get_stock_price(stock_symbols[asset])
    elif asset in crypto_symbols:
        new_prices[asset] = get_crypto_price(crypto_symbols[asset])
    else:
        new_prices[asset] = None

df["現價(USD)"] = df["資產"].map(new_prices).fillna(df["現價(USD)"])

# Recalculate
df["現值(USD)"] = df["現價(USD)"] * df["持有數量"]
df["損益(USD)"] = df["現值(USD)"] - df["投入(USD)"]
df["損益率"] = df["損益(USD)"] / df["投入(USD)"]



# === Totals BEFORE adding CASH (only positions, exclude cash) ===
total_invested = float(df["投入(USD)"].sum())             # 總投入金額（不含現金）
total_pl_usd   = float(df["損益(USD)"].sum())             # 目前總盈虧（USD）
total_pl_pct   = (total_pl_usd / total_invested) if total_invested != 0 else 0.0  # 總盈虧%

# Total value
total_value = df["現值(USD)"].sum()

# Cash position 目前持有現金
cash_usd = 97.87

# 如果已經存在 CASH，先刪掉
df = df[df["資產"] != "CASH"]

# 再新增新的 CASH 列
cash_row = {
    "資產": "CASH",
    "投入(USD)": cash_usd,
    "現價(USD)": None,
    "持有數量": None,
    "現值(USD)": cash_usd,
    "損益(USD)": 0,
    "損益率": 0,
}
df = pd.concat([df, pd.DataFrame([cash_row])], ignore_index=True)

# Portfolio weight
total_value_with_cash = total_value + cash_usd
df["倉位比例"] = df["現值(USD)"] / total_value_with_cash

# Sort by weight
df = df.sort_values(by="倉位比例", ascending=False).reset_index(drop=True)

# === 讓最左邊索引從 1 開始，並命名為 '#' ===
df.index = df.index + 1      # 將 0-based 變成 1-based
df.index.name = "#"          # 設定索引欄位名稱顯示在表頭

# === Display DataFrame (colored) ===
def highlight_profit(val):
    color = "red" if val < 0 else "green"
    return f"color: {color}"

styled_df = (
    df.style
    .format({
        "投入(USD)": "{:,.2f}",
        "現價(USD)": "{:,.2f}",
        "持有數量": "{:,.4f}",
        "現值(USD)": "{:,.2f}",
        "損益(USD)": "{:,.2f}",
        "損益率": "{:.2%}",
        "倉位比例": "{:.2%}"
    })
    .applymap(highlight_profit, subset=["損益(USD)", "損益率"])
)

display(styled_df)

# === Portfolio totals (positions only, exclude cash) ===
print("📊 Portfolio Totals (exclude cash)")
print(f"   Total Invested: {total_invested:,.2f} USD")
print(f"   Total P/L:      {total_pl_usd:,.2f} USD")
print(f"   Total P/L %:    {total_pl_pct:.2%}")

# === Portfolio summary ===
print(f"💰 Total Portfolio Value (incl. cash): {total_value_with_cash:,.2f} USD")
print(f"   Cash Position: {cash_usd:,.2f} USD ({cash_usd/total_value_with_cash:.2%})")

# === Bar chart: portfolio weights ===
plt.figure(figsize=(10,6))
bars = plt.bar(df["資產"], df["倉位比例"]*100)

# Labels
plt.title("Portfolio Allocation by Asset", fontsize=14)
plt.ylabel("Weight (%)", fontsize=12)
plt.xticks(rotation=45, ha="right")

# Annotate bars with %
for bar, pct in zip(bars, df["倉位比例"]*100):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
             f"{pct:.1f}%", ha="center", va="bottom", fontsize=9)

plt.tight_layout()
plt.show()
