import sys
import time
import math
import requests
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf

# ============== 基本設定 ==============
file_path = r"C:\Users\Leon\Desktop\程式語言資料\stock\投資組合損益.xlsx"
html_out  = "portfolio.html"
csv_out   = "portfolio.csv"
png_out   = "allocation.png"

# 偵測是否在 Jupyter 環境（決定輸出方式）
IS_JUPYTER = "ipykernel" in sys.modules

# === 股票 / ETF 代碼表 ===
stock_symbols = {
    "TSLA": "TSLA", "NVDA": "NVDA", "AMD": "AMD", "TSM": "TSM", "AAPL": "AAPL",
    "NIO": "NIO", "VGT": "VGT", "FIG": "FIG", "VOO": "VOO", "VTI": "VTI",
    "IBM": "IBM", "CENN": "CENN", "QQQ": "QQQ"
}

# === 加密幣代碼（Coingecko 的 ids） ===
crypto_symbols = {
    "ETH": "ethereum", "ADA": "cardano", "FIL": "filecoin", "SOL": "solana",
    "DOGE": "dogecoin", "TAO": "bittensor", "ATH": "ath", "COMP": "compound",
    "IOTA": "iota", "VET": "vechain", "CELR": "celer-network", "XTZ": "tezos",
    "ZEC": "zcash", "LUNC": "terra-luna", "LOOKS": "looksrare", "TRUMP": "trumpcoin",
    "BNB": "binancecoin",
}

# 目前持有現金（USD）
cash_usd = 132.43


# ============== 抓價函式（穩健版） ==============
def get_stock_price(symbol: str) -> float | None:
    """
    從 yfinance 抓近 5 天收盤價，取最後一筆有效值（避免非交易日/空表）。
    """
    try:
        # 使用 download 比 Ticker.history 更穩，且可 dropna
        hist = yf.download(symbol, period="5d", progress=False)
        if hist is None or hist.empty:
            return None
        close = hist["Close"].dropna()
        if close.empty:
            return None
        return float(close.iloc[-1])
    except Exception:
        return None


def get_crypto_prices_batch(ids: list[str]) -> dict[str, float | None]:
    """
    一次向 Coingecko 批次要價；回傳 {id: price or None}
    """
    if not ids:
        return {}
    url = "https://api.coingecko.com/api/v3/simple/price"
    # Coingecko 批次 ids 以逗號分隔
    params = {"ids": ",".join(ids), "vs_currencies": "usd"}
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json() or {}
        out: dict[str, float | None] = {}
        for cid in ids:
            out[cid] = data.get(cid, {}).get("usd")
        return out
    except Exception:
        # 失敗時全部回 None（也不會中斷主流程）
        return {cid: None for cid in ids}


# ============== 載入與轉型 ==============
# 讀 Excel
df = pd.read_excel(file_path)

# 確保數值欄位型態正確（把字串/空白轉成 NaN 後再為 float）
for col in ["投入(USD)", "現價(USD)", "持有數量", "現值(USD)", "損益(USD)", "損益率"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# ============== 更新報價 ==============
# 依照 df["資產"] 判斷該抓股票或加密幣價格
assets = list(df["資產"].astype(str))

# 股票：逐一抓價（數量通常不多，穩就好）
stock_price_map: dict[str, float | None] = {}
# 加密幣：先收集出現過的 ids 批次抓價
crypto_ids_needed = set()

for asset in assets:
    if asset in stock_symbols:
        stock_price_map[asset] = get_stock_price(stock_symbols[asset])
    elif asset in crypto_symbols:
        crypto_ids_needed.add(crypto_symbols[asset])

crypto_price_by_id = get_crypto_prices_batch(sorted(list(crypto_ids_needed)))

# 組合成 asset -> price 的新字典
new_prices: dict[str, float | None] = {}
for asset in assets:
    if asset in stock_symbols:
        new_prices[asset] = stock_price_map.get(asset)
    elif asset in crypto_symbols:
        cid = crypto_symbols[asset]
        new_prices[asset] = crypto_price_by_id.get(cid)
    else:
        new_prices[asset] = None

# 寫回現價（若抓不到就保留舊值）
if "現價(USD)" not in df.columns:
    df["現價(USD)"] = pd.NA
df["現價(USD)"] = df["資產"].map(new_prices).fillna(df["現價(USD)"])

# 重新計算現值 / 損益 / 損益率
df["現值(USD)"] = df["現價(USD)"] * df["持有數量"]
df["損益(USD)"] = df["現值(USD)"] - df["投入(USD)"]
df["損益率"]   = df["損益(USD)"] / df["投入(USD)"]

# ============== 總結（加入 CASH 前） ==============
# 注意：這裡先排除已存在的 CASH
df = df[df["資產"] != "CASH"].copy()

total_invested = float(pd.to_numeric(df["投入(USD)"], errors="coerce").fillna(0).sum())  # 總投入（不含現金）
total_pl_usd   = float(pd.to_numeric(df["損益(USD)"], errors="coerce").fillna(0).sum())  # 總損益
total_pl_pct   = (total_pl_usd / total_invested) if total_invested > 0 else 0.0

# 組合總現值（不含現金）
total_value_positions = float(pd.to_numeric(df["現值(USD)"], errors="coerce").fillna(0).sum())

# ============== 加入現金列 ==============
cash_row = {
    "資產": "CASH",
    "投入(USD)": cash_usd,
    "現價(USD)": math.nan,
    "持有數量": math.nan,
    "現值(USD)": cash_usd,
    "損益(USD)": 0.0,
    "損益率": 0.0,
}
df = pd.concat([df, pd.DataFrame([cash_row])], ignore_index=True)

# 倉位比例（含現金）
total_value_with_cash = total_value_positions + cash_usd
df["倉位比例"] = pd.to_numeric(df["現值(USD)"], errors="coerce") / (total_value_with_cash if total_value_with_cash > 0 else 1)

# 依倉位比例排序，並把索引從 1 開始
df = df.sort_values(by="倉位比例", ascending=False).reset_index(drop=True)
df.index = df.index + 1
df.index.name = "#"

# ============== 一致性快篩（抓怪數據，如 VGT 類型） ==============
issues = []
for i, row in df.iterrows():
    asset    = str(row["資產"])
    invested = float(row.get("投入(USD)", 0) or 0)
    mktval   = float(row.get("現值(USD)", 0) or 0)
    qty      = float(row.get("持有數量", 0) or 0)
    price    = float(row.get("現價(USD)", 0) or 0)

    # 1) 若「現值/投入」過高（>3），且不是現金，提醒檢查投入是否少記/加倉未更新
    if asset != "CASH" and invested > 0 and mktval > 0 and (mktval / invested) > 3:
        issues.append(f"{asset}: 現值/投入 = {mktval/invested:.2f}，可能是投入金額未涵蓋加倉/拆股，請檢查。")

    # 2) 若有價有量但「價格×數量」與現值差異過大，也提醒
    calc_val = price * qty
    if asset != "CASH" and price > 0 and qty > 0:
        if abs(calc_val - mktval) > max(1.0, 0.01 * mktval):
            issues.append(f"{asset}: 現值({mktval:.2f}) 與 價格×數量({calc_val:.2f}) 差異較大，請檢查數量/價格。")

# ============== 輸出（Jupyter 彩色；腳本純文字 + 存 HTML） ==============
def highlight_profit(val):
    try:
        return "color: green" if float(val) >= 0 else "color: red"
    except Exception:
        return ""

# 風格化（供 Jupyter 與 HTML 檔使用）
styled_df = (
    df.style
    .format({
        "投入(USD)": "{:,.2f}",
        "現價(USD)": "{:,.2f}",
        "持有數量": "{:,.4f}",
        "現值(USD)": "{:,.2f}",
        "損益(USD)": "{:,.2f}",
        "損益率": "{:.2%}",
        "倉位比例": "{:.2%}",
    }, na_rep="—")
    .applymap(highlight_profit, subset=["損益(USD)", "損益率"])
)

if IS_JUPYTER:
    # 在 Jupyter 顯示漂亮表格
    display(styled_df)
else:
    # 在終端機輸出純文字（避免看到 <style> HTML）
    print(df.to_string(index=True, justify="center", max_colwidth=20))

# 存成 HTML 方便用瀏覽器看
with open(html_out, "w", encoding="utf-8") as f:
    f.write(styled_df.to_html())

# ============== 總結輸出 ==============
print("\n📊 Portfolio Totals (exclude cash)")
print(f"   Total Invested: {total_invested:,.2f} USD")
print(f"   Total P/L:      {total_pl_usd:,.2f} USD")
print(f"   Total P/L %:    {total_pl_pct:.2%}")

print(f"\n💰 Total Portfolio Value (incl. cash): {total_value_with_cash:,.2f} USD")
print(f"   Cash Position: {cash_usd:,.2f} USD ({(cash_usd / total_value_with_cash if total_value_with_cash>0 else 0):.2%})")

if issues:
    print("\n⚠️ 一致性檢查提醒：")
    for msg in issues:
        print(" - " + msg)

# ============== 倉位比例長條圖（並存檔） ==============
plt.figure(figsize=(10, 6))
bars = plt.bar(df["資產"], df["倉位比例"] * 100)

plt.title("Portfolio Allocation by Asset", fontsize=14)
plt.ylabel("Weight (%)", fontsize=12)
plt.xticks(rotation=45, ha="right")

for bar, pct in zip(bars, df["倉位比例"] * 100):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
             f"{pct:.1f}%", ha="center", va="bottom", fontsize=9)

plt.tight_layout()
plt.savefig(png_out, dpi=150)
if IS_JUPYTER:
    plt.show()
else:
    print(f"\n🖼 已輸出圖檔：{png_out}")

# ============== 存 CSV（方便匯整/追蹤） ==============
df.to_csv(csv_out, encoding="utf-8-sig")
print(f"✅ 已輸出：{csv_out}")
print(f"✅ 已輸出：{html_out}")
