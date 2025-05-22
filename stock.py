import time
import datetime
import itertools

import numpy as np
import pandas as pd
import yfinance as yf
import pandas_ta as ta
from tqdm import tqdm

from river import preprocessing, ensemble, metrics

# ─── PARAMETERS ───────────────────────────────────────────────────────────────
SYMBOL        = "AAPL"
HIST_MINUTES  = 50
WINDOW_MOM    = 5
BUDGET        = 10.0

# ─── ONLINE MODEL & METRIC ────────────────────────────────────────────────────
model  = (
    preprocessing.StandardScaler() |
    ensemble.AdaptiveRandomForestClassifier(n_models=20, seed=42)
)
metric = metrics.Rolling(metrics.Accuracy(), window_size=200)

# ─── STORAGE ──────────────────────────────────────────────────────────────────
print("DEBUG: Downloading warm-up history…")
df_hist  = yf.download(SYMBOL, period=f"{HIST_MINUTES}m", interval="1m", progress=False).dropna()
# compute static indicators once
df_hist.ta.rsi(length=14, append=True)
df_hist.ta.macd(append=True)
df_hist.ta.bbands(length=20, std=2, append=True)
df_hist.ta.atr(length=14, append=True)
df_hist.ta.obv(append=True)
df_hist["vwap"]  = (df_hist.Volume * (df_hist.High+df_hist.Low+df_hist.Close)/3).cumsum() / df_hist.Volume.cumsum()
df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()
df_hist.dropna(inplace=True)
print(f"DEBUG: Warm-up complete, {len(df_hist)} bars loaded.")

pending = None
cash, pos = BUDGET, 0.0

# ─── FEATURE EXTRACTION ───────────────────────────────────────────────────────
def make_features(df: pd.DataFrame) -> dict:
    last = df.iloc[-1]
    prev = df.iloc[-2]
    return {
        "pct_chg": (last.Close - prev.Close) / prev.Close,
        "mom5":    (last.Close - df.Close.shift(WINDOW_MOM).iloc[-1]) / df.Close.shift(WINDOW_MOM).iloc[-1],
        "rsi14":   last.RSI_14,
        "macd":    last.MACD_12_26_9,
        "macd_sig":last.MACDs_12_26_9,
        "bb_up":   last.BBUpper_20_2.0,
        "bb_low":  last.BBLower_20_2.0,
        "bb_bw":   (last.BBUpper_20_2.0 - last.BBLower_20_2.0) / last.BBMiddle_20_2.0,
        "bb_pctb": (last.Close - last.BBLower_20_2.0) / (last.BBUpper_20_2.0 - last.BBLower_20_2.0),
        "atr14":   last.ATR_14,
        "obv":     last.OBV,
        "vol_chg": (last.Volume - prev.Volume) / prev.Volume,
        "vwap":    last.vwap,
        "volat":   last.volat,
    }

# ─── MAIN LOOP (tqdm-wrapped) ─────────────────────────────────────────────────
for iteration in tqdm(itertools.count(1), desc="Signal iterations"):
    # ⏱ wait until next exact minute
    now   = datetime.datetime.now()
    to_sleep = 60 - now.second - now.microsecond/1e6
    print(f"DEBUG: Sleeping for {to_sleep:.2f}s to align to minute boundary")
    time.sleep(to_sleep)

    # fetch the last 2 bars
    df2 = yf.download(SYMBOL, period="2m", interval="1m", progress=False).dropna()
    if len(df2) < 2:
        print("DEBUG: Insufficient data, retrying…")
        continue

    prev_bar, curr_bar = df2.iloc[-2], df2.iloc[-1]
    ts = curr_bar.name.to_pydatetime()
    print(f"DEBUG: [{ts:%H:%M}] Fetched bars — prev_close={prev_bar.Close:.2f}, curr_close={curr_bar.Close:.2f}")

    # 1) Learn from last signal
    if pending:
        true_lbl = int(curr_bar.Close > pending["price"])
        print(f"DEBUG: Learning from last prediction: was_buy={pending['pred']==1}, actual_up={true_lbl==1}")
        model.learn_one(pending["feats"], true_lbl)
        metric.update(true_lbl, pending["pred"])
        print(f"DEBUG: Model updated — rolling accuracy={metric.get():.3f}")
        pending = None

    # 2) Update history & recompute indicators
    df_hist = pd.concat([df_hist, curr_bar.to_frame().T]).iloc[-HIST_MINUTES:]
    print("DEBUG: Recomputing indicators…")
    df_hist.ta.rsi(length=14, append=True)
    df_hist.ta.macd(append=True)
    df_hist.ta.bbands(length=20, std=2, append=True)
    df_hist.ta.atr(length=14, append=True)
    df_hist.ta.obv(append=True)
    df_hist["vwap"]  = (df_hist.Volume * (df_hist.High+df_hist.Low+df_hist.Close)/3).cumsum() / df_hist.Volume.cumsum()
    df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()
    df_hist.dropna(inplace=True)

    # 3) Feature extraction
    feats = make_features(df_hist)
    print(f"DEBUG: Features at {ts:%H:%M} → {feats}")

    # 4) Predict
    y_pred = model.predict_one(feats) or 0
    p_up   = model.predict_proba_one(feats).get(1, 0.0)
    action = "BUY" if y_pred else "SELL"
    print(f"DEBUG: Prediction → action={action}, P(up)={p_up:.2f}")

    # 5) Execute (simulated) with sizing
    price = curr_bar.Close
    if action == "BUY" and cash > 0:
        qty = cash / price
        pos += qty
        cash = 0.0
        print(f"[{ts:%H:%M}] BUY  {qty:.6f} shares @${price:.2f}")
    elif action == "SELL" and pos > 0:
        cash += pos * price
        print(f"[{ts:%H:%M}] SELL {pos:.6f} shares @${price:.2f}")
        pos = 0.0
    else:
        print(f"[{ts:%H:%M}] HOLD — cash=${cash:.2f}, pos={pos:.6f}")

    # 6) Explain
    try:
        imps = model.feature_importances_
        top2 = sorted(imps.items(), key=lambda x: abs(x[1]), reverse=True)[:2]
        reason = "; ".join(f"{n}→{w:.3f}" for n, w in top2)
    except Exception:
        reason = "—"

    print(f"         P(↑)={p_up:.2f} | {reason}")

    # 7) Stash for feedback
    pending = {"feats": feats, "price": price, "pred": y_pred}

    print(f"DEBUG: End of iteration {iteration} — cash=${cash:.2f}, pos={pos:.6f}\n")