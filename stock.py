# main.py

# ─── Monkey-patch NumPy so pandas_ta can import NaN ──────────────────────────────
import numpy as np
if not hasattr(np, "NaN"):
    np.NaN = np.nan

import time
import datetime
import itertools
from collections import deque

import pandas as pd
import yfinance as yf
import pandas_ta as ta
from tqdm import tqdm

from river import preprocessing
from river.ensemble import BaggingClassifier
from river.tree import HoeffdingTreeClassifier

# ─── PARAMETERS ────────────────────────────────────────────────────────────────
SYMBOL         = "AAPL"      # Ticker to trade
HIST_MINUTES   = 50          # How many minutes of history to keep for indicators
WINDOW_MOM     = 5           # Momentum look-back
BUDGET         = 10.0        # USD starting cash
METRIC_WINDOW  = 200         # Rolling window size for performance metrics

# ─── ONLINE MODEL ──────────────────────────────────────────────────────────────
model = (
    preprocessing.StandardScaler()
    | BaggingClassifier(
        model=HoeffdingTreeClassifier(),
        n_models=15,
        seed=42
      )
)

# ─── PERFORMANCE BUFFER ────────────────────────────────────────────────────────
# We'll store (true_label, prediction) pairs here:
metric_buffer = deque(maxlen=METRIC_WINDOW)
prev_acc = None  # for trend-confidence calculation

# ─── WARM-UP: download initial history & compute indicators ────────────────────
print("DEBUG: Downloading warm-up history…")
df_hist = (
    yf.download(SYMBOL,
                period=f"{HIST_MINUTES}m",
                interval="1m",
                progress=False)
      .dropna()
)

# Compute static TA indicators
df_hist.ta.rsi(length=14, append=True)
df_hist.ta.macd(append=True)
df_hist.ta.bbands(length=20, std=2, append=True)
df_hist.ta.atr(length=14, append=True)
df_hist.ta.obv(append=True)

# VWAP & rolling volatility
df_hist["vwap"]  = (df_hist.Volume
                   * (df_hist.High + df_hist.Low + df_hist.Close) / 3).cumsum() \
                  / df_hist.Volume.cumsum()
df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()

df_hist.dropna(inplace=True)
print(f"DEBUG: Warm-up complete, {len(df_hist)} bars loaded.")

pending = None
cash, pos = BUDGET, 0.0  # start fully in cash

# ─── FEATURE ENGINEERING ───────────────────────────────────────────────────────
def make_features(df: pd.DataFrame) -> dict:
    last = df.iloc[-1]
    prev = df.iloc[-2]
    return {
        "pct_chg":  (last.Close - prev.Close) / prev.Close,
        "mom5":     (last.Close - df.Close.shift(WINDOW_MOM).iloc[-1])
                     / df.Close.shift(WINDOW_MOM).iloc[-1],
        "rsi14":    last["RSI_14"],
        "macd":     last["MACD_12_26_9"],
        "macd_sig": last["MACDs_12_26_9"],
        "bb_up":    last["BBU_20_2.0"],
        "bb_mid":   last["BBM_20_2.0"],
        "bb_low":   last["BBL_20_2.0"],
        "bb_bw":    last["BBB_20_2.0"],
        "bb_pctb":  last["BBP_20_2.0"],
        "atr14":    last["ATR_14"],
        "obv":      last["OBV"],
        "vol_chg":  (last.Volume - prev.Volume) / prev.Volume,
        "vwap":     last["vwap"],
        "volat":    last["volat"],
    }

# ─── METRICS CALCULATION ───────────────────────────────────────────────────────
def compute_metrics(buffer: deque) -> tuple[float, float, float]:
    """Return (accuracy, precision, recall) over the last WINDOW entries."""
    n = len(buffer)
    if n == 0:
        return 0.0, 0.0, 0.0
    tp = sum(1 for t, p in buffer if p == 1 and t == 1)
    tn = sum(1 for t, p in buffer if p == 0 and t == 0)
    fp = sum(1 for t, p in buffer if p == 1 and t == 0)
    fn = sum(1 for t, p in buffer if p == 0 and t == 1)
    acc  = (tp + tn) / n
    prec = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    rec  = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    return acc, prec, rec

# ─── MAIN LOOP (tqdm-wrapped) ──────────────────────────────────────────────────
for iteration in tqdm(itertools.count(1), desc="Signal iterations"):
    # ── 1) Align to the next exact minute ──────────────────────────────────────
    now      = datetime.datetime.now()
    to_sleep = 60 - now.second - now.microsecond / 1e6
    print(f"DEBUG: Sleeping for {to_sleep:.2f}s to align with minute boundary")
    time.sleep(to_sleep)

    # ── 2) Fetch the two most recent 1-min bars ────────────────────────────────
    df2 = (yf.download(SYMBOL,
                       period="2m",
                       interval="1m",
                       progress=False)
             .dropna())
    if len(df2) < 2:
        print("DEBUG: Insufficient data, retrying…")
        continue
    prev_bar, curr_bar = df2.iloc[-2], df2.iloc[-1]
    ts = curr_bar.name.to_pydatetime()

    # ── 3) Learn from the last signal ─────────────────────────────────────────
    if pending:
        true_lbl = int(curr_bar.Close > pending["price"])
        print(f"DEBUG: Learning from last – pred={pending['pred']}, true={true_lbl}")
        model.learn_one(pending["feats"], true_lbl)
        metric_buffer.append((true_lbl, pending["pred"]))
        acc, prec, rec = compute_metrics(metric_buffer)
        # performance-trend confidence
        if prev_acc is not None:
            perf_conf = max(0.0, min(1.0, 0.5 + (acc - prev_acc)))
        else:
            perf_conf = 0.5
        prev_acc = acc
        print(
            f"[{ts:%H:%M}] METRICS → Acc={acc:.3f}, Prec={prec:.3f}, "
            f"Rec={rec:.3f}, PerfConf={perf_conf:.2f}"
        )
        pending = None

    # ── 4) Update history & recompute indicators ─────────────────────────────
    df_hist = pd.concat([df_hist, curr_bar.to_frame().T]).iloc[-HIST_MINUTES:]
    df_hist.ta.rsi(length=14, append=True)
    df_hist.ta.macd(append=True)
    df_hist.ta.bbands(length=20, std=2, append=True)
    df_hist.ta.atr(length=14, append=True)
    df_hist.ta.obv(append=True)
    df_hist["vwap"]  = (
        df_hist.Volume * (df_hist.High + df_hist.Low + df_hist.Close) / 3
    ).cumsum() / df_hist.Volume.cumsum()
    df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()
    df_hist.dropna(inplace=True)

    # ── 5) Feature extraction & prediction ────────────────────────────────────
    feats  = make_features(df_hist)
    y_pred = model.predict_one(feats) or 0
    p_up   = model.predict_proba_one(feats).get(1, 0.0)
    action = "BUY" if y_pred else "SELL"
    print(f"[{ts:%H:%M}] SIGNAL → {action} | P(up)={p_up:.2f}")

    # ── 6) Simulated execution & sizing ───────────────────────────────────────
    price = curr_bar.Close
    if action == "BUY" and cash > 0:
        qty = cash / price
        pos += qty
        cash = 0.0
        print(f"DEBUG: Executed BUY  {qty:.6f} shares @${price:.2f}")
    elif action == "SELL" and pos > 0:
        cash += pos * price
        print(f"DEBUG: Executed SELL {pos:.6f} shares @${price:.2f}")
        pos = 0.0
    else:
        print(f"DEBUG: HOLD — cash=${cash:.2f}, pos={pos:.6f}")

    # ── 7) Explain via top-2 feature importances ──────────────────────────────
    try:
        imps = model.feature_importances_
        top2 = sorted(imps.items(), key=lambda x: abs(x[1]), reverse=True)[:2]
        reason = "; ".join(f"{n}→{w:.3f}" for n, w in top2)
    except Exception:
        reason = "—"
    print(f"Reason: {reason}\n")

    # ── 8) Stash for next-minute feedback ─────────────────────────────────────
    pending = {"feats": feats, "price": price, "pred": y_pred}