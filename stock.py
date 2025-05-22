# ─── Monkey-patch NumPy so pandas_ta can import NaN ─────────────────────────────
import numpy as np
if not hasattr(np, "NaN"):
    np.NaN = np.nan

import time
import datetime
import itertools

import pandas as pd
import yfinance as yf
import pandas_ta as ta
from tqdm import tqdm

from river import preprocessing, metrics, ensemble, tree

# ─── PARAMETERS ────────────────────────────────────────────────────────────────
SYMBOL        = "AAPL"
HIST_MINUTES  = 50
WINDOW_MOM    = 5
BUDGET        = 10.0

# ─── ONLINE MODEL & ROLLING METRICS ─────────────────────────────────────────────
model       = (
    preprocessing.StandardScaler()
    | ensemble.BaggingClassifier(
        model=tree.HoeffdingTreeClassifier(),  # base learner
        n_models=15,
        seed=42
      )
)
metric_acc  = metrics.Rolling(metrics.Accuracy(),  window_size=200)
metric_prec = metrics.Rolling(metrics.Precision(), window_size=200)
metric_rec  = metrics.Rolling(metrics.Recall(),    window_size=200)
prev_acc    = None  # for trend-confidence

# ─── WARM-UP HISTORY ────────────────────────────────────────────────────────────
print("DEBUG: Downloading warm-up history…")
df_hist = (
    yf.download(SYMBOL, period=f"{HIST_MINUTES}m", interval="1m", progress=False)
      .dropna()
)
df_hist.ta.rsi(length=14, append=True)
df_hist.ta.macd(append=True)
df_hist.ta.bbands(length=20, std=2, append=True)
df_hist.ta.atr(length=14, append=True)
df_hist.ta.obv(append=True)
df_hist["vwap"]  = (df_hist.Volume * (df_hist.High + df_hist.Low + df_hist.Close) / 3).cumsum() \
                   / df_hist.Volume.cumsum()
df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()
df_hist.dropna(inplace=True)
print(f"DEBUG: Warm-up complete, {len(df_hist)} bars loaded.")

pending = None
cash, pos = BUDGET, 0.0

# ─── FEATURE EXTRACTION ────────────────────────────────────────────────────────
def make_features(df: pd.DataFrame) -> dict:
    last, prev = df.iloc[-1], df.iloc[-2]
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

# ─── MAIN LOOP (tqdm-wrapped) ───────────────────────────────────────────────────
for iteration in tqdm(itertools.count(1), desc="Signal iterations"):
    now      = datetime.datetime.now()
    to_sleep = 60 - now.second - now.microsecond / 1e6
    print(f"DEBUG: Sleeping {to_sleep:.2f}s")
    time.sleep(to_sleep)

    df2 = yf.download(SYMBOL, period="2m", interval="1m", progress=False).dropna()
    if len(df2) < 2:
        print("DEBUG: Insufficient data, retrying…")
        continue
    prev_bar, curr_bar = df2.iloc[-2], df2.iloc[-1]
    ts = curr_bar.name.to_pydatetime()

    # ── Learn from last prediction ──────────────────────────────────────────────
    if pending:
        true_lbl = int(curr_bar.Close > pending["price"])
        print(f"DEBUG: Learning – pred={pending['pred']}, true={true_lbl}")
        model.learn_one(pending["feats"], true_lbl)
        metric_acc.update(true_lbl, pending["pred"])
        metric_prec.update(true_lbl, pending["pred"])
        metric_rec.update(true_lbl, pending["pred"])
        curr_acc  = metric_acc.get()
        curr_prec = metric_prec.get()
        curr_rec  = metric_rec.get()
        if prev_acc is not None:
            diff      = curr_acc - prev_acc
            perf_conf = max(0.0, min(1.0, 0.5 + diff))
        else:
            perf_conf = 0.5
        prev_acc = curr_acc
        print(f"[{ts:%H:%M}] METRICS → Acc={curr_acc:.3f}, Prec={curr_prec:.3f}, "
              f"Rec={curr_rec:.3f}, PerfConf={perf_conf:.2f}")
        pending = None

    # ── Update history & recompute indicators ────────────────────────────────────
    df_hist = pd.concat([df_hist, curr_bar.to_frame().T]).iloc[-HIST_MINUTES:]
    df_hist.ta.rsi(length=14, append=True)
    df_hist.ta.macd(append=True)
    df_hist.ta.bbands(length=20, std=2, append=True)
    df_hist.ta.atr(length=14, append=True)
    df_hist.ta.obv(append=True)
    df_hist["vwap"]  = (df_hist.Volume * (df_hist.High + df_hist.Low + df_hist.Close) / 3).cumsum() \
                       / df_hist.Volume.cumsum()
    df_hist["volat"] = df_hist.Close.pct_change().rolling(14).std()
    df_hist.dropna(inplace=True)

    # ── Predict & execute ───────────────────────────────────────────────────────
    feats  = make_features(df_hist)
    y_pred = model.predict_one(feats) or 0
    p_up   = model.predict_proba_one(feats).get(1, 0.0)
    action = "BUY" if y_pred else "SELL"
    print(f"[{ts:%H:%M}] SIGNAL → {action} | P(up)={p_up:.2f}")

    price = curr_bar.Close
    if action == "BUY" and cash > 0:
        qty = cash / price
        pos += qty
        cash = 0.0
        print(f"DEBUG: Executed BUY {qty:.6f} shares @${price:.2f}")
    elif action == "SELL" and pos > 0:
        cash += pos * price
        print(f"DEBUG: Executed SELL {pos:.6f} shares @${price:.2f}")
        pos = 0.0
    else:
        print(f"DEBUG: HOLD — cash=${cash:.2f}, pos={pos:.6f}")

    # ── Explain via top-2 feature importances ────────────────────────────────────
    try:
        imps = model.feature_importances_
        top2 = sorted(imps.items(), key=lambda x: abs(x[1]), reverse=True)[:2]
        reason = "; ".join(f"{n}→{w:.3f}" for n, w in top2)
    except:
        reason = "—"
    print(f"Reason: {reason}\n")

    # ── Stash for the next update ────────────────────────────────────────────────
    pending = {"feats": feats, "price": price, "pred": y_pred}