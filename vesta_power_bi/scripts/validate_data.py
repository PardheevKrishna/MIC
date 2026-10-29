"""Validate the synthetic Vesta CSVs against every rule in docs/SPEC_transcribed.md.

Usage:  python scripts/validate_data.py
Prints one PASS/FAIL line per check (plus WARN lines for soft observations) and exits 1 on any FAIL.
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PARENT = ROOT / "data" / "vesta_Document.csv"
CHILD = ROOT / "data" / "vesta_DocumentVersion.csv"

results: list[tuple[str, bool, str]] = []
warnings: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    results.append((name, bool(ok), detail))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  -- {detail}" if detail else ""))


def warn(msg: str) -> None:
    warnings.append(msg)
    print(f"WARN  {msg}")


def nb(s: pd.Series) -> pd.Series:
    """Nonblank mask."""
    return s.str.strip() != ""


def load(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str, keep_default_na=False)


p = load(PARENT)
c = load(CHILD)

PARENT_COLS = ["document_id", "document_type_code", "business_document_id", "created_by", "created_on",
               "last_modified_by", "last_modified_on", "model_id", "is_active", "description"]
CHILD_COLS = ["document_version_id", "document_id", "template_id", "template_version", "version_no", "state",
              "based_on_version_id", "created_by", "created_on", "approved_by", "approved_on", "last_modified_by",
              "last_modified_on", "section_id", "template_driver_mapping", "mar_id", "var_id",
              "committed_to_vesta_by", "committed_to_vesta_on", "var_approverId", "var_ReferenceId",
              "published_date", "published_by", "mar_referenceNumber", "created_by_name",
              "last_modified_by_name", "ack_template_version"]

# ---------------------------------------------------------------- schema
check("parent column set/order", list(p.columns) == PARENT_COLS, str(list(p.columns)))
check("child column set/order", list(c.columns) == CHILD_COLS, str(list(c.columns)))

# whitespace hygiene
for name, df in (("parent", p), ("child", c)):
    bad = [(col, i) for col in df.columns for i, v in df[col].items() if v != v.strip()]
    check(f"{name} no leading/trailing whitespace in values", not bad, f"{len(bad)} cells e.g. {bad[:3]}")

# ---------------------------------------------------------------- row counts
check("parent row count == 131", len(p) == 131, f"got {len(p)}")
check("child row count == 145", len(c) == 145, f"got {len(c)}")

# ---------------------------------------------------------------- integer parsing helpers
INT_RE = re.compile(r"^\d+$")


def all_int(s: pd.Series) -> bool:
    return s[nb(s)].str.match(INT_RE).all()


for col in ["document_id", "business_document_id", "model_id"]:
    check(f"parent {col} integer-formatted", all_int(p[col]))
for col in ["document_version_id", "document_id", "template_id", "template_version", "based_on_version_id",
            "ack_template_version"]:
    check(f"child {col} integer-formatted (when nonblank)", all_int(c[col]))

pid = p["document_id"].astype(int)
cid = c["document_version_id"].astype(int)
c_doc = c["document_id"].astype(int)

# ---------------------------------------------------------------- uniqueness / ranges / gaps
check("parent document_id unique (131 distinct)", pid.is_unique and pid.nunique() == 131, f"{pid.nunique()} distinct")
check("child document_version_id unique (145 distinct)", cid.is_unique and cid.nunique() == 145,
      f"{cid.nunique()} distinct")
check("parent document_id range 5..228", pid.min() == 5 and pid.max() == 228, f"{pid.min()}..{pid.max()}")
check("child document_version_id range 1004..1257", cid.min() == 1004 and cid.max() == 1257,
      f"{cid.min()}..{cid.max()}")
check("parent document_id gapped (non-contiguous)", pid.max() - pid.min() + 1 > len(pid))
check("child document_version_id gapped (non-contiguous)", cid.max() - cid.min() + 1 > len(cid))
check("parent rows sorted ascending by document_id", pid.is_monotonic_increasing)
check("child rows sorted ascending by document_version_id", cid.is_monotonic_increasing)
check("all ids positive", (pid > 0).all() and (cid > 0).all())

# ---------------------------------------------------------------- FK integrity
orphans = sorted(set(c_doc) - set(pid))
check("every child document_id exists in parent (no orphans)", not orphans, f"orphans={orphans[:10]}")
childless = sorted(set(pid) - set(c_doc))
check("every parent document has >=1 version", not childless, f"childless={childless[:10]}")

# ---------------------------------------------------------------- version-count distribution
vc = c_doc.value_counts()
dist = Counter(int(x) for x in vc.values)
expected_dist = {1: 123, 2: 3, 3: 4, 4: 1}
check("version-count distribution 123x1, 3x2, 4x3, 1x4", dict(dist) == expected_dist, f"got {dict(sorted(dist.items()))}")
check("8 documents have history", (vc > 1).sum() == 8, f"got {(vc > 1).sum()}")

# ---------------------------------------------------------------- type split & business id
tc = p["document_type_code"].value_counts().to_dict()
check("document_type_code split MDD 118 / MVD 13", tc == {"MDD": 118, "MVD": 13}, f"got {tc}")
bmap = {"MDD": "100110", "MVD": "100101"}
bad_b = p[p["business_document_id"] != p["document_type_code"].map(bmap)]
check("business_document_id mapping MDD->100110, MVD->100101", bad_b.empty, f"{len(bad_b)} bad rows")

# ---------------------------------------------------------------- distinct counts
def distinct(s: pd.Series) -> int:
    return s[nb(s)].nunique()


for df_name, df, col, exp in [
    ("parent", p, "created_by", 80),
    ("parent", p, "model_id", 118),
    ("parent", p, "description", 47),
    ("child", c, "created_by", 84),
    ("child", c, "template_id", 87),
    ("child", c, "template_version", 35),
    ("child", c, "mar_id", 27),
    ("child", c, "var_id", 23),
    ("child", c, "mar_referenceNumber", 37),
]:
    got = distinct(df[col])
    check(f"{df_name} {col} distinct == {exp}", got == exp, f"got {got}")

# ---------------------------------------------------------------- per-column nonblank counts
parent_nonblank = {"document_id": 131, "document_type_code": 131, "business_document_id": 131, "created_by": 131,
                   "created_on": 131, "last_modified_by": 0, "last_modified_on": 0, "model_id": 131,
                   "is_active": 0, "description": 131}
for col, exp in parent_nonblank.items():
    got = int(nb(p[col]).sum())
    check(f"parent {col} nonblank == {exp}", got == exp, f"got {got}")

child_nonblank = {
    "document_version_id": 145, "document_id": 145, "template_id": 145, "template_version": 145,
    "version_no": 145, "state": 145, "based_on_version_id": 12, "created_by": 145, "created_on": 145,
    "approved_by": 3, "approved_on": 3, "last_modified_by": 144, "last_modified_on": 110,
    "section_id": 0, "template_driver_mapping": 0, "mar_id": 32, "var_id": 40,
    "committed_to_vesta_by": 45, "committed_to_vesta_on": 45, "var_approverId": 2, "var_ReferenceId": 3,
    "published_date": 5, "published_by": 4, "mar_referenceNumber": 42, "created_by_name": 97,
    "last_modified_by_name": 119, "ack_template_version": 8,
}
for col, exp in child_nonblank.items():
    got = int(nb(c[col]).sum())
    check(f"child {col} nonblank == {exp}", got == exp, f"got {got}")

# ---------------------------------------------------------------- lifecycle distribution by state x type
c = c.merge(p[["document_id", "document_type_code", "model_id"]], on="document_id", how="left")
STATES = ["Draft", "Lock", "Retired", "Publish", "Commit to Vesta"]
check("state values within allowed set", set(c["state"]) <= set(STATES), f"got {sorted(set(c['state']))}")
exp_state = {"Draft": (94, 83, 11), "Lock": (31, 31, 0), "Retired": (11, 6, 5), "Publish": (6, 4, 2),
             "Commit to Vesta": (3, 3, 0)}
for st, (tot, mdd, mvd) in exp_state.items():
    sub = c[c["state"] == st]
    g = (len(sub), int((sub["document_type_code"] == "MDD").sum()), int((sub["document_type_code"] == "MVD").sum()))
    check(f"lifecycle {st}: total {tot} / MDD {mdd} / MVD {mvd}", g == (tot, mdd, mvd), f"got {g}")

mvd_bad = c[(c["document_type_code"] == "MVD") & c["state"].isin(["Lock", "Commit to Vesta"])]
check("Lock / Commit to Vesta only on MDD", mvd_bad.empty, f"{len(mvd_bad)} MVD rows")

# ---------------------------------------------------------------- state-dependent workflow table
wf_exp = {  # based_on, mar_id, var_id, committed_on, published_date
    "Draft": (5, 0, 9, 0, 0),
    "Lock": (0, 20, 15, 30, 0),
    "Retired": (4, 6, 10, 6, 0),
    "Publish": (3, 4, 4, 6, 5),
    "Commit to Vesta": (0, 2, 2, 3, 0),
}
for st, exp in wf_exp.items():
    sub = c[c["state"] == st]
    g = tuple(int(nb(sub[col]).sum()) for col in
              ["based_on_version_id", "mar_id", "var_id", "committed_to_vesta_on", "published_date"])
    check(f"workflow {st} (based_on, mar_id, var_id, committed_on, published_date) == {exp}", g == exp, f"got {g}")

# approval only on Publish state, and published fields only on Publish
for col in ["approved_by", "approved_on", "published_by", "published_date"]:
    bad = c[nb(c[col]) & (c["state"] != "Publish")]
    check(f"{col} only on Publish rows", bad.empty, f"{len(bad)} rows in other states")
pub = c[c["state"] == "Publish"]
check("Publish rows always have committed_to_vesta_on", nb(pub["committed_to_vesta_on"]).all())

# ---------------------------------------------------------------- datetime formats & windows
DT_RE = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
DT_COLS_C = ["created_on", "approved_on", "last_modified_on", "committed_to_vesta_on", "published_date"]
for col in DT_COLS_C:
    s = c[col][nb(c[col])]
    check(f"child {col} format YYYY-MM-DD HH:MM:SS", s.str.match(DT_RE).all(),
          f"bad={s[~s.str.match(DT_RE)].tolist()[:3]}")
s = p["created_on"]
check("parent created_on format YYYY-MM-DD HH:MM:SS", s.str.match(DT_RE).all())


def dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s.where(nb(s)), format="%Y-%m-%d %H:%M:%S", errors="coerce")


for col in DT_COLS_C:
    c[col + "_dt"] = dt(c[col])
p["created_on_dt"] = dt(p["created_on"])

lo, hi = pd.Timestamp("2026-06-24 14:13:21"), pd.Timestamp("2026-10-28 10:31:30")
cc = c["created_on_dt"]
check("child created_on min == 2026-06-24 14:13:21", cc.min() == lo, f"got {cc.min()}")
check("child created_on max == 2026-10-28 10:31:30", cc.max() == hi, f"got {cc.max()}")
check("child created_on all within window", cc.between(lo, hi).all())
pc = p["created_on_dt"]
check("parent created_on within child window", pc.between(lo, hi).all(), f"{pc.min()}..{pc.max()}")

lm = c["last_modified_on_dt"].dropna()
check("child last_modified_on min date == 2026-07-10", lm.min().date() == pd.Timestamp("2026-07-10").date(),
      f"got {lm.min()}")
check("child last_modified_on max date == 2026-10-28", lm.max().date() == pd.Timestamp("2026-10-28").date(),
      f"got {lm.max()}")

for col in ["approved_on", "committed_to_vesta_on", "published_date"]:
    s = c[col + "_dt"].dropna()
    in_year = (s.dt.year == 2026).all()
    check(f"child {col} in 2026", in_year, f"{s.min()}..{s.max()}" if len(s) else "none")
    # nothing should be in the future relative to the export window end (spec 'today' is 2026-10-28/29)
    if len(s) and s.max() > pd.Timestamp("2026-10-28 23:59:59"):
        warn(f"{col} has values after 2026-10-28 (max {s.max()})")

lm_mask = c["last_modified_on_dt"].notna()
bad = c[lm_mask & (c["last_modified_on_dt"] < c["created_on_dt"])]
check("last_modified_on >= created_on", bad.empty, f"{len(bad)} rows e.g. {bad['document_version_id'].tolist()[:5]}")

for col in ["approved_on", "committed_to_vesta_on", "published_date"]:
    m = c[col + "_dt"].notna()
    bad = c[m & (c[col + "_dt"] < c["created_on_dt"])]
    check(f"{col} >= created_on", bad.empty, f"{len(bad)} rows e.g. {bad['document_version_id'].tolist()[:5]}")

m = c["approved_on_dt"].notna() & c["published_date_dt"].notna()
bad = c[m & (c["approved_on_dt"] > c["published_date_dt"])]
if not bad.empty:
    warn(f"approved_on after published_date on {bad['document_version_id'].tolist()}")
m = c["committed_to_vesta_on_dt"].notna() & c["published_date_dt"].notna()
bad = c[m & (c["published_date_dt"] < c["committed_to_vesta_on_dt"])]
if not bad.empty:
    warn(f"published_date before committed_to_vesta_on on {bad['document_version_id'].tolist()}")

# ---------------------------------------------------------------- pairing rules
def paired(a: str, b: str) -> tuple[bool, str]:
    diff = c[nb(c[a]) != nb(c[b])]
    return diff.empty, f"{len(diff)} mismatched rows e.g. {diff['document_version_id'].tolist()[:5]}"


ok, d = paired("committed_to_vesta_by", "committed_to_vesta_on")
check("committed_to_vesta_by <-> committed_to_vesta_on paired", ok, d)
ok, d = paired("approved_by", "approved_on")
check("approved_by <-> approved_on paired", ok, d)
bad = c[nb(c["published_by"]) & ~nb(c["published_date"])]
check("published_by only with published_date", bad.empty, f"{len(bad)} rows")
bad = c[(nb(c["approved_by"]) | nb(c["approved_on"])) & ~nb(c["published_date"])]
check("approved_by/on only with nonblank published_date", bad.empty, f"{len(bad)} rows")
bad = c[nb(c["last_modified_on"]) & ~nb(c["last_modified_by"])]
check("last_modified_on implies last_modified_by", bad.empty, f"{len(bad)} rows")
bad = c[nb(c["last_modified_by_name"]) & ~nb(c["last_modified_by"])]
check("last_modified_by_name implies last_modified_by", bad.empty, f"{len(bad)} rows")
bad = c[nb(c["var_approverId"]) & ~nb(c["var_id"])]
if not bad.empty:
    warn(f"var_approverId without var_id on {bad['document_version_id'].tolist()}")
bad = c[nb(c["var_ReferenceId"]) & ~nb(c["var_id"])]
if not bad.empty:
    warn(f"var_ReferenceId without var_id on {bad['document_version_id'].tolist()}")

# ---------------------------------------------------------------- based_on references
c["_cid"] = c["document_version_id"].astype(int)
by_id = c.set_index("_cid")
bo = c[nb(c["based_on_version_id"])]
bad_ref = []
for _, r in bo.iterrows():
    ref = int(r["based_on_version_id"])
    if ref not in by_id.index:
        bad_ref.append((r["document_version_id"], "missing"))
        continue
    tgt = by_id.loc[ref]
    if tgt["document_id"] != r["document_id"]:
        bad_ref.append((r["document_version_id"], "other doc"))
    elif not (tgt["created_on_dt"] < r["created_on_dt"]):
        bad_ref.append((r["document_version_id"], "not earlier"))
    elif ref == int(r["document_version_id"]):
        bad_ref.append((r["document_version_id"], "self"))
check("based_on_version_id references an earlier version of the same document", not bad_ref, str(bad_ref[:5]))
multi_docs = set(vc[vc > 1].index)
non_first = 0
for d_id in multi_docs:
    non_first += vc[d_id] - 1
check("based_on present on some but not all successive records",
      0 < len(bo) < non_first, f"based_on={len(bo)}, successive rows={non_first}")
bo_on_first = []
for d_id, g in c.groupby("document_id"):
    g = g.sort_values("created_on_dt")
    first = g.iloc[0]
    if first["based_on_version_id"] != "":
        bo_on_first.append(first["document_version_id"])
check("based_on_version_id never on a document's first version", not bo_on_first, str(bo_on_first[:5]))
single_bo = c[nb(c["based_on_version_id"]) & c["document_id"].astype(int).map(vc).eq(1)]
check("based_on_version_id never on single-version documents", single_bo.empty, f"{len(single_bo)} rows")

# ---------------------------------------------------------------- version_no
VNO_RE = re.compile(r"^2026\.\d+(\.\d+)?$")
check("version_no matches ^2026\\.\\d+(\\.\\d+)?$", c["version_no"].str.match(VNO_RE).all(),
      f"bad={c.loc[~c['version_no'].str.match(VNO_RE), 'version_no'].tolist()[:5]}")
SIX = {"2026.1", "2026.1.1", "2026.1.2", "2026.1.3", "2026.2.1", "2026.2.2"}
check("version_no uses exactly the six spec values", set(c["version_no"]) == SIX,
      f"got {sorted(set(c['version_no']))}")


def vkey(v: str) -> tuple[int, ...]:
    parts = [int(x) for x in v.split(".")]
    return tuple(parts + [0] * (3 - len(parts)))


# ---------------------------------------------------------------- chains
chain_bad = []
label_dup = []
id_order = []
paths = {}
for d_id, g in c.groupby("document_id"):
    g = g.sort_values("created_on_dt")
    ts = g["created_on_dt"].tolist()
    if any(not (b > a) for a, b in zip(ts, ts[1:])):
        chain_bad.append(d_id)
    labels = g["version_no"].tolist()
    if len(set(labels)) != len(labels):
        label_dup.append(d_id)
    keys = [vkey(x) for x in labels]
    if any(not (b > a) for a, b in zip(keys, keys[1:])):
        chain_bad.append(f"{d_id}:label-order")
    ids = g["_cid"].tolist()
    if ids != sorted(ids):
        id_order.append(d_id)
    if len(g) > 1:
        paths[d_id] = " -> ".join(g["state"].tolist())
check("multi-version chains: created_on strictly increasing & labels increasing", not chain_bad, str(chain_bad[:5]))
check("version labels unique within each document", not label_dup, str(label_dup[:5]))
if id_order:
    warn(f"document_version_id order disagrees with created_on order for docs {id_order}")
else:
    check("document_version_id ascends with created_on within each document", True)
four = [p_ for d, p_ in paths.items() if vc[int(d)] == 4]
check("four-version chain is Retired -> Publish -> Retired -> Draft",
      four == ["Retired -> Publish -> Retired -> Draft"], str(four))
print("INFO  chain paths:", "; ".join(f"{k}: {v}" for k, v in sorted(paths.items())))
for d, path in paths.items():
    states = path.split(" -> ")
    # Only the last version of a doc should normally be "current"; earlier versions being Draft/Lock is odd
    for s_ in states[:-1]:
        if s_ in ("Draft", "Lock", "Commit to Vesta"):
            warn(f"doc {d} has non-terminal earlier version in state {s_}: {path}")
            break

# ---------------------------------------------------------------- parent created == first version
first = c.sort_values("created_on_dt").groupby("document_id").first()
pm = p.set_index("document_id")
mis_on = [d for d in pm.index if pm.loc[d, "created_on"] != first.loc[d, "created_on"]]
mis_by = [d for d in pm.index if pm.loc[d, "created_by"] != first.loc[d, "created_by"]]
check("parent created_on == first version created_on", not mis_on, str(mis_on[:5]))
check("parent created_by == first version created_by", not mis_by, str(mis_by[:5]))

# ---------------------------------------------------------------- document_id vs model_id
mids = set(p["model_id"].astype(int))
clash = sorted(set(pid) & mids)
check("document_id never equal to any model_id", not clash, str(clash[:5]))
# model reuse: expect pairs preferably MDD/MVD
mcount = p["model_id"].value_counts()
reused = mcount[mcount > 1]
check("13 model_id values reused (131-118)", int((mcount - 1).sum()) == 13, f"extra rows={int((mcount - 1).sum())}")
pair_types = []
for m_, n in reused.items():
    types = sorted(p.loc[p["model_id"] == m_, "document_type_code"].tolist())
    pair_types.append(tuple(types))
same_type = [t for t in pair_types if len(set(t)) == 1]
if same_type:
    warn(f"{len(same_type)} reused model_ids share the same doc type: {same_type}")
print(f"INFO  reused model ids: {Counter(pair_types)}")
mvd_models = set(p.loc[p["document_type_code"] == "MVD", "model_id"])
mdd_models = set(p.loc[p["document_type_code"] == "MDD", "model_id"])
lone_mvd = mvd_models - mdd_models
if lone_mvd:
    warn(f"{len(lone_mvd)} MVD documents whose model has no MDD: {sorted(lone_mvd)}")
# MVD should not predate the MDD of the same model (soft)
for m_ in mvd_models & mdd_models:
    rows = p[p["model_id"] == m_]
    mdd_on = rows.loc[rows["document_type_code"] == "MDD", "created_on_dt"].min()
    mvd_on = rows.loc[rows["document_type_code"] == "MVD", "created_on_dt"].min()
    if mvd_on < mdd_on:
        warn(f"model {m_}: MVD created before MDD")

# ---------------------------------------------------------------- description
desc_bad = p[[not d.startswith(f"{t} created from template: {t} |") for d, t in
              zip(p["description"], p["document_type_code"])]]
check("description starts with '<type> created from template: <type> |'", desc_bad.empty, f"{len(desc_bad)} rows")
check("description pipe-delimited with 4 segments",
      p["description"].str.split(": ", n=1).str[1].str.split(" | ", regex=False).str.len().eq(4).all())

# ---------------------------------------------------------------- display names per user id
def name_consistency(id_col: str, name_col: str) -> dict:
    sub = c[nb(c[id_col]) & nb(c[name_col])]
    g = sub.groupby(id_col)[name_col].nunique()
    return g[g > 1].to_dict()


inc1 = name_consistency("created_by", "created_by_name")
inc2 = name_consistency("last_modified_by", "last_modified_by_name")
check("created_by_name consistent per created_by", not inc1, str(inc1))
check("last_modified_by_name consistent per last_modified_by", not inc2, str(inc2))
allmap = pd.concat([
    c.loc[nb(c["created_by_name"]), ["created_by", "created_by_name"]].set_axis(["uid", "nm"], axis=1),
    c.loc[nb(c["last_modified_by_name"]), ["last_modified_by", "last_modified_by_name"]].set_axis(["uid", "nm"], axis=1),
])
g = allmap.groupby("uid")["nm"].nunique()
check("display name consistent per user id across both name columns", (g <= 1).all(), str(g[g > 1].to_dict()))
g2 = allmap.groupby("nm")["uid"].nunique()
check("each display name maps to one user id", (g2 <= 1).all(), str(g2[g2 > 1].to_dict()))

# ---------------------------------------------------------------- value formats
UID_RE = re.compile(r"^[a-z]\d{5,6}$", re.I)
uid_cols = [("parent", p, "created_by")] + [("child", c, col) for col in
            ["created_by", "approved_by", "last_modified_by", "committed_to_vesta_by", "published_by"]]
for nm, df, col in uid_cols:
    s = df[col][nb(df[col])]
    bad = s[~s.str.match(UID_RE)]
    check(f"{nm} {col} format letter+5-6 digits", bad.empty, f"bad={bad.tolist()[:5]}")
s = c["var_approverId"][nb(c["var_approverId"])]
if not s.str.match(UID_RE).all():
    warn(f"var_approverId not user-id-like: {s.tolist()}")

def fmt_check(label: str, s: pd.Series, pattern: str) -> None:
    s = s[nb(s)]
    ok = s.str.match(pattern)
    check(label, ok.all(), f"bad={s[~ok].tolist()[:5]}")


fmt_check("mar_referenceNumber format MA+digits", c["mar_referenceNumber"], r"^MA\d+$")
fmt_check("var_id format digits or P+digits", c["var_id"], r"^P?\d+$")
fmt_check("mar_id numeric string or letters+digits (e.g. CM500120)", c["mar_id"], r"^([A-Z]{2})?\d+$")
s = c["mar_id"][nb(c["mar_id"])]
print(f"INFO  mar_id prefixed: {int(s.str.match(r'^[A-Z]').sum())}/{len(s)}; "
      f"var_id P-prefixed: {int(c['var_id'][nb(c['var_id'])].str.startswith('P').sum())}/{int(nb(c['var_id']).sum())}")
s = c["var_ReferenceId"][nb(c["var_ReferenceId"])]
print(f"INFO  var_ReferenceId values: {s.tolist()}; var_approverId: {c['var_approverId'][nb(c['var_approverId'])].tolist()}")

tv = c["template_version"].astype(int)
check("template_version low positive integers", (tv >= 1).all() and tv.max() <= 100, f"{tv.min()}..{tv.max()}")
atv = c["ack_template_version"][nb(c["ack_template_version"])].astype(int)
check("ack_template_version small positive integers", (atv >= 1).all() and atv.max() <= 100,
      f"{atv.tolist()}")
for col in ["created_by_name", "last_modified_by_name"]:
    fmt_check(f"{col} looks like 'First Last'", c[col], r"^[A-Z][a-zA-Z'\-]+ [A-Z][a-zA-Z'\-]+$")

# mar_referenceNumber vs mar_id relationship (soft)
both = c[nb(c["mar_id"]) & nb(c["mar_referenceNumber"])]
print(f"INFO  rows with mar_id & mar_referenceNumber: {len(both)}; mar_ref without mar_id: "
      f"{int((nb(c['mar_referenceNumber']) & ~nb(c['mar_id'])).sum())}")

# template_id consistency per document (soft): versions of the same doc normally share a template
tpl_var = c.groupby("document_id")["template_id"].nunique()
if (tpl_var > 1).any():
    print(f"INFO  docs whose versions use different template_id: {int((tpl_var > 1).sum())}")

# initial version label conventions (soft)
firsts = c.sort_values("created_on_dt").groupby("document_id").first()
init = Counter(zip(firsts["document_type_code"], firsts["version_no"]))
print(f"INFO  initial version labels by type: {dict(init)}")
mdd_init = Counter(firsts.loc[firsts["document_type_code"] == "MDD", "version_no"])
if mdd_init and mdd_init.most_common(1)[0][0] != "2026.1.1":
    warn(f"initial MDD label most common is not 2026.1.1: {mdd_init}")

# ---------------------------------------------------------------- summary
fails = [r for r in results if not r[1]]
print()
print(f"SUMMARY: {len(results) - len(fails)} PASS, {len(fails)} FAIL, {len(warnings)} WARN")
sys.exit(1 if fails else 0)
