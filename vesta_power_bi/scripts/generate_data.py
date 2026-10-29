"""Generate synthetic vesta_Document / vesta_DocumentVersion CSVs.

Follows docs/SPEC_transcribed.md (the README reconstruction guide). Every count in the spec is hit
exactly; all IDs, users and names are invented. Deterministic for a given SEED.
"""
from __future__ import annotations

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

SEED = 2026
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data"

WINDOW_START = datetime(2026, 6, 24, 14, 13, 21)
WINDOW_END = datetime(2026, 10, 28, 10, 31, 30)
LM_MIN_DAY = datetime(2026, 7, 10, 8, 0, 0)
LM_MAX = datetime(2026, 10, 28, 17, 45, 0)
FMT = "%Y-%m-%d %H:%M:%S"

rng = random.Random(SEED)

# --------------------------------------------------------------------------- users
FIRST = ["Jordan", "Avery", "Riley", "Morgan", "Casey", "Taylor", "Quinn", "Reese", "Skyler", "Rowan",
         "Emerson", "Finley", "Harper", "Kendall", "Logan", "Parker", "Sawyer", "Dakota", "Hayden", "Jamie",
         "Alexis", "Blair", "Cameron", "Drew", "Elliot", "Frankie", "Gray", "Indigo", "Jesse", "Kai"]
LAST = ["Lee", "Kim", "Park", "Rivera", "Nguyen", "Patel", "Shah", "Okafor", "Lindqvist", "Moreau",
        "Tanaka", "Haddad", "Novak", "Castillo", "Brennan", "Iyer", "Mensah", "Kowalski", "Duarte", "Asante",
        "Varga", "Ortiz", "Holm", "Rahman", "Silva", "Achebe", "Laurent", "Sato", "Fischer", "Abara"]


def make_users(n: int, taken: set[str]) -> list[str]:
    out = []
    while len(out) < n:
        letter = rng.choice("ukaezr")
        digits = rng.choice([5, 6, 6, 6])
        uid = letter + str(rng.randint(10 ** (digits - 1), 10 ** digits - 1))
        if uid not in taken:
            taken.add(uid)
            out.append(uid)
    return out


taken_ids: set[str] = set()
CREATORS = make_users(80, taken_ids)         # parent creators (exactly 80 distinct)
SUCCESSOR_CREATORS = make_users(4, taken_ids)  # +4 -> 84 distinct version creators
REVIEWERS = make_users(14, taken_ids)        # modifiers / committers / approvers
ALL_USERS = CREATORS + SUCCESSOR_CREATORS + REVIEWERS
names = set()
USER_NAME: dict[str, str] = {}
for uid in ALL_USERS:
    while True:
        nm = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
        if nm not in names:
            names.add(nm)
            USER_NAME[uid] = nm
            break

# ------------------------------------------------------------------- taxonomy / templates
MODEL_TYPES = ["Quantitative Model", "Qualitative Model", "AI/ML Model", "Statistical Model"]
RISKS = ["Low", "Moderate", "High"]
CATEGORIES = ["General", "Vendor", "Credit", "Market", "Operational", "Regulatory"]
all_tax = [(m, r, c) for m in MODEL_TYPES for r in RISKS for c in CATEGORIES]
weights = [(3 if m == "Quantitative Model" else 1) * (2 if r == "Moderate" else 1) for m, r, c in all_tax]
mdd_tax: list[tuple] = []
while len(mdd_tax) < 38:  # 38 MDD descriptions + 9 MVD descriptions = 47 distinct
    t = rng.choices(all_tax, weights)[0]
    if t not in mdd_tax:
        mdd_tax.append(t)


def desc(doc_type: str, tax: tuple) -> str:
    m, r, c = tax
    return f"{doc_type} created from template: {doc_type} | {m} | {r} risk | {c}"


template_ids = sorted(rng.sample(range(12, 160), 87))
rng.shuffle(template_ids)
mdd_templates = template_ids[:76]
mvd_template_pool = template_ids[76:]
tv_values = sorted(rng.sample(range(1, 41), 35))  # 35 distinct low positive revisions

TEMPLATE_VERSION = {}
for i, t in enumerate(mdd_templates + mvd_template_pool):
    TEMPLATE_VERSION[t] = tv_values[i] if i < len(tv_values) else rng.choice(tv_values)
TEMPLATE_TAX = {t: mdd_tax[i // 2] for i, t in enumerate(mdd_templates)}  # 2 templates per description

# ----------------------------------------------------------------------- MDD documents
span = (WINDOW_END - WINDOW_START).total_seconds()


def rand_time(lo: datetime, hi: datetime) -> datetime:
    sec = rng.uniform(0, (hi - lo).total_seconds())
    return min(max(business(lo + timedelta(seconds=sec)), lo), hi).replace(microsecond=0)


def business(t: datetime) -> datetime:
    """Move a timestamp onto a weekday between 08:00 and 18:00 (a bank workflow tool)."""
    if t.weekday() >= 5:
        t += timedelta(days=7 - t.weekday())
    if t.hour < 8 or t.hour > 17:
        t = t.replace(hour=rng.randint(8, 17))
    return t


docs: list[dict] = []
mdd_template_assign = mdd_templates + [rng.choice(mdd_templates) for _ in range(118 - 76)]
rng.shuffle(mdd_template_assign)
mdd_creators = CREATORS + [rng.choice(CREATORS) for _ in range(118 - 80)]
rng.shuffle(mdd_creators)
model_ids = sorted(rng.sample(range(31001, 31800), 118))
for i in range(118):
    t = mdd_template_assign[i]
    docs.append({
        "type": "MDD", "template_id": t, "tax": TEMPLATE_TAX[t], "created_by": mdd_creators[i],
        "created_on": rand_time(WINDOW_START, WINDOW_END - timedelta(days=1)),
    })
docs.sort(key=lambda d: d["created_on"])
docs[0]["created_on"] = WINDOW_START
for d, m in zip(docs, model_ids):  # model ids ascend with creation order
    d["model_id"] = m

# ------------------------------------------------------------ MVD documents (paired)
cutoff = datetime(2026, 10, 5)
candidates = [d for d in docs if d["created_on"] < cutoff]
rng.shuffle(candidates)
paired: list[dict] = []
seen_tax: list[tuple] = []
for d in candidates:  # 9 distinct taxonomies first
    if d["tax"] not in seen_tax and len(seen_tax) < 9:
        seen_tax.append(d["tax"])
        paired.append(d)
for d in candidates:  # then 4 more re-using those taxonomies
    if len(paired) == 13:
        break
    if d not in paired and d["tax"] in seen_tax:
        paired.append(d)
assert len(paired) == 13 and len({d["tax"] for d in paired}) == 9

mvd_docs = []
for d in paired:
    lo = d["created_on"] + timedelta(days=1)
    hi = min(d["created_on"] + timedelta(days=21), WINDOW_END - timedelta(days=2))
    mvd_docs.append({
        # validation is written by a different person than the development document
        "type": "MVD", "tax": d["tax"], "created_by": rng.choice([c for c in CREATORS if c != d["created_by"]]),
        "model_id": d["model_id"],
        "created_on": rand_time(lo, hi), "pair": d,
    })
# MVD templates: one per description, plus 2 extra for descriptions with >1 doc -> 11 templates
by_tax: dict[tuple, list] = {}
for d in mvd_docs:
    by_tax.setdefault(d["tax"], []).append(d)
pool = list(mvd_template_pool)
for tax, group in by_tax.items():
    base = pool.pop()
    for d in group:
        d["template_id"] = base
for d in [d for g in by_tax.values() for d in g[1:]][:2]:  # 2 extra templates -> 11 MVD templates
    d["template_id"] = pool.pop()
assert not pool
docs.extend(mvd_docs)
for d in docs:
    d["description"] = desc(d["type"], d["tax"])
    d["template_version"] = TEMPLATE_VERSION[d["template_id"]]

# ---------------------------------------------------------------- version chains
CHAINS = {
    # name: (doc type, [(state, label), ...])
    "MVD4": ("MVD", [("Retired", "2026.1"), ("Publish", "2026.1.1"), ("Retired", "2026.1.2"), ("Draft", "2026.2.1")]),
    "MVD3": ("MVD", [("Retired", "2026.1"), ("Retired", "2026.1.1"), ("Draft", "2026.2.1")]),
    "A3": ("MDD", [("Retired", "2026.1.1"), ("Publish", "2026.1.2"), ("Draft", "2026.2.1")]),
    "B3": ("MDD", [("Retired", "2026.1.1"), ("Retired", "2026.1.2"), ("Publish", "2026.1.3")]),
    "C3": ("MDD", [("Retired", "2026.1.1"), ("Retired", "2026.2.1"), ("Draft", "2026.2.2")]),
    "D2": ("MDD", [("Publish", "2026.1.1"), ("Draft", "2026.2.1")]),
    "E2": ("MDD", [("Publish", "2026.1.1"), ("Lock", "2026.1.2")]),
    "F2": ("MDD", [("Retired", "2026.1.1"), ("Lock", "2026.1.2")]),
}
chain_cutoff = datetime(2026, 8, 31)
early_mvd = [d for d in docs if d["type"] == "MVD" and d["created_on"] < chain_cutoff]
early_mdd = [d for d in docs if d["type"] == "MDD" and d["created_on"] < chain_cutoff]
rng.shuffle(early_mvd)
rng.shuffle(early_mdd)
for name, (typ, steps) in CHAINS.items():
    d = (early_mvd if typ == "MVD" else early_mdd).pop()
    d["chain"] = steps

singles = {
    "MDD": ["Draft"] * 80 + ["Lock"] * 29 + ["Commit to Vesta"] * 3,
    "MVD": ["Draft"] * 9 + ["Retired"] * 1 + ["Publish"] * 1,
}
for typ in ("MDD", "MVD"):
    pool_docs = [d for d in docs if d["type"] == typ and "chain" not in d]
    states = singles[typ]
    assert len(pool_docs) == len(states), (typ, len(pool_docs), len(states))
    # later lifecycle states skew towards older documents
    pool_docs.sort(key=lambda d: d["created_on"])
    advanced = [s for s in states if s != "Draft"]
    drafts = [s for s in states if s == "Draft"]
    adv_slots = set(sorted(rng.sample(range(int(len(pool_docs) * 0.85)), len(advanced))))
    rng.shuffle(advanced)
    ai = iter(advanced)
    di = iter(drafts)
    for i, d in enumerate(pool_docs):
        st = next(ai) if i in adv_slots else next(di)
        label = "2026.1" if typ == "MVD" else ("2026.1" if rng.random() < 0.05 else "2026.1.1")
        if typ == "MVD" and rng.random() < 0.2:
            label = "2026.1.1"
        d["chain"] = [(st, label)]

# an MVD that has progressed (published/retired) should belong to a model whose MDD has left Draft
ADVANCED = ("Lock", "Commit to Vesta", "Publish")
for mvd in [d for d in docs if d["type"] == "MVD" and any(st != "Draft" for st, _ in d["chain"])]:
    mdd = mvd["pair"]
    if len(mdd["chain"]) == 1 and mdd["chain"][0][0] == "Draft":
        paired_ids = {id(d["pair"]) for d in docs if d["type"] == "MVD"}
        swap = [d for d in docs if d["type"] == "MDD" and len(d["chain"]) == 1 and d["chain"][0][0] in ADVANCED
                and id(d) not in paired_ids and d["created_on"] < mvd["created_on"]]
        other = rng.choice(swap)
        mdd["chain"], other["chain"] = other["chain"], mdd["chain"]

# the latest-created document marks the end of the window
last_single = max((d for d in docs if len(d["chain"]) == 1), key=lambda d: d["created_on"])
last_single["created_on"] = WINDOW_END

# ------------------------------------------------------------------ version rows
versions: list[dict] = []
for d in docs:
    prev = None
    t = d["created_on"]
    for seq, (state, label) in enumerate(d["chain"]):
        if seq == 0:
            creator = d["created_by"]
        else:
            gap = timedelta(days=rng.uniform(2, 18), hours=rng.uniform(0, 6))
            t = min(max(business(prev["created_on"] + gap), LM_MIN_DAY + timedelta(days=1, hours=2, minutes=13, seconds=8)),
                    WINDOW_END - timedelta(hours=3))
            t = t.replace(microsecond=0)
            creator = d["created_by"]
        v = {"doc": d, "seq": seq, "state": state, "version_no": label, "created_by": creator,
             "created_on": t, "template_id": d["template_id"], "template_version": d["template_version"],
             "prev": prev}
        versions.append(v)
        prev = v
assert len(versions) == 145
# four successor versions were created by someone other than the parent's creator -> 84 distinct
successors = [v for v in versions if v["seq"] > 0]
for v, uid in zip(rng.sample(successors, 4), SUCCESSOR_CREATORS):
    v["created_by"] = uid

# predecessor rows received their last edit shortly before the successor version was opened
for v in versions:
    if v["prev"] is not None:
        p = v["prev"]
        lo = max(p["created_on"] + timedelta(hours=1), LM_MIN_DAY)
        hi = v["created_on"] - timedelta(minutes=rng.randint(25, 300))
        p["lm_forced"] = rand_time(lo, max(lo, hi))

# based_on: successor rows in Draft/Retired/Publish (12 rows); Lock successors have none
for v in versions:
    v["based_on"] = v["prev"] if (v["prev"] is not None and v["state"] in ("Draft", "Retired", "Publish")) else None

# ---------------------------------------------------------------- workflow quotas
def pick(rows: list[dict], n: int, prefer=None) -> list[dict]:
    rows = list(rows)
    rng.shuffle(rows)
    if prefer:
        rows.sort(key=lambda r: 0 if prefer(r) else 1)
    assert len(rows) >= n
    return rows[:n]


by_state: dict[str, list] = {}
for v in versions:
    by_state.setdefault(v["state"], []).append(v)

QUOTA = {  # state: (mar_id, var_id, committed, published)
    "Draft": (0, 9, 0, 0),
    "Lock": (20, 15, 30, 0),
    "Retired": (6, 10, 6, 0),
    "Publish": (4, 4, 6, 5),
    "Commit to Vesta": (2, 2, 3, 0),
}
for st, (n_mar, n_var, n_com, n_pub) in QUOTA.items():
    rows = by_state[st]
    for v in pick(rows, n_com, prefer=lambda r: "lm_forced" in r):
        v["committed"] = True
    for v in pick(rows, n_mar, prefer=lambda r: r.get("committed")):
        v["has_mar"] = True
    for v in pick(rows, n_var, prefer=lambda r: r.get("has_mar") or r.get("committed")):
        v["has_var"] = True
    for v in pick(rows, n_pub, prefer=lambda r: r.get("committed")):
        v["published"] = True

pub_rows = [v for v in versions if v.get("published")]
for v in pick(pub_rows, 4):
    v["published_by_set"] = True
for v in pick([v for v in pub_rows if v.get("published_by_set")], 3):
    v["approved"] = True
for v in pick([v for v in versions if v.get("has_var") and v["state"] in ("Publish", "Retired")], 2,
              prefer=lambda r: r.get("approved")):
    v["var_approver"] = True
for v in pick([v for v in versions if v.get("has_var") and v["state"] in ("Publish", "Retired", "Lock")], 3,
              prefer=lambda r: r["state"] == "Publish"):
    v["var_ref"] = True
for v in pick(pub_rows, 5) + pick(by_state["Lock"], 3, prefer=lambda r: r.get("committed")):
    v["ack"] = rng.randint(1, 5)

# ------------------------------------------------------- last modified + timestamps
non_committed = [v for v in versions if not v.get("committed") and "lm_forced" not in v]
needed_lm = 110 - sum(1 for v in versions if v.get("committed") or "lm_forced" in v)
lm_rows = pick(non_committed, needed_lm, prefer=lambda r: r["state"] != "Draft")
for v in lm_rows:
    v["has_lm"] = True
for v in versions:
    if v.get("committed") or "lm_forced" in v:
        v["has_lm"] = True
assert sum(1 for v in versions if v.get("has_lm")) == 110

no_lmb = pick([v for v in versions if not v.get("has_lm") and v["state"] == "Draft"], 1)[0]

for v in versions:
    st = v["state"]
    if v is no_lmb:
        v["last_modified_by"] = None
    elif st == "Draft":
        v["last_modified_by"] = v["created_by"] if rng.random() < 0.85 else rng.choice(REVIEWERS)
    else:
        v["last_modified_by"] = rng.choice(REVIEWERS) if rng.random() < 0.7 else v["created_by"]
    if not v.get("has_lm"):
        v["last_modified_on"] = None
        continue
    if "lm_forced" in v:
        lm = v["lm_forced"]
    else:
        lo = max(v["created_on"] + timedelta(hours=1), LM_MIN_DAY)
        hi = min(max(v["created_on"] + timedelta(days=rng.uniform(1, 20)), lo + timedelta(days=1)), LM_MAX)
        lm = rand_time(lo, hi)
    v["last_modified_on"] = lm
    # workflow events precede the last edit: commit, then approval, then publication
    if v.get("committed"):
        com = lm - timedelta(minutes=rng.randint(12, 60 * 30), seconds=rng.randint(0, 59))
        v["committed_on"] = max(com, v["created_on"] + timedelta(minutes=rng.randint(20, 90)))
        if v["committed_on"] > lm:
            v["committed_on"] = lm
    if v.get("published"):
        span = (lm - v["committed_on"]).total_seconds()
        v["published_on"] = v["committed_on"] + timedelta(seconds=int(span * rng.uniform(0.4, 0.9)))
        v["approved_on"] = v["committed_on"] + timedelta(seconds=int(span * rng.uniform(0.1, 0.35)))

# pin the span of modification dates: first on 2026-07-10, last on 2026-10-28
lm_set = [v for v in versions if v["last_modified_on"] and "lm_forced" not in v and not v.get("committed")]
first = min((v for v in lm_set if v["created_on"] <= datetime(2026, 7, 10, 8, 0)), key=lambda v: v["created_on"])
first["last_modified_on"] = datetime(2026, 7, 10, 9, 41, 7)
latest = max((v for v in lm_set if v is not first), key=lambda v: v["created_on"])
latest["last_modified_on"] = max(latest["created_on"] + timedelta(minutes=37), datetime(2026, 10, 28, 15, 2, 44))

# ------------------------------------------------------------ id assignment
doc_ids = sorted([5, 228] + rng.sample(range(6, 228), 129))
docs.sort(key=lambda d: d["created_on"])
for d, did in zip(docs, doc_ids):
    d["document_id"] = did
ver_ids = sorted([1004, 1257] + rng.sample(range(1005, 1257), 143))
versions.sort(key=lambda v: (v["created_on"], v["doc"]["created_on"], v["seq"]))
for v, vid in zip(versions, ver_ids):
    v["document_version_id"] = vid


# ------------------------------------------------- reference values with exact distinct counts
def assign_distinct(rows: list[dict], target: int, key: str, new_value):
    """Give rows values so that exactly `target` distinct values exist.

    Rows of the same document share a value first (a workflow reference carried through the chain);
    then documents of the same model share; then counts are adjusted.
    """
    groups: dict[int, list] = {}
    for r in rows:
        groups.setdefault(r["doc"]["document_id"], []).append(r)
    glist = list(groups.values())
    rng.shuffle(glist)
    values = [new_value() for _ in range(len(glist))]
    if len(glist) > target:
        # merge surplus groups into existing values (prefer same model)
        for i in range(target, len(glist)):
            model = glist[i][0]["doc"]["model_id"]
            same = [j for j in range(target) if glist[j][0]["doc"]["model_id"] == model]
            values[i] = values[same[0] if same else rng.randrange(target)]
    for g, val in zip(glist, values):
        for r in g:
            r[key] = val
    distinct = len({r[key] for r in rows})
    if distinct < target:
        # split: give later rows of multi-row groups their own value
        splittable = [r for g in glist if len(g) > 1 for r in sorted(g, key=lambda x: x["seq"])[1:]]
        rng.shuffle(splittable)
        for r in splittable[: target - distinct]:
            r[key] = new_value()
    assert len({r[key] for r in rows}) == target, (key, len({r[key] for r in rows}), target)


used_codes: set[str] = set()


def unique_digits() -> str:
    while True:
        s = str(rng.randint(500100, 509999))
        if s not in used_codes:
            used_codes.add(s)
            return s


def mar_value() -> str:
    d = unique_digits()
    return ("CM" + d) if rng.random() < 0.4 else d


def var_value() -> str:
    d = str(rng.randint(91000, 99999))
    return ("P" + d) if rng.random() < 0.3 else d


mar_rows = [v for v in versions if v.get("has_mar")]
var_rows = [v for v in versions if v.get("has_var")]
assign_distinct(mar_rows, 27, "mar_id", mar_value)
assign_distinct(var_rows, 23, "var_id", var_value)

# mar_referenceNumber: MA + digits of the MAR id for MAR rows, plus 10 extra rows with new refs
for v in mar_rows:
    v["mar_ref"] = "MA" + "".join(ch for ch in v["mar_id"] if ch.isdigit())
extra_ref_rows = pick([v for v in versions if not v.get("has_mar")], 10,
                      prefer=lambda r: r["state"] in ("Lock", "Retired"))
for v in extra_ref_rows:
    v["mar_ref"] = "MA" + unique_digits()
assert len([v for v in versions if v.get("mar_ref")]) == 42
assert len({v["mar_ref"] for v in versions if v.get("mar_ref")}) == 37

# ----------------------------------------------------------------- display names
cb_blank = set(id(v) for v in pick(versions, 48))
lmb_candidates = [v for v in versions if v is not no_lmb]
lmb_blank = set(id(v) for v in pick(lmb_candidates, 25)) | {id(no_lmb)}


# --------------------------------------------------------------------- write CSV
def fmt(t):
    return t.strftime(FMT) if t else ""


DOC_COLS = ["document_id", "document_type_code", "business_document_id", "created_by", "created_on",
            "last_modified_by", "last_modified_on", "model_id", "is_active", "description"]
VER_COLS = ["document_version_id", "document_id", "template_id", "template_version", "version_no", "state",
            "based_on_version_id", "created_by", "created_on", "approved_by", "approved_on", "last_modified_by",
            "last_modified_on", "section_id", "template_driver_mapping", "mar_id", "var_id",
            "committed_to_vesta_by", "committed_to_vesta_on", "var_approverId", "var_ReferenceId",
            "published_date", "published_by", "mar_referenceNumber", "created_by_name", "last_modified_by_name",
            "ack_template_version"]

OUT.mkdir(exist_ok=True)
with open(OUT / "vesta_Document.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(DOC_COLS)
    for d in sorted(docs, key=lambda d: d["document_id"]):
        w.writerow([d["document_id"], d["type"], 100110 if d["type"] == "MDD" else 100101, d["created_by"],
                    fmt(d["created_on"]), "", "", d["model_id"], "", d["description"]])

with open(OUT / "vesta_DocumentVersion.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(VER_COLS)
    for v in sorted(versions, key=lambda v: v["document_version_id"]):
        lmb = v["last_modified_by"]
        lmo = v["last_modified_on"]
        committed = v.get("committed")
        published = v.get("published")
        pub_by = lmb if v.get("published_by_set") else ""
        # approver is independent of the publisher and of the document's creator
        approver = ""
        if v.get("approved"):
            approver = rng.choice([r for r in REVIEWERS if r not in (lmb, v["created_by"], v["doc"]["created_by"])])
            v["approver"] = approver
        w.writerow([
            v["document_version_id"], v["doc"]["document_id"], v["template_id"], v["template_version"],
            v["version_no"], v["state"],
            v["based_on"]["document_version_id"] if v["based_on"] else "",
            v["created_by"], fmt(v["created_on"]),
            approver, fmt(v["approved_on"]) if v.get("approved") else "",
            lmb or "", fmt(lmo), "", "",
            v.get("mar_id", ""), v.get("var_id", ""),
            lmb if committed else "", fmt(v.get("committed_on")) if committed else "",
            (v.get("approver") or rng.choice(REVIEWERS)) if v.get("var_approver") else "",
            str(rng.randint(91000, 99999)) if v.get("var_ref") else "",
            fmt(v.get("published_on")) if published else "", pub_by,
            v.get("mar_ref", ""),
            "" if id(v) in cb_blank else USER_NAME[v["created_by"]],
            "" if id(v) in lmb_blank else USER_NAME[lmb],
            v.get("ack", ""),
        ])

print(f"wrote {len(docs)} documents and {len(versions)} versions to {OUT}")
