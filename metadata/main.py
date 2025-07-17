import re
import pandas as pd

# —— Helpers —— #

def split_fields(select_list: str) -> list:
    """
    Split a SELECT list string by commas at top‐level (ignoring commas in parentheses).
    """
    fields, buf, depth = [], '', 0
    for c in select_list:
        if c == '(':
            depth += 1; buf += c
        elif c == ')':
            depth = max(depth-1, 0); buf += c
        elif c == ',' and depth == 0:
            fields.append(buf.strip()); buf = ''
        else:
            buf += c
    if buf.strip():
        fields.append(buf.strip())
    return fields

def extract_inner_pass_thru(segment: str) -> str:
    """
    If we see "SELECT * FROM CONNECTION TO eiw(...)", grab the inner SQL.
    """
    m = re.search(r"connection\s+to\s+eiw\s*\((.*)\)\s*;", segment, re.IGNORECASE|re.DOTALL)
    return m.group(1).strip() if m else ""


# —— Core Parsing —— #

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    """
    Given one PROC SQL ... QUIT; block, returns a DataFrame of the 7 metadata columns,
    exploding any multi‐source pulls into one row per source.
    """
    # 1) Find target dataset
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select', segment, re.IGNORECASE)
    if not m_ds:
        return pd.DataFrame()
    ds = m_ds.group(1)

    # 2) Handle pass-thru SELECT * FROM CONNECTION TO eiw(...)
    if re.search(r'select\s*\*\s*from\s*connection\s+to\s+eiw', segment, re.IGNORECASE):
        inner = extract_inner_pass_thru(segment)
        if inner:
            # wrap it into a fake CREATE so we can re‐parse the SELECT list
            segment = f"create table {ds} as {inner};"

    # 3) Pull out source definitions
    sources = {}
    # main FROM t1
    m1 = re.search(r'from\s+([^\s]+)\s+as\s+(t1)\b', segment, re.IGNORECASE)
    if m1:
        sources['t1'] = dict(table=m1.group(1), fields=[], logic="")
    # join t2
    m2 = re.search(r'join\s+([^\s]+)\s+as\s+(t2)\b', segment, re.IGNORECASE)
    if m2:
        sources['t2'] = dict(table=m2.group(1), fields=[], logic="")
    # subqueries in LEFT JOIN
    for m in re.finditer(r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(t\d+)', 
                        segment, re.IGNORECASE|re.DOTALL):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tbl = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        fl = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from', subq, re.IGNORECASE|re.DOTALL)
        sources[alias] = dict(
            table = tbl.group(1) if tbl else "",
            fields = [f.strip() for f in fl.group(1).split(',')] if fl else [],
            logic = subq
        )

    # 4) Extract the SELECT list
    sel = re.search(rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b', 
                    segment, re.IGNORECASE|re.DOTALL)
    bits = split_fields(sel.group(1)) if sel else []

    rows = []
    for frag in bits:
        frag = frag.strip()
        # alias?
        am = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$', frag, re.IGNORECASE)
        if am:
            expr, fname = am.group(1).strip(), am.group(2).strip()
        else:
            expr = frag
            cm = re.match(r'\b(t\d+)\.(\w+)$', expr, re.IGNORECASE)
            fname = cm.group(2) if cm else expr

        ltype = ('straight pull' 
                 if re.match(r'\b(t\d+)\.\w+', expr, re.IGNORECASE) and not am 
                 else 'derived')

        aliases = list(dict.fromkeys(re.findall(r'\b(t\d+)\.', expr, re.IGNORECASE)))
        if aliases:
            for a in aliases:
                info = sources.get(a.lower())
                if not info: 
                    continue
                rows.append({
                    'dataset_name': ds,
                    'field_name'  : fname,
                    'logic_type'  : ltype,
                    'expression'  : expr,
                    'table'       : info['table'],
                    'fields'      : ', '.join(info['fields']),
                    'logic'       : info['logic']
                })
        else:
            rows.append({
                'dataset_name': ds,
                'field_name'  : fname,
                'logic_type'  : ltype,
                'expression'  : expr,
                'table'       : "",
                'fields'      : "",
                'logic'       : ""
            })

    return pd.DataFrame(rows)


def parse_sas_file(path: str) -> pd.DataFrame:
    """
    Reads a .sas file, strips out comments, finds every PROC SQL…QUIT; block,
    parses each, and concatenates into one DataFrame.
    """
    txt = open(path, 'r').read()
    # remove /* … */ blocks
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    # remove * … ; line comments
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    # find every PROC SQL…QUIT;
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# —— Main: hard-coded paths —— #

if __name__ == '__main__':
    input_sas    = './code.sas'               # ← put your .sas here
    output_excel = './metadata_mapping.xlsx'  # ← Excel out here

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Extracted {len(df)} rows — saved to {output_excel}")
