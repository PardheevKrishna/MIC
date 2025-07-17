import re
import pandas as pd

# —— Helpers —— #

def split_fields(select_list: str) -> list:
    """
    Split a SELECT list string by commas at top‐level (ignoring commas inside parentheses).
    """
    out, buf, depth = [], '', 0
    for c in select_list:
        if c == '(':
            depth += 1; buf += c
        elif c == ')':
            depth = max(depth-1, 0); buf += c
        elif c == ',' and depth == 0:
            out.append(buf.strip()); buf = ''
        else:
            buf += c
    if buf.strip(): out.append(buf.strip())
    return out

def extract_pass_thru_inner(segment: str) -> str:
    """
    If segment has "FROM CONNECTION TO ...(...)", return the inner SQL.
    """
    m = re.search(
        r'connection\s+to\s+\w+\s*\((.*)\)\s*;', 
        segment, re.IGNORECASE|re.DOTALL
    )
    return m.group(1).strip() if m else ""

# —— Core parsing for one PROC SQL block —— #

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    """
    Parse one PROC SQL...QUIT segment to extract:
      - dataset_name, field_name, logic_type, expression
    """
    # 1) target dataset
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select',
                     segment, re.IGNORECASE)
    if not m_ds:
        return pd.DataFrame()
    ds = m_ds.group(1)

    # 2) handle pass-through SELECT * FROM CONNECTION to ...
    if re.search(r'select\s*\*\s*from\s*connection\s+to',
                 segment, re.IGNORECASE):
        inner = extract_pass_thru_inner(segment)
        if inner:
            segment = f"create table {ds} as {inner};"

    # 3) gather source aliases for straight-pull detection
    sources = set()
    for m in re.finditer(r'\b(?:from|join)\s+[^\s]+\s+as\s+(\w+)\b',
                         segment, re.IGNORECASE):
        sources.add(m.group(1).lower())
    # also catch subqueries without 'AS': ") t2 on"
    for m in re.finditer(r'\)\s*(\w+)\s+on\b', segment, re.IGNORECASE):
        sources.add(m.group(1).lower())

    # 4) extract SELECT list
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE|re.DOTALL
    )
    selects = m_sel.group(1) if m_sel else ""
    fragments = split_fields(selects)

    # 5) build rows
    records = []
    for frag in fragments:
        flat = re.sub(r'\s+', ' ', frag).strip()
        # detect "expr AS alias"
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$',
                        flat, re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1), m_as.group(2)
        else:
            expr = flat
            # simple alias.col or bare column
            m_col = re.match(r'^(\w+\.\w+|\w+)$', expr)
            fname = m_col.group(1) if m_col else expr

        # classify straight vs derived
        sp = re.match(r'^(\w+)\.(\w+)$', expr)
        logic_type = 'straight pull' if sp and sp.group(1).lower() in sources else 'derived'

        records.append({
            'dataset_name': ds,
            'field_name'  : fname,
            'logic_type'  : logic_type,
            'expression'  : expr
        })

    return pd.DataFrame(records)

# —— Parse full SAS file and explode F column —— #

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path, 'r').read()
    # strip comments
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    # find each PROC SQL...QUIT
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    # extract all tX.col occurrences from expression into 'f'
    df['f'] = df['expression'].apply(
        lambda expr: [m.group(2) for m in re.finditer(r'\b(t\d+)\.(\w+)\b', expr)]
    )
    # explode so each row has exactly one f
    df = df.explode('f').reset_index(drop=True)

    return df[['dataset_name','field_name','logic_type','expression','f']]

# ——— Main (hard-coded) —————————————————————————————————————————

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Saved {len(df)} rows → {output_excel}")
