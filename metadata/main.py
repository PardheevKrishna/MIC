import re
import pandas as pd

# ——— Helpers —————————————————————————————————————————————————

def split_fields(select_list: str) -> list:
    """
    Split a SELECT list string by commas at top-level (ignoring commas inside parentheses).
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
    If segment has "FROM CONNECTION TO ... ( ... )", return the inner SQL.
    """
    m = re.search(
        r'connection\s+to\s+\w+\s*\((.*)\)\s*;',
        segment, re.IGNORECASE|re.DOTALL
    )
    return m.group(1).strip() if m else ""


# ——— Core parsing of one PROC SQL block ——————————————————————————————

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    # 1) target dataset
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select',
                     segment, re.IGNORECASE)
    if not m_ds:
        return pd.DataFrame()
    ds = m_ds.group(1)

    # 2) unwrap pass-thru if needed
    if re.search(r'select\s*\*\s*from\s*connection\s+to',
                 segment, re.IGNORECASE):
        inner = extract_pass_thru_inner(segment)
        if inner:
            segment = f"create table {ds} as {inner};"

    # 3) gather source aliases → table, fields, logic
    sources = {}

    # 3a) any "FROM X AS alias" or "JOIN X AS alias"
    for m in re.finditer(
        r'\b(?:from|join)\s+([^\s]+)\s+as\s+(\w+)\b',
        segment, re.IGNORECASE
    ):
        tbl, alias = m.group(1), m.group(2).lower()
        sources[alias] = {'table': tbl, 'fields': [], 'logic': ''}

    # 3b) LEFT JOIN (sub-query) AS alias
    for m in re.finditer(
        r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(\w+)',
        segment, re.IGNORECASE|re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(
            r'select\s+(?:distinct\s+)?(.*?)\s+from',
            subq, re.IGNORECASE|re.DOTALL
        )
        sources[alias] = {
            'table' : tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic' : subq
        }

    # 3c) INNER JOIN (sub-query) alias ON ...  (no "AS")
    for m in re.finditer(
        r'inner\s+join\s*\(\s*(select\b.*?\))\s*\)\s*(\w+)\s+on\b',
        segment, re.IGNORECASE|re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(
            r'select\s+(?:distinct\s+)?(.*?)\s+from',
            subq, re.IGNORECASE|re.DOTALL
        )
        sources[alias] = {
            'table' : tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic' : subq
        }

    # 4) pull out the SELECT list for ds
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE|re.DOTALL
    )
    selects = m_sel.group(1) if m_sel else ""
    fragments = split_fields(selects)

    # 5) parse each fragment
    records = []
    for frag in fragments:
        # flatten to one line for alias detection
        flat = re.sub(r'\s+', ' ', frag).strip()

        # detect "expr AS alias"
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$',
                        flat, re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1), m_as.group(2)
        else:
            expr = flat
            # simple alias.col
            m_col = re.match(r'^(\w+)\.(\w+)$', expr)
            fname = m_col.group(2) if m_col else expr

        # classify straight-pull if exactly alias.col and alias known
        sp = re.match(r'^(\w+)\.(\w+)$', expr)
        if sp and sp.group(1).lower() in sources:
            logic_type = 'straight pull'
        else:
            logic_type = 'derived'

        # find all aliases used in expr
        used = list(dict.fromkeys(
            re.findall(r'\b(\w+)\.', expr)
        ))
        if used:
            for a in used:
                info = sources.get(a.lower(), {})
                records.append({
                    'dataset_name': ds,
                    'field_name'  : fname,
                    'logic_type'  : logic_type,
                    'expression'  : expr,
                    'table'       : info.get('table',''),
                    'fields'      : ', '.join(info.get('fields',[])),
                    'logic'       : info.get('logic','')
                })
        else:
            # e.g. a literal, a constant, or an expression without table refs
            records.append({
                'dataset_name': ds,
                'field_name'  : fname,
                'logic_type'  : logic_type,
                'expression'  : expr,
                'table'       : '',
                'fields'      : '',
                'logic'       : ''
            })

    return pd.DataFrame(records)


# ——— Parse a full SAS file ——————————————————————————————————————

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path, 'r').read()
    # strip out /* … */ and * … ;
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    # find all PROC SQL…QUIT; segments
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# ——— Main: hard-coded paths —————————————————————————————————

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Extracted {len(df)} rows → {output_excel}")