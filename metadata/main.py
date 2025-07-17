import re
import pandas as pd

# ——— Helpers —————————————————————————————————————————————————

def split_fields(select_list: str) -> list:
    """
    Split a SELECT list string by commas at top‐level (ignore commas inside parentheses).
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
    if buf.strip(): 
        out.append(buf.strip())
    return out

def extract_pass_thru_inner(segment: str) -> str:
    """
    If segment has "FROM CONNECTION TO ...(...)", return the inner SQL.
    """
    m = re.search(r'connection\s+to\s+\w+\s*\((.*)\)\s*;', 
                  segment, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""


# ——— Core parsing of one PROC SQL block ——————————————————————————————

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    # 1) target dataset
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select',
                     segment, re.IGNORECASE)
    if not m_ds:
        return pd.DataFrame()
    ds = m_ds.group(1)

    # 2) unwrap pass‐thru if needed
    if re.search(r'select\s*\*\s*from\s*connection\s+to', 
                 segment, re.IGNORECASE):
        inner = extract_pass_thru_inner(segment)
        if inner:
            segment = f"create table {ds} as {inner};"

    # 3) gather source aliases → tables and subqueries
    sources = {}
    # any FROM or JOIN alias
    for m in re.finditer(r'\b(?:from|join)\s+([^\s]+)\s+as\s+(\w+)\b',
                         segment, re.IGNORECASE):
        sources[m.group(2).lower()] = {'table': m.group(1),
                                       'fields': [], 'logic': ''}
    # subquery LEFT JOIN
    for m in re.finditer(
        r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(\w+)',
        segment, re.IGNORECASE | re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from',
                        subq, re.IGNORECASE | re.DOTALL)
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic': subq
        }

    # 4) pull out the SELECT list
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE | re.DOTALL
    )
    selects = m_sel.group(1) if m_sel else ''
    fragments = split_fields(selects)

    # 5) parse each fragment
    records = []
    for frag in fragments:
        # flatten all whitespace to single spaces for easy alias matching
        flat = re.sub(r'\s+', ' ', frag).strip()

        # detect "expr AS alias"
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$',
                        flat, re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1).strip(), m_as.group(2).strip()
        else:
            expr = flat
            # wildcard expansion (handled elsewhere)
            if expr.endswith('.*') or expr == '*':
                alias = expr.split('.')[0] if '.' in expr else ''
                tbl   = sources.get(alias,{}).get('table','')
                for col in sources.get(alias,{}).get('fields',[]):
                    records.append({
                        'dataset_name': ds,
                        'field_name'  : col,
                        'logic_type'  : 'straight pull',
                        'expression'  : f"{alias}.{col}",
                        'table'       : tbl,
                        'fields'      : '',
                        'logic'       : ''
                    })
                continue
            # simple alias.col
            m_col = re.match(r'^(\w+)\.(\w+)$', expr)
            if m_col:
                fname = m_col.group(2)
            else:
                fname = expr

        # classify straight vs derived
        sp = re.match(r'^(\w+)\.(\w+)$', expr)
        if sp and sp.group(1).lower() in sources:
            logic_type = 'straight pull'
        else:
            logic_type = 'derived'

        # which aliases drive this expression?
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


# ——— Parse the entire .sas file —————————————————————————————————

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path, 'r').read()
    # strip block comments and star-comments
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    # grab every PROC SQL…QUIT; segment
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# ——— Main (hard-coded paths) —————————————————————————————————

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Extracted {len(df)} rows → {output_excel}")
