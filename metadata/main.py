import re
import pandas as pd

# ——— Helpers —————————————————————————————————————————————————

def split_fields(select_list: str) -> list:
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
    m = re.search(
        r'connection\s+to\s+\w+\s*\((.*)\)\s*;',
        segment, re.IGNORECASE | re.DOTALL
    )
    return m.group(1).strip() if m else ""

# ——— Core parsing of one PROC SQL block ——————————————————————————————

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select',
                     segment, re.IGNORECASE)
    if not m_ds:
        return pd.DataFrame()
    ds = m_ds.group(1)

    if re.search(r'select\s*\*\s*from\s*connection\s+to',
                 segment, re.IGNORECASE):
        inner = extract_pass_thru_inner(segment)
        if inner:
            segment = f"create table {ds} as {inner};"

    sources = {}
    for m in re.finditer(r'\b(?:from|join)\s+([^\s]+)\s+as\s+(\w+)\b',
                         segment, re.IGNORECASE):
        tbl, alias = m.group(1), m.group(2).lower()
        sources[alias] = {'table': tbl, 'fields': [], 'logic': ''}

    for m in re.finditer(
        r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(\w+)', 
        segment, re.IGNORECASE | re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(
            r'select\s+(?:distinct\s+)?(.*?)\s+from',
            subq, re.IGNORECASE | re.DOTALL
        )
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic': subq
        }

    for m in re.finditer(
        r'inner\s+join\s*\(\s*(select\b.*?\))\s*\)\s*(\w+)\s+on\b',
        segment, re.IGNORECASE | re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(
            r'select\s+(?:distinct\s+)?(.*?)\s+from',
            subq, re.IGNORECASE | re.DOTALL
        )
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic': subq
        }

    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE | re.DOTALL
    )
    selects = m_sel.group(1) if m_sel else ""
    fragments = split_fields(selects)

    records = []
    for frag in fragments:
        flat = re.sub(r'\s+', ' ', frag).strip()
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$',
                        flat, re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1), m_as.group(2)
        else:
            expr = flat
            m_col = re.match(r'^(\w+)\.(\w+)$', expr)
            fname = m_col.group(2) if m_col else expr

        sp = re.match(r'^(\w+)\.(\w+)$', expr)
        logic_type = 'straight pull' if sp and sp.group(1).lower() in sources else 'derived'

        used = list(dict.fromkeys(re.findall(r'\b(\w+)\.', expr)))
        records.append({
            'dataset_name': ds,
            'field_name'  : fname,
            'logic_type'  : logic_type,
            'expression'  : expr,
            'table'       : ', '.join(
                                sources[a]['table'] for a in used if a in sources
                            ),
            'logic'       : ', '.join(
                                sources[a]['logic'] for a in used if a in sources
                            )
        })

    return pd.DataFrame(records)

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path).read()
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    # extract fields from expression and explode
    df['fields'] = df['expression'].apply(lambda exp:
        re.findall(r'\b\w+\.(\w+)\b', exp))
    df = df.explode('fields').reset_index(drop=True)

    # drop any exact duplicates
    df = df.drop_duplicates()

    # reorder columns so 'fields' is before 'logic'
    df = df[
        ['dataset_name','field_name','logic_type',
         'expression','table','fields','logic']
    ]

    df.to_excel(output_excel, index=False)
    print(f"Extracted {len(df)} rows → {output_excel}")
