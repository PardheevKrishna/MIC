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
    if buf.strip():
        out.append(buf.strip())
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
      - dataset_name, field_name, logic_type, expression, table, fields (source), logic
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

    # 3) gather source aliases → table, fields, logic
    sources = {}
    # a) FROM/JOIN with AS
    for m in re.finditer(r'\b(?:from|join)\s+([^\s]+)\s+as\s+(\w+)\b',
                         segment, re.IGNORECASE):
        alias = m.group(2).lower()
        sources[alias] = { 'table': m.group(1),
                           'fields': [], 'logic': '' }
    # b) LEFT JOIN subqueries with AS
    for m in re.finditer(
        r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(\w+)',
        segment, re.IGNORECASE|re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from',
                        subq, re.IGNORECASE|re.DOTALL)
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic': subq
        }
    # c) INNER JOIN subqueries without AS
    for m in re.finditer(
        r'inner\s+join\s*\(\s*(select\b.*?\))\s*\)\s*(\w+)\s+on\b',
        segment, re.IGNORECASE|re.DOTALL
    ):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        tblm = re.search(r'from\s+([^\s]+)', subq, re.IGNORECASE)
        flm = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from',
                        subq, re.IGNORECASE|re.DOTALL)
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'fields': split_fields(flm.group(1)) if flm else [],
            'logic': subq
        }

    # 4) extract SELECT list
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE|re.DOTALL
    )
    selects = m_sel.group(1) if m_sel else ""
    fragments = split_fields(selects)

    # 5) parse fragments
    records = []
    for frag in fragments:
        flat = re.sub(r'\s+', ' ', frag).strip()

        # detect "expr AS alias"
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$', flat, re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1), m_as.group(2)
        else:
            expr = flat
            m_col = re.match(r'^(\w+\.\w+|\w+)$', expr)
            fname = m_col.group(1) if m_col else expr

        # straight vs derived
        sp = re.match(r'^(\w+)\.(\w+)$', expr)
        logic_type = 'straight pull' if sp and sp.group(1).lower() in sources else 'derived'

        # determine table & logic by first alias in expr
        alias_used = re.search(r'\b(\w+)\.', expr)
        if alias_used:
            info = sources.get(alias_used.group(1).lower(), {})
            table = info.get('table','')
            logic = info.get('logic','')
            src_fields = info.get('fields',[])
        else:
            table = ''
            logic = ''
            src_fields = []

        records.append({
            'dataset_name': ds,
            'field_name'  : fname,
            'logic_type'  : logic_type,
            'expression'  : expr,
            'table'       : table,
            'fields_src'  : ', '.join(src_fields),
            'logic'       : logic
        })

    return pd.DataFrame(records)

# —— Full file parse & explode fields —— #

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path, 'r').read()
    # strip comments
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    # derive fields from expression into 'fields'
    df['fields'] = df['expression'].apply(
        lambda expr: re.findall(r'\b(t\d+)\.(\w+)\b', expr)
    )
    df['fields'] = df['fields'].apply(lambda lst: [col for _,col in lst])

    # explode so each row has one field
    df = df.explode('fields').reset_index(drop=True)
    df['fields'] = df['fields'].fillna('')

    # select final six columns
    return df[[
        'dataset_name',
        'field_name',
        'logic_type',
        'expression',
        'table',
        'fields',
        'logic'
    ]]

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Saved {len(df)} rows → {output_excel}")
