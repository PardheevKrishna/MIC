import re
import pandas as pd

# ——— Helpers —————————————————————————————————————————————————

def split_fields(select_list: str) -> list:
    """Split by top-level commas, ignoring any inside parentheses."""
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
    If segment uses "SELECT * FROM CONNECTION TO ...(...)",
    return the inner SQL.
    """
    m = re.search(r'connection\s+to\s+\w+\s*\((.*)\)\s*;', 
                  segment, re.IGNORECASE|re.DOTALL)
    return m.group(1).strip() if m else ""


# ——— Build per-table column mapping ——————————————————————————————

def build_table_columns(segment: str,
                        existing_map: dict[str,list[str]]) -> list[str]:
    """
    Given one PROC SQL block, return the explicit field list for
    its target table.  If it’s a pass-through, we parse the inner SELECT;
    if it lists wildcards, we expand them from `existing_map`.
    """
    # 1) find target dataset
    m_ds = re.search(r'create\s+table\s+(\w+)\s+as\s+select',
                     segment, re.IGNORECASE)
    if not m_ds:
        return []
    ds = m_ds.group(1)

    # 2) if it’s a pass-thru wildcard, unwrap it
    if re.search(r'select\s*\*\s*from\s*connection\s+to', 
                 segment, re.IGNORECASE):
        inner = extract_pass_thru_inner(segment)
        if inner:
            # rebuild as a plain SELECT for parsing:
            segment = f"create table {ds} as {inner};"

    # 3) get the SELECT list text
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE|re.DOTALL
    )
    select_list = m_sel.group(1) if m_sel else ""
    frags = split_fields(select_list)

    cols = []
    for frag in frags:
        frag = frag.strip()
        # explicit alias "X as Y"
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$', frag,
                        re.IGNORECASE)
        if m_as:
            expr, alias = m_as.group(1).strip(), m_as.group(2).strip()
            cols.append(alias)
            continue

        # wildcard "alias.*" or just "*"
        if frag.endswith('.*') or frag == '*':
            alias = frag.split('.')[0] if '.' in frag else ''
            if alias in existing_map:
                cols.extend(existing_map[alias])
            continue

        # simple column reference tX.colname
        m_col = re.match(r'\b\w+\.(\w+)$', frag)
        if m_col:
            cols.append(m_col.group(1))
            continue

        # otherwise it’s some expression; we can’t name it
        # so skip it for building the table’s column map
        # (we’ll still capture it later in the metadata)
    return cols


# ——— Parse metadata rows for one PROC SQL block ————————————————————

def parse_proc_sql_segment(segment: str,
                           table_map: dict[str,list[str]]) -> pd.DataFrame:
    """
    Returns the 7-column metadata for one PROC SQL…QUIT; block,
    expanding any tX.* using `table_map`.
    """
    # 1) dataset name
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

    # 3) discover all source aliases → table names
    sources = {}
    # any "from X as alias" or "join X as alias"
    for m in re.finditer(r'\b(?:from|join)\s+([^\s]+)\s+as\s+(\w+)\b',
                         segment, re.IGNORECASE):
        tbl, alias = m.group(1), m.group(2)
        sources[alias.lower()] = {'table':tbl, 'logic': '', 'fields': []}

    # subqueries: LEFT JOIN ( SELECT … ) AS alias
    for m in re.finditer(r'left\s+join\s*\(\s*(select\b.*?\))\s+as\s+(\w+)',
                         segment, re.IGNORECASE|re.DOTALL):
        subq, alias = m.group(1).strip(), m.group(2).lower()
        # overwrite table with the real table inside subq
        tblm = re.search(r'from\s+([^\s]+)', subq,
                         re.IGNORECASE)
        flm = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from',
                        subq, re.IGNORECASE|re.DOTALL)
        sources[alias] = {
            'table': tblm.group(1) if tblm else '',
            'logic': subq,
            'fields': split_fields(flm.group(1)) if flm else []
        }

    # 4) get SELECT list
    m_sel = re.search(
        rf'create\s+table\s+{ds}\s+as\s+select\s+(.*?)\s+from\b',
        segment, re.IGNORECASE|re.DOTALL
    )
    bits = split_fields(m_sel.group(1)) if m_sel else []

    rows = []
    for frag in bits:
        frag = frag.strip()
        # alias?
        m_as = re.match(r'(.+?)\s+as\s+([A-Za-z0-9_]+)$', frag,
                        re.IGNORECASE)
        if m_as:
            expr, fname = m_as.group(1), m_as.group(2)
        else:
            expr = frag
            # wildcard?
            if expr.endswith('.*') or expr == '*':
                alias = expr.split('.')[0] if '.' in expr else ''
                for col in table_map.get(sources.get(alias,{}).get('table',''), []):
                    rows.append({
                        'dataset_name': ds,
                        'field_name'  : col,
                        'logic_type'  : 'straight pull',
                        'expression'  : f"{alias}.{col}" if alias else col,
                        'table'       : sources.get(alias,{}).get('table',''),
                        'fields'      : '',
                        'logic'       : ''
                    })
                continue
            # simple tX.col
            m_col = re.match(r'\b(t\d+)\.(\w+)$', expr, re.IGNORECASE)
            if m_col:
                fname = m_col.group(2)
            else:
                fname = expr

        # logic type
        logic_type = ('straight pull'
                      if re.match(r'\b(t\d+)\.\w+', expr, re.IGNORECASE)
                         and not m_as
                      else 'derived')

        # find all aliases used
        aliases = list(dict.fromkeys(
            re.findall(r'\b(t\d+)\.', expr, re.IGNORECASE)
        ))
        if aliases:
            for a in aliases:
                info = sources.get(a.lower(), {})
                rows.append({
                    'dataset_name': ds,
                    'field_name'  : fname,
                    'logic_type'  : logic_type,
                    'expression'  : expr,
                    'table'       : info.get('table',''),
                    'fields'      : ', '.join(info.get('fields',[])),
                    'logic'       : info.get('logic','')
                })
        else:
            rows.append({
                'dataset_name': ds,
                'field_name'  : fname,
                'logic_type'  : logic_type,
                'expression'  : expr,
                'table'       : '',
                'fields'      : '',
                'logic'       : ''
            })

    return pd.DataFrame(rows)


# ——— Walk the entire SAS file —————————————————————————————————

def parse_sas_file(path: str) -> pd.DataFrame:
    txt = open(path).read()
    # strip /* ... */ and * ... ;
    txt = re.sub(r'/\*.*?\*/', '', txt, flags=re.DOTALL)
    txt = re.sub(r'^\s*\*.*?;', '', txt, flags=re.MULTILINE)

    # find all PROC SQL…QUIT; segments in order
    segments = re.findall(r'(?is)proc\s+sql\b.*?\bquit\b\s*;', txt)

    table_map = {}
    all_rows = []

    # first pass: build table_map in file order
    for seg in segments:
        # discover dataset name
        m_ds = re.search(r'create\s+table\s+(\w+)', seg,
                         re.IGNORECASE)
        if not m_ds:
            continue
        ds = m_ds.group(1)
        cols = build_table_columns(seg, table_map)
        table_map[ds] = cols

        # parse metadata rows for this segment
        df_seg = parse_proc_sql_segment(seg, table_map)
        if not df_seg.empty:
            all_rows.append(df_seg)

    if all_rows:
        return pd.concat(all_rows, ignore_index=True)
    else:
        return pd.DataFrame()


# ——— Main (hard-coded paths) ————————————————————————————————

if __name__ == '__main__':
    input_sas    = '/mnt/data/code.sas'
    output_excel = '/mnt/data/metadata_mapping.xlsx'

    df = parse_sas_file(input_sas)
    df.to_excel(output_excel, index=False)
    print(f"Saved {len(df)} rows → {output_excel}")
