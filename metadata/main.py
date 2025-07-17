import re
import pandas as pd

def split_fields(select_list: str) -> list:
    """
    Split a SELECT list string by commas at top-level (ignoring commas inside parentheses).
    """
    fields = []
    buff = ''
    depth = 0
    for c in select_list:
        if c == '(':
            depth += 1
            buff += c
        elif c == ')':
            depth = max(depth - 1, 0)
            buff += c
        elif c == ',' and depth == 0:
            fields.append(buff.strip())
            buff = ''
        else:
            buff += c
    if buff.strip():
        fields.append(buff.strip())
    return fields

def parse_proc_sql_segment(segment: str) -> pd.DataFrame:
    """
    Parses a single PROC SQL ... QUIT; segment and returns a DataFrame with:
      - dataset_name
      - field_name   (underscores preserved)
      - logic_type   ('straight pull' or 'derived')
      - expression   (raw SQL expression)
      - table        (one source table per row)
      - fields       (one source field list per row)
      - logic        (full subquery text or empty for main tables)
    Explodes multi-source fields into separate rows.
    """
    records = []
    ds_match = re.search(r'create table\s+(\w+)\s+as select', segment, re.IGNORECASE)
    if not ds_match:
        return pd.DataFrame(records)
    ds = ds_match.group(1)

    # Extract source definitions
    sources_info = {}
    # Main FROM (t1)
    m1 = re.search(r'from\s+([^\s]+)\s+as\s+(t1)', segment, re.IGNORECASE)
    if m1:
        sources_info['t1'] = {'table': m1.group(1), 'fields': [], 'logic': ''}
    # Direct JOIN (t2)
    m2 = re.search(r'join\s+([^\s]+)\s+as\s+(t2)', segment, re.IGNORECASE)
    if m2:
        sources_info['t2'] = {'table': m2.group(1), 'fields': [], 'logic': ''}
    # Subqueries in LEFT JOIN
    subq_pattern = re.compile(
        r'left join\s*\(\s*(select\b.*?\))\s+as\s+(t\d+)', re.IGNORECASE | re.DOTALL
    )
    for m in subq_pattern.finditer(segment):
        subquery = m.group(1).strip()
        alias = m.group(2).lower()
        table_match = re.search(r'from\s+([^\s]+)', subquery, re.IGNORECASE)
        fields_match = re.search(r'select\s+(?:distinct\s+)?(.*?)\s+from', subquery, re.IGNORECASE | re.DOTALL)
        fields_list = [f.strip() for f in fields_match.group(1).split(',')] if fields_match else []
        sources_info[alias] = {
            'table': table_match.group(1) if table_match else '',
            'fields': fields_list,
            'logic': subquery
        }

    # Extract SELECT list
    sel_match = re.search(r'create table\s+' + re.escape(ds) + r'\s+as select\s+(.*?)\s+from\b',
                          segment, re.IGNORECASE | re.DOTALL)
    select_list = sel_match.group(1) if sel_match else ''
    fragments = split_fields(select_list)

    for frag in fragments:
        frag = frag.strip()
        alias_match = re.match(r'(.+?)\s+AS\s+([A-Za-z0-9_]+)$', frag, re.IGNORECASE)
        if alias_match:
            expr = alias_match.group(1).strip()
            field_name = alias_match.group(2).strip()
        else:
            expr = frag
            col_match = re.match(r'\b(t\d+)\.(\w+)$', expr, re.IGNORECASE)
            field_name = col_match.group(2) if col_match else expr

        logic_type = ('straight pull'
                      if re.match(r'\b(t\d+)\.\w+', expr, re.IGNORECASE) and not alias_match
                      else 'derived')

        aliases_used = list(dict.fromkeys(re.findall(r'\b(t\d+)\.', expr, re.IGNORECASE)))
        if aliases_used:
            for a in aliases_used:
                key = a.lower()
                if key in sources_info:
                    info = sources_info[key]
                    records.append({
                        'dataset_name': ds,
                        'field_name': field_name,
                        'logic_type': logic_type,
                        'expression': expr,
                        'table': info['table'],
                        'fields': ', '.join(info['fields']),
                        'logic': info['logic']
                    })
        else:
            records.append({
                'dataset_name': ds,
                'field_name': field_name,
                'logic_type': logic_type,
                'expression': expr,
                'table': '',
                'fields': '',
                'logic': ''
            })

    return pd.DataFrame(records)

def parse_sas_file(input_path: str) -> pd.DataFrame:
    """
    Reads a .sas file, extracts all PROC SQL segments, parses each,
    and concatenates the resulting metadata.
    """
    with open(input_path, 'r') as f:
        content = f.read()
    segments = re.findall(r'proc sql\b.*?quit\s*;', content, re.IGNORECASE | re.DOTALL)
    dfs = [parse_proc_sql_segment(seg) for seg in segments]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

if __name__ == '__main__':
    # Hard-coded file paths
    input_sas = './code.sas'                
    output_excel = './metadata_mapping.xlsx'     

    # Parse and save
    df_metadata = parse_sas_file(input_sas)
    df_metadata.to_excel(output_excel, index=False)
    print(f"Metadata mapping saved to {output_excel}")
