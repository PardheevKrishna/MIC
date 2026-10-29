"""Build the VestaDashboard Power BI project (PBIP: TMDL semantic model + PBIR report).

Output: <root>/VestaDashboard.pbip, VestaDashboard.SemanticModel/, VestaDashboard.Report/
Open VestaDashboard.pbip in Power BI Desktop. The CSV folder is the `DataFolder` parameter
(Transform data > Edit parameters) and defaults to <root>/data/.
"""
from __future__ import annotations

import json
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NAME = "VestaDashboard"
SM = ROOT / f"{NAME}.SemanticModel"
RP = ROOT / f"{NAME}.Report"
BASE_THEME = "Fluent2-CY26SU10"
BASE_THEME_FILE = Path(__file__).resolve().parent / "assets" / f"{BASE_THEME}.json"
CUSTOM_THEME = "WellsFargoVesta.json"
VERSION_AT_IMPORT = {"visual": "2.14.0", "report": "3.4.0", "page": "2.3.1"}

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
VISUAL_SCHEMA = f"{SCHEMA}/visualContainer/2.4.0/schema.json"
PAGE_SCHEMA = f"{SCHEMA}/page/2.1.0/schema.json"
PAGES_SCHEMA = f"{SCHEMA}/pagesMetadata/1.1.0/schema.json"
REPORT_SCHEMA = f"{SCHEMA}/report/3.3.0/schema.json"
VERSION_SCHEMA = f"{SCHEMA}/versionMetadata/1.0.0/schema.json"
PLATFORM_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json"

# Wells Fargo-inspired palette
WF_RED = "#D71E28"
WF_GOLD = "#FFCD41"
WF_CHARCOAL = "#3B3331"
WF_BURGUNDY = "#8C1D18"
WF_ORANGE = "#F28C28"
WF_TAUPE = "#B6ADA5"
WF_CREAM = "#F7F3EF"
WF_LINE = "#E3DCD5"
STATE_COLORS = {"Draft": WF_GOLD, "Lock": WF_CHARCOAL, "Commit to Vesta": WF_ORANGE,
                "Publish": WF_RED, "Retired": WF_TAUPE}
TYPE_COLORS = {"MDD": WF_RED, "MVD": WF_GOLD}

NS = uuid.UUID("6d1f9a1e-6c4b-4f7e-9d0c-5a1b2c3d4e5f")


def guid(key: str) -> str:
    return str(uuid.uuid5(NS, key))


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path: Path, obj) -> None:
    write(path, json.dumps(obj, indent=2) + "\n")


# ============================================================================ semantic model
DOC_COLS = [  # name, tmdl type, M type, format, summarize
    ("document_id", "int64", "Int64.Type", "0", "none"),
    ("document_type_code", "string", "type text", None, "none"),
    ("business_document_id", "int64", "Int64.Type", "0", "none"),
    ("created_by", "string", "type text", None, "none"),
    ("created_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("last_modified_by", "string", "type text", None, "none"),
    ("last_modified_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("model_id", "int64", "Int64.Type", "0", "none"),
    ("is_active", "boolean", "type logical", '"""TRUE"";""TRUE"";""FALSE"""', "none"),
    ("description", "string", "type text", None, "none"),
]
DOC_DERIVED = [("Model Type", "string"), ("Risk Tier", "string"), ("Usage Category", "string")]

VER_COLS = [
    ("document_version_id", "int64", "Int64.Type", "0", "none"),
    ("document_id", "int64", "Int64.Type", "0", "none"),
    ("template_id", "int64", "Int64.Type", "0", "none"),
    ("template_version", "int64", "Int64.Type", "0", "none"),
    ("version_no", "string", "type text", None, "none"),
    ("state", "string", "type text", None, "none"),
    ("based_on_version_id", "int64", "Int64.Type", "0", "none"),
    ("created_by", "string", "type text", None, "none"),
    ("created_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("approved_by", "string", "type text", None, "none"),
    ("approved_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("last_modified_by", "string", "type text", None, "none"),
    ("last_modified_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("section_id", "string", "type text", None, "none"),
    ("template_driver_mapping", "string", "type text", None, "none"),
    ("mar_id", "string", "type text", None, "none"),
    ("var_id", "string", "type text", None, "none"),
    ("committed_to_vesta_by", "string", "type text", None, "none"),
    ("committed_to_vesta_on", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("var_approverId", "string", "type text", None, "none"),
    ("var_ReferenceId", "string", "type text", None, "none"),
    ("published_date", "dateTime", "type datetime", "yyyy-mm-dd hh:nn", "none"),
    ("published_by", "string", "type text", None, "none"),
    ("mar_referenceNumber", "string", "type text", None, "none"),
    ("created_by_name", "string", "type text", None, "none"),
    ("last_modified_by_name", "string", "type text", None, "none"),
    ("ack_template_version", "int64", "Int64.Type", "0", "sum"),
]

T = "\t"


def q(name: str) -> str:
    """Quote a TMDL object name when needed."""
    return name if name.replace("_", "").isalnum() else "'" + name.replace("'", "''") + "'"


def tmdl_column(table: str, name: str, dtype: str, fmt: str | None = None, summarize: str = "none",
                source: str | None = None, expression: str | None = None, hidden: bool = False,
                sort_by: str | None = None, calc_table: bool = False) -> str:
    head = f"{T}column {q(name)}" + (f" = {expression}" if expression else "")
    lines = [head, f"{T*2}dataType: {dtype}"]
    if fmt:
        lines.append(f"{T*2}formatString: {fmt}")
    if hidden:
        lines.append(f"{T*2}isHidden")
    if calc_table:
        lines.append(f"{T*2}isNameInferred")
    lines.append(f"{T*2}lineageTag: {guid(table + '.' + name)}")
    lines.append(f"{T*2}summarizeBy: {summarize}")
    if source:
        lines.append(f"{T*2}sourceColumn: {source}")
    if sort_by:
        lines.append(f"{T*2}sortByColumn: {q(sort_by)}")
    lines.append("")
    lines.append(f"{T*2}annotation SummarizationSetBy = User")
    lines.append("")
    return "\n".join(lines)


def tmdl_measure(table: str, name: str, expr: str, fmt: str | None = None, folder: str | None = None) -> str:
    lines = [f"{T}measure {q(name)} = {expr}"]
    if fmt:
        lines.append(f"{T*2}formatString: {fmt}")
    if folder:
        lines.append(f"{T*2}displayFolder: {folder}")
    lines.append(f"{T*2}lineageTag: {guid(table + '.m.' + name)}")
    lines.append("")
    return "\n".join(lines)


def m_block(m: str) -> str:
    return "\n".join(f"{T*4}{ln}" if ln.strip() else "" for ln in m.strip("\n").splitlines())


def csv_partition(table: str, file: str, cols, extra_steps: str = "", final: str = "Typed") -> str:
    types = ", ".join(f'{{"{c[0]}", {c[2]}}}' for c in cols)
    m = f"""
let
    Source = Csv.Document(File.Contents(DataFolder & "{file}"), [Delimiter = ",", Columns = {len(cols)}, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    BlanksToNull = Table.ReplaceValue(Promoted, "", null, Replacer.ReplaceValue, Table.ColumnNames(Promoted)),
    Typed = Table.TransformColumnTypes(BlanksToNull, {{{types}}}, "en-US"){extra_steps}
in
    {final}
"""
    return f"{T}partition {q(table)} = m\n{T*2}mode: import\n{T*2}source =\n{m_block(m)}\n"


def build_semantic_model() -> None:
    if SM.exists():
        shutil.rmtree(SM)
    d = SM / "definition"
    write_json(SM / ".platform", {
        "$schema": PLATFORM_SCHEMA,
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid("semanticmodel")},
    })
    write_json(SM / "definition.pbism", {"version": "4.2", "settings": {}})
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")

    tables = ["vesta_Document", "vesta_DocumentVersion", "Models", "Calendar", "Workflow Field", "Key Measures"]
    model = [
        "model Model",
        "\tculture: en-US",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tsourceQueryCulture: en-US",
        "\tvalueFilterBehavior: independent",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        'annotation PBI_QueryOrder = ["DataFolder","vesta_Document","vesta_DocumentVersion","Models"]',
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        'annotation PBI_ProTooling = ["DevMode"]',
        "",
    ] + [f"ref table {q(t)}" for t in tables] + ["", "ref cultureInfo en-US", "", ""]
    write(d / "model.tmdl", "\n".join(model))
    write(d / "cultures" / "en-US.tmdl", "cultureInfo en-US\n\n")

    data_dir = str(ROOT / "data") + "\\"
    write(d / "expressions.tmdl",
          f'expression DataFolder = "{data_dir}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]\n'
          f"\tlineageTag: {guid('DataFolder')}\n\n"
          "\tannotation PBI_ResultType = Text\n\n"
          "\tannotation PBI_NavigationStepName = Navigation\n\n")

    # ---- vesta_Document
    parts = [f"table vesta_Document\n\tlineageTag: {guid('vesta_Document')}\n"]
    for name, dt, _, fmt, summ in DOC_COLS:
        parts.append(tmdl_column("vesta_Document", name, dt, fmt, summ, source=name))
    for name, dt in DOC_DERIVED:
        parts.append(tmdl_column("vesta_Document", name, dt, source=name))
    doc_steps = """,
    Taxonomy = Table.AddColumn(Typed, "Taxonomy", each Text.Split(Text.AfterDelimiter([description], "template: "), " | "), type list),
    ModelType = Table.AddColumn(Taxonomy, "Model Type", each try [Taxonomy]{1} otherwise null, type text),
    RiskTier = Table.AddColumn(ModelType, "Risk Tier", each try Text.Replace([Taxonomy]{2}, " risk", "") otherwise null, type text),
    Usage = Table.AddColumn(RiskTier, "Usage Category", each try [Taxonomy]{3} otherwise null, type text),
    Result = Table.RemoveColumns(Usage, {"Taxonomy"})"""
    parts.append(csv_partition("vesta_Document", "vesta_Document.csv", DOC_COLS, doc_steps, "Result"))
    parts.append("\tannotation PBI_ResultType = Table\n\n")
    write(d / "tables" / "vesta_Document.tmdl", "\n".join(parts))

    # ---- vesta_DocumentVersion
    V = "vesta_DocumentVersion"
    parts = [f"table {V}\n\tlineageTag: {guid(V)}\n"]
    for name, dt, _, fmt, summ in VER_COLS:
        parts.append(tmdl_column(V, name, dt, fmt, summ, source=name,
                                 sort_by="State Order" if name == "state" else None))
    calc_cols = [
        ("Created Date", "dateTime", "yyyy-mm-dd",
         f"DATE(YEAR({V}[created_on]), MONTH({V}[created_on]), DAY({V}[created_on]))"),
        ("Version Seq", "int64", "0",
         f"VAR d = {V}[document_id] VAR c = {V}[created_on] RETURN COUNTROWS(FILTER({V}, {V}[document_id] = d && {V}[created_on] <= c))"),
        ("Is Latest", "boolean", '"""TRUE"";""TRUE"";""FALSE"""',
         f"{V}[created_on] = CALCULATE(MAX({V}[created_on]), ALLEXCEPT({V}, {V}[document_id]))"),
        ("Document Type", "string", None, f"RELATED(vesta_Document[document_type_code])"),
        ("Model ID", "string", None, f"FORMAT(RELATED(vesta_Document[model_id]), \"0\")"),
    ]
    for name, dt, fmt, expr in calc_cols:
        parts.append(tmdl_column(V, name, dt, fmt, expression=expr))
    # state sort order is computed in Power Query (a DAX column derived from state would be circular)
    parts.append(tmdl_column(V, "State Order", "int64", "0", source="State Order", hidden=True))
    order_step = """,
    StateOrder = Table.AddColumn(Typed, "State Order", each if [state] = "Draft" then 1 else if [state] = "Lock" then 2
        else if [state] = "Commit to Vesta" then 3 else if [state] = "Publish" then 4 else if [state] = "Retired" then 5 else 9, Int64.Type)"""
    parts.append(csv_partition(V, "vesta_DocumentVersion.csv", VER_COLS, order_step, "StateOrder"))
    parts.append("\tannotation PBI_ResultType = Table\n\n")
    write(d / "tables" / f"{V}.tmdl", "\n".join(parts))

    # ---- Models dimension (one row per model_id) built from the document query
    parts = [f"table Models\n\tlineageTag: {guid('Models')}\n"]
    for name, dt, fmt in [("model_id", "int64", "0"), ("Model ID", "string", None), ("Document Types", "string", None),
                          ("Model Type", "string", None), ("Risk Tier", "string", None),
                          ("Usage Category", "string", None)]:
        parts.append(tmdl_column("Models", name, dt, fmt, source=name))
    m = """
let
    Source = vesta_Document,
    Grouped = Table.Group(Source, {"model_id"}, {
        {"Document Types", each Text.Combine(List.Sort(List.Distinct([document_type_code])), " + "), type text},
        {"Model Type", each List.First(List.RemoveNulls([Model Type])), type text},
        {"Risk Tier", each List.First(List.RemoveNulls([Risk Tier])), type text},
        {"Usage Category", each List.First(List.RemoveNulls([Usage Category])), type text}}),
    WithLabel = Table.AddColumn(Grouped, "Model ID", each Text.From([model_id]), type text),
    Ordered = Table.ReorderColumns(WithLabel, {"model_id", "Model ID", "Document Types", "Model Type", "Risk Tier", "Usage Category"}),
    Sorted = Table.Sort(Ordered, {{"model_id", Order.Ascending}})
in
    Sorted
"""
    parts.append(f"\tpartition Models = m\n\t\tmode: import\n\t\tsource =\n{m_block(m)}\n")
    parts.append("\tannotation PBI_ResultType = Table\n\n")
    write(d / "tables" / "Models.tmdl", "\n".join(parts))

    # ---- Calendar (calculated)
    parts = [f"table Calendar\n\tlineageTag: {guid('Calendar')}\n"]
    for name, dt, fmt in [("Date", "dateTime", "yyyy-mm-dd"), ("Week Start", "dateTime", "dd mmm"),
                          ("Month", "string", None), ("Month Num", "int64", "0")]:
        parts.append(tmdl_column("Calendar", name, dt, fmt, source=f"[{name}]", calc_table=True,
                                 sort_by="Month Num" if name == "Month" else None))
    cal = ('ADDCOLUMNS(CALENDAR(DATE(2026, 6, 1), DATE(2026, 10, 31)), "Week Start", [Date] - WEEKDAY([Date], 2) + 1, '
           '"Month", FORMAT([Date], "mmm yyyy"), "Month Num", MONTH([Date]))')
    parts.append(f"\tpartition Calendar = calculated\n\t\tmode: import\n\t\tsource = {cal}\n")
    parts.append("\tannotation PBI_Id = " + guid("cal-id").replace("-", "") + "\n\n")
    write(d / "tables" / "Calendar.tmdl", "\n".join(parts))

    # ---- Workflow Field (calculated, disconnected) for the field-coverage chart
    WF = "Workflow Field"
    parts = [f"table {q(WF)}\n\tlineageTag: {guid(WF)}\n"]
    parts.append(tmdl_column(WF, "Field", "string", source="[Field]", calc_table=True, sort_by="Field Order"))
    parts.append(tmdl_column(WF, "Field Order", "int64", "0", source="[Field Order]", calc_table=True, hidden=True))
    fields = ["Committed to Vesta", "MAR Reference", "VAR ID", "MAR ID", "Based On Version", "Published",
              "Approved"]
    rows = ", ".join(f'{{"{f}", {i + 1}}}' for i, f in enumerate(fields))
    parts.append(f'\tpartition {q(WF)} = calculated\n\t\tmode: import\n\t\tsource = DATATABLE("Field", STRING, "Field Order", INTEGER, {{{rows}}})\n')
    write(d / "tables" / f"{WF}.tmdl", "\n".join(parts))

    # ---- Key Measures
    KM = "Key Measures"
    nb = lambda c: f"CALCULATE(COUNTROWS({V}), NOT ISBLANK({V}[{c}]))"
    measures = [
        ("Documents", "COUNTROWS(vesta_Document)", "#,0", "KPIs"),
        ("Versions", f"COUNTROWS({V})", "#,0", "KPIs"),
        ("Model Count", "DISTINCTCOUNT(vesta_Document[model_id]) + 0", "#,0", "KPIs"),
        ("Multi-Version Documents",
         f"COUNTROWS(FILTER(vesta_Document, CALCULATE(COUNTROWS({V})) > 1)) + 0", "#,0", "KPIs"),
        ("Committed Versions", f"{nb('committed_to_vesta_on')} + 0", "#,0", "KPIs"),
        ("Published Versions", f"{nb('published_date')} + 0", "#,0", "KPIs"),
        ("Approved Versions", f"{nb('approved_on')} + 0", "#,0", "KPIs"),
        ("Commit Rate", "DIVIDE([Committed Versions], [Versions])", "0.0%", "KPIs"),
        ("Draft Share", f'DIVIDE(CALCULATE([Versions], {V}[state] = "Draft"), [Versions])', "0.0%", "KPIs"),
        ("Avg Versions per Document", "DIVIDE([Versions], [Documents])", "0.00", "KPIs"),
        ("Selected Model",
         "IF(HASONEVALUE('Models'[Model ID]), \"Model \" & SELECTEDVALUE('Models'[Model ID]), "
         "\"All models (\" & FORMAT([Model Count], \"0\") & \") - choose a Model ID\")", None, "Labels"),
        ("Selected Model Detail",
         "IF(HASONEVALUE('Models'[Model ID]), SELECTEDVALUE('Models'[Model Type]) & \"  |  \" & SELECTEDVALUE('Models'[Risk Tier]) "
         "& \" risk  |  \" & SELECTEDVALUE('Models'[Usage Category]) & \"  |  Documents: \" & SELECTEDVALUE('Models'[Document Types]), "
         "\"Portfolio view - search the Model ID dropdown to focus on one model\")", None, "Labels"),
        ("Latest State",
         f'VAR v = FILTER({V}, {V}[Is Latest]) VAR n = COUNTROWS(v) RETURN '
         f'IF(n = 0, "-", IF(n <= 2, CONCATENATEX(v, {V}[Document Type] & ": " & {V}[state], "  /  ", {V}[Document Type]), '
         f'FORMAT(CALCULATE(COUNTROWS({V}), {V}[Is Latest], {V}[state] = "Draft"), "0") & " docs in Draft"))', None, "Labels"),
        ("Latest Version",
         f'VAR v = FILTER({V}, {V}[Is Latest]) VAR n = COUNTROWS(v) RETURN '
         f'IF(n = 0, "-", IF(n <= 2, CONCATENATEX(v, {V}[Document Type] & ": " & {V}[version_no], "  /  ", {V}[Document Type]), '
         f'FORMAT(n, "0") & " latest versions"))', None, "Labels"),
        ("Lifecycle Days",
         f"VAR f = MIN({V}[created_on]) VAR l = MAXX({V}, MAX({V}[created_on], COALESCE({V}[last_modified_on], {V}[created_on]))) "
         f"RETURN IF(ISBLANK(f), BLANK(), DATEDIFF(f, l, DAY))", "#,0", "KPIs"),
        ("Field Populated",
         f'SWITCH(SELECTEDVALUE(\'{WF}\'[Field]), "Last Modified", {nb("last_modified_on")}, '
         f'"MAR Reference", {nb("mar_referenceNumber")}, "VAR ID", {nb("var_id")}, '
         f'"Committed to Vesta", {nb("committed_to_vesta_on")}, "MAR ID", {nb("mar_id")}, '
         f'"Based On Version", {nb("based_on_version_id")}, "Published", {nb("published_date")}, '
         f'"Approved", {nb("approved_on")}) + 0', "#,0", "Workflow"),
        ("Field Coverage %", "DIVIDE([Field Populated], [Versions])", "0%", "Workflow"),
    ]
    parts = [f"table {q(KM)}\n\tlineageTag: {guid(KM)}\n"]
    for name, expr, fmt, folder in measures:
        parts.append(tmdl_measure(KM, name, expr, fmt, folder))
    parts.append(tmdl_column(KM, "Dummy", "int64", "0", source="[Dummy]", calc_table=True, hidden=True))
    parts.append(f'\tpartition {q(KM)} = calculated\n\t\tmode: import\n\t\tsource = ROW("Dummy", 0)\n')
    write(d / "tables" / f"{KM}.tmdl", "\n".join(parts))

    rel = [
        (f"{V}.document_id", "vesta_Document.document_id"),
        ("vesta_Document.model_id", "Models.model_id"),
        (f"{V}.'Created Date'", "Calendar.Date"),
    ]
    out = []
    for frm, to in rel:
        out.append(f"relationship {guid('rel' + frm + to)}\n\tfromColumn: {frm}\n\ttoColumn: {to}\n")
    write(d / "relationships.tmdl", "\n".join(out) + "\n")


# =================================================================================== report
def lit(v: str) -> dict:
    return {"expr": {"Literal": {"Value": v}}}


def s(text: str) -> dict:  # string literal
    return lit("'" + text.replace("'", "''") + "'")


def color(hex_: str) -> dict:
    return {"solid": {"color": s(hex_)}}


def num(v: float) -> dict:
    return lit(f"{v}D")


TRUE, FALSE = lit("true"), lit("false")


def col(entity: str, prop: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def meas(prop: str, entity: str = "Key Measures") -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def proj(field: dict, display: str | None = None) -> dict:
    k = "Column" if "Column" in field else "Measure"
    ent = field[k]["Expression"]["SourceRef"]["Entity"]
    prop = field[k]["Property"]
    p = {"field": field, "queryRef": f"{ent}.{prop}", "nativeQueryRef": prop}
    if display:
        p["displayName"] = display
    return p


def container(title: str | None, bg: str = "#FFFFFF", border: bool = True) -> dict:
    o = {
        "background": [{"properties": {"show": TRUE, "color": color(bg), "transparency": num(0)}}],
        "border": [{"properties": {"show": lit("true" if border else "false"), "color": color(WF_LINE),
                                   "radius": num(6)}}],
        "dropShadow": [{"properties": {"show": FALSE}}],
        "subTitle": [{"properties": {"show": FALSE}}],
        "padding": [{"properties": {"top": num(8), "bottom": num(8), "left": num(10), "right": num(10)}}],
    }
    if title:
        o["title"] = [{"properties": {"show": TRUE, "text": s(title), "fontColor": color(WF_CHARCOAL),
                                      "fontSize": num(12), "bold": TRUE, "fontFamily": s("Segoe UI Semibold")}}]
    else:
        o["title"] = [{"properties": {"show": FALSE}}]
    return o


def value_selector(field: dict, value: str) -> dict:
    return {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": field,
                                                 "Right": {"Literal": {"Value": "'" + value + "'"}}}}}]}


def state_colors(field: dict) -> list:
    return [{"properties": {"fill": color(c)}, "selector": value_selector(field, st)} for st, c in STATE_COLORS.items()]


class Page:
    def __init__(self, key: str, display: str, bg: str = WF_CREAM, width: int = 1280, height: int = 720):
        self.name = guid("page-" + key).replace("-", "")[:20]
        self.display, self.bg, self.width, self.height = display, bg, width, height
        self.visuals: list[dict] = []

    def add(self, key: str, x, y, w, h, visual: dict, z: int | None = None, extra: dict | None = None):
        name = guid(self.name + key).replace("-", "")[:20]
        z = len(self.visuals) * 1000 if z is None else z
        v = {"$schema": VISUAL_SCHEMA, "name": name,
             "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
             "visual": visual}
        if extra:
            v.update(extra)
        self.visuals.append(v)


def textbox(runs: list[tuple[str, dict]], bg: str | None = None, align: str = "left") -> dict:
    vis = {"visualType": "textbox",
           "objects": {"general": [{"properties": {"paragraphs": [
               {"textRuns": [{"value": t, "textStyle": st} for t, st in runs], "horizontalTextAlignment": align}]}}]},
           "visualContainerObjects": container(None, bg or "#FFFFFF", border=False)}
    if bg is None:
        vis["visualContainerObjects"]["background"] = [{"properties": {"show": FALSE}}]
    return vis


def shape_bar(fill: str) -> dict:
    return {"visualType": "shape",
            "objects": {"shape": [{"properties": {"tileShape": s("rectangle")}}],
                        "fill": [{"properties": {"show": TRUE, "fillColor": color(fill), "transparency": num(0)}}],
                        "outline": [{"properties": {"show": FALSE}}]},
            "visualContainerObjects": {"background": [{"properties": {"show": FALSE}}],
                                       "title": [{"properties": {"show": FALSE}}]}}


def card(measure: str, label: str, accent: str = WF_RED, size: int = 26) -> dict:
    return {"visualType": "card",
            "query": {"queryState": {"Values": {"projections": [proj(meas(measure), label)]}}},
            "objects": {"labels": [{"properties": {"color": color(accent), "fontSize": num(size),
                                                   "fontFamily": s("Segoe UI Semibold")}}],
                        "categoryLabels": [{"properties": {"show": TRUE, "color": color(WF_CHARCOAL),
                                                           "fontSize": num(10)}}]},
            "visualContainerObjects": container(None)}


def text_card(measure: str, color_: str, size: int, show_label: bool = False, bg: str = "#FFFFFF") -> dict:
    v = card(measure, measure, color_, size)
    v["objects"]["categoryLabels"] = [{"properties": {"show": lit("true" if show_label else "false")}}]
    v["objects"]["labels"][0]["properties"]["fontFamily"] = s("Segoe UI")
    v["visualContainerObjects"] = container(None, bg, border=False)
    return v


def header(page: Page, title: str, subtitle: str) -> None:
    page.add("hdr", 0, 0, page.width, 64, textbox(
        [(title, {"fontWeight": "bold", "fontSize": "20pt", "color": "#FFFFFF", "fontFamily": "Segoe UI"}),
         ("   " + subtitle, {"fontSize": "11pt", "color": WF_GOLD, "fontFamily": "Segoe UI"})], bg=WF_RED))
    page.add("gold", 0, 64, page.width, 5, shape_bar(WF_GOLD))


V = "vesta_DocumentVersion"
D = "vesta_Document"


def model_explorer() -> Page:
    p = Page("explorer", "Model Explorer")
    header(p, "Vesta Model Document Explorer", "Search and select a Model ID to see its documents, versions and workflow")

    # searchable single-select dropdown slicer on Model ID
    slicer = {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [proj(col("Models", "Model ID"))]}}},
        "objects": {
            "data": [{"properties": {"mode": s("Dropdown")}}],
            "general": [{"properties": {"selfFilterEnabled": TRUE, "orientation": lit("0D")}}],
            "selection": [{"properties": {"singleSelect": TRUE, "strictSingleSelect": FALSE, "selectAllCheckboxEnabled": FALSE}}],
            "header": [{"properties": {"show": TRUE, "text": s("Model ID  (type to search)"), "fontColor": color(WF_RED),
                                       "textSize": num(11), "bold": TRUE}}],
            "items": [{"properties": {"fontColor": color(WF_CHARCOAL), "textSize": num(12), "background": color("#FFFFFF")}}],
        },
        "visualContainerObjects": container(None),
    }
    p.add("slicer", 16, 82, 280, 78, slicer)
    p.add("title", 308, 82, 956, 44, text_card("Selected Model", WF_RED, 20))
    p.add("subtitle", 308, 122, 956, 38, text_card("Selected Model Detail", WF_CHARCOAL, 11))

    kpis = [("Documents", "Documents", WF_RED, 26), ("Versions", "Versions", WF_RED, 26),
            ("Latest State", "Latest State", WF_CHARCOAL, 13), ("Latest Version", "Latest Version", WF_CHARCOAL, 13),
            ("Committed Versions", "Committed to Vesta", WF_BURGUNDY, 26),
            ("Published Versions", "Published", WF_BURGUNDY, 26), ("Lifecycle Days", "Lifecycle (days)", WF_CHARCOAL, 26)]
    w = (1248 - 6 * 8) / 7
    for i, (m, lbl, c, sz) in enumerate(kpis):
        p.add(f"kpi{i}", round(16 + i * (w + 8)), 170, round(w), 92, card(m, lbl, c, sz))

    state = col(V, "state")
    donut = {"visualType": "donutChart",
             "query": {"queryState": {"Category": {"projections": [proj(state, "State")]},
                                      "Y": {"projections": [proj(meas("Versions"))]}}},
             "objects": {"dataPoint": state_colors(state),
                         "legend": [{"properties": {"show": TRUE, "position": s("Bottom")}}],
                         "labels": [{"properties": {"show": TRUE, "labelStyle": s("Data value")}}]},
             "visualContainerObjects": container("Versions by Lifecycle State")}
    p.add("donut", 16, 272, 372, 236, donut)

    timeline = {"visualType": "columnChart",
                "query": {"queryState": {
                    "Category": {"projections": [proj(col("Calendar", "Week Start"), "Week")]},
                    "Series": {"projections": [proj(state, "State")]},
                    "Y": {"projections": [proj(meas("Versions"))]}},
                    "sortDefinition": {"sort": [{"field": col("Calendar", "Week Start"), "direction": "Ascending"}],
                                       "isDefaultSort": True}},
                "objects": {"dataPoint": state_colors(state),
                            "legend": [{"properties": {"show": TRUE, "position": s("Top")}}],
                            "categoryAxis": [{"properties": {"axisType": s("Categorical"), "showAxisTitle": FALSE}}],
                            "valueAxis": [{"properties": {"showAxisTitle": FALSE, "gridlineShow": FALSE}}]},
                "visualContainerObjects": container("Version Timeline (versions created per week)")}
    p.add("timeline", 396, 272, 476, 236, timeline)

    wf = col("Workflow Field", "Field")
    bar = {"visualType": "clusteredBarChart",
           "query": {"queryState": {"Category": {"projections": [proj(wf, "Workflow field")]},
                                    "Y": {"projections": [proj(meas("Field Populated"), "Versions with field")]}},
                     "sortDefinition": {"sort": [{"field": col("Workflow Field", "Field"), "direction": "Ascending"}],
                                        "isDefaultSort": True}},
           "objects": {"dataPoint": [{"properties": {"fill": color(WF_RED)}}],
                       "labels": [{"properties": {"show": TRUE, "color": color(WF_CHARCOAL), "fontSize": num(9)}}],
                       "categoryAxis": [{"properties": {"showAxisTitle": FALSE, "fontSize": num(9),
                                                        "innerPadding": num(15)}}],
                       "valueAxis": [{"properties": {"show": FALSE, "gridlineShow": FALSE}}]},
           "visualContainerObjects": container("Workflow Field Coverage")}
    p.add("wfbar", 880, 272, 384, 236, bar)

    vt_cols = [(col(V, "document_version_id"), "Version ID"), (col(V, "document_id"), "Doc ID"),
               (col(V, "Document Type"), "Type"), (col(V, "version_no"), "Version"), (state, "State"),
               (col(V, "created_on"), "Created"), (col(V, "created_by"), "Created By"),
               (col(V, "based_on_version_id"), "Based On"), (col(V, "committed_to_vesta_on"), "Committed to Vesta"),
               (col(V, "published_date"), "Published"), (col(V, "mar_id"), "MAR ID"), (col(V, "var_id"), "VAR ID")]
    table_obj = {"columnHeaders": [{"properties": {"backColor": color(WF_RED), "fontColor": color("#FFFFFF"),
                                                   "bold": TRUE}}],
                 "values": [{"properties": {"backColorPrimary": color("#FFFFFF"), "backColorSecondary": color("#FBF7F3")}}],
                 "grid": [{"properties": {"gridHorizontalColor": color(WF_LINE)}}]}
    vtable = {"visualType": "tableEx",
              "query": {"queryState": {"Values": {"projections": [proj(f, n) for f, n in vt_cols]}},
                        "sortDefinition": {"sort": [{"field": col(V, "created_on"), "direction": "Ascending"}]}},
              "objects": table_obj,
              "visualContainerObjects": container("Version History")}
    p.add("vtable", 16, 516, 856, 192, vtable)

    dt_cols = [(col(D, "document_id"), "Doc ID"), (col(D, "document_type_code"), "Type"),
               (col(D, "created_on"), "Created"), (col(D, "created_by"), "Created By"),
               (col(D, "description"), "Template Description")]
    dtable = {"visualType": "tableEx",
              "query": {"queryState": {"Values": {"projections": [proj(f, n) for f, n in dt_cols]}}},
              "objects": table_obj,
              "visualContainerObjects": container("Documents for this Model")}
    p.add("dtable", 880, 516, 384, 192, dtable)
    return p


def overview() -> Page:
    p = Page("overview", "Portfolio Overview")
    header(p, "Vesta Document Portfolio", "All models  |  2026 lifecycle extract (synthetic)")
    kpis = [("Model Count", "Models"), ("Documents", "Documents"), ("Versions", "Versions"),
            ("Multi-Version Documents", "Docs with History"), ("Commit Rate", "Committed to Vesta"),
            ("Published Versions", "Published"), ("Draft Share", "Draft Share")]
    w = (1248 - 6 * 8) / 7
    for i, (m, lbl) in enumerate(kpis):
        p.add(f"kpi{i}", round(16 + i * (w + 8)), 82, round(w), 92, card(m, lbl, WF_RED if i < 4 else WF_BURGUNDY))

    state = col(V, "state")
    dtype = col(V, "Document Type")
    stacked = {"visualType": "barChart",
               "query": {"queryState": {"Category": {"projections": [proj(state, "State")]},
                                        "Series": {"projections": [proj(dtype, "Document Type")]},
                                        "Y": {"projections": [proj(meas("Versions"))]}},
                         "sortDefinition": {"sort": [{"field": state, "direction": "Ascending"}], "isDefaultSort": True}},
               "objects": {"dataPoint": [{"properties": {"fill": color(c)}, "selector": value_selector(dtype, t)}
                                         for t, c in TYPE_COLORS.items()],
                           "legend": [{"properties": {"show": TRUE, "position": s("Top")}}],
                           "labels": [{"properties": {"show": TRUE}}],
                           "valueAxis": [{"properties": {"show": FALSE, "gridlineShow": FALSE}}],
                           "categoryAxis": [{"properties": {"showAxisTitle": FALSE}}]},
               "visualContainerObjects": container("Lifecycle State by Document Type")}
    p.add("stack", 16, 184, 408, 254, stacked)

    trend = {"visualType": "columnChart",
             "query": {"queryState": {"Category": {"projections": [proj(col("Calendar", "Week Start"), "Week")]},
                                      "Series": {"projections": [proj(state, "State")]},
                                      "Y": {"projections": [proj(meas("Versions"))]}},
                       "sortDefinition": {"sort": [{"field": col("Calendar", "Week Start"), "direction": "Ascending"}],
                                          "isDefaultSort": True}},
             "objects": {"dataPoint": state_colors(state),
                         "legend": [{"properties": {"show": TRUE, "position": s("Top")}}],
                         "categoryAxis": [{"properties": {"axisType": s("Categorical"), "showAxisTitle": FALSE}}],
                         "valueAxis": [{"properties": {"showAxisTitle": FALSE}}]},
             "visualContainerObjects": container("Versions Created per Week")}
    p.add("trend", 432, 184, 832, 254, trend)

    risk = col("Models", "Risk Tier")
    donut = {"visualType": "donutChart",
             "query": {"queryState": {"Category": {"projections": [proj(risk, "Risk Tier")]},
                                      "Y": {"projections": [proj(meas("Model Count"), "Models")]}}},
             "objects": {"dataPoint": [{"properties": {"fill": color(c)}, "selector": value_selector(risk, r)}
                                       for r, c in {"High": WF_RED, "Moderate": WF_GOLD, "Low": WF_TAUPE}.items()],
                         "legend": [{"properties": {"show": TRUE, "position": s("Bottom")}}],
                         "labels": [{"properties": {"show": TRUE, "labelStyle": s("Data value")}}]},
             "visualContainerObjects": container("Models by Risk Tier")}
    p.add("risk", 16, 448, 300, 256, donut)

    mtype = col("Models", "Model Type")
    bar = {"visualType": "clusteredBarChart",
           "query": {"queryState": {"Category": {"projections": [proj(mtype, "Model Type")]},
                                    "Y": {"projections": [proj(meas("Model Count"), "Models")]}},
                     "sortDefinition": {"sort": [{"field": meas("Model Count"), "direction": "Descending"}], "isDefaultSort": True}},
           "objects": {"dataPoint": [{"properties": {"fill": color(WF_RED)}}],
                       "labels": [{"properties": {"show": TRUE}}],
                       "valueAxis": [{"properties": {"show": FALSE, "gridlineShow": FALSE}}],
                       "categoryAxis": [{"properties": {"showAxisTitle": FALSE}}]},
           "visualContainerObjects": container("Models by Model Type")}
    p.add("mtype", 324, 448, 316, 256, bar)

    usage = col("Models", "Usage Category")
    ubar = {"visualType": "clusteredColumnChart",
            "query": {"queryState": {"Category": {"projections": [proj(usage, "Usage Category")]},
                                     "Y": {"projections": [proj(meas("Model Count"), "Models")]}},
                      "sortDefinition": {"sort": [{"field": meas("Model Count"), "direction": "Descending"}], "isDefaultSort": True}},
            "objects": {"dataPoint": [{"properties": {"fill": color(WF_GOLD)}}],
                        "labels": [{"properties": {"show": TRUE}}],
                        "valueAxis": [{"properties": {"show": FALSE, "gridlineShow": FALSE}}],
                        "categoryAxis": [{"properties": {"showAxisTitle": FALSE}}]},
            "visualContainerObjects": container("Models by Usage Category")}
    p.add("usage", 648, 448, 300, 256, ubar)

    wf = col("Workflow Field", "Field")
    cov = {"visualType": "clusteredBarChart",
           "query": {"queryState": {"Category": {"projections": [proj(wf, "Workflow field")]},
                                    "Y": {"projections": [proj(meas("Field Coverage %"), "Coverage")]}},
                     "sortDefinition": {"sort": [{"field": wf, "direction": "Ascending"}], "isDefaultSort": True}},
           "objects": {"dataPoint": [{"properties": {"fill": color(WF_CHARCOAL)}}],
                       "labels": [{"properties": {"show": TRUE, "fontSize": num(9)}}],
                       "valueAxis": [{"properties": {"show": FALSE, "gridlineShow": FALSE}}],
                       "categoryAxis": [{"properties": {"showAxisTitle": FALSE, "fontSize": num(9),
                                                        "innerPadding": num(15)}}]},
           "visualContainerObjects": container("Workflow Field Coverage %")}
    p.add("cov", 956, 448, 308, 256, cov)
    return p


def theme() -> dict:
    return {
        "name": "Wells Fargo Vesta",
        "dataColors": [WF_RED, WF_GOLD, WF_CHARCOAL, WF_BURGUNDY, WF_ORANGE, WF_TAUPE, "#E8868B", "#C9A227",
                       "#6E6259", "#F4B6A6"],
        "foreground": WF_CHARCOAL,
        "foregroundNeutralSecondary": "#5E5550",
        "background": "#FFFFFF",
        "backgroundLight": "#F4F0ED",
        "tableAccent": WF_RED,
        "good": "#2E7D32", "neutral": WF_GOLD, "bad": WF_RED,
        "maximum": WF_RED, "center": WF_GOLD, "minimum": "#FFF4D6",
        "hyperlink": WF_RED, "visitedHyperlink": WF_BURGUNDY,
        "textClasses": {
            "callout": {"fontSize": 26, "fontFace": "Segoe UI Semibold", "color": WF_RED},
            "title": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": WF_CHARCOAL},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": WF_CHARCOAL},
            "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": WF_CHARCOAL},
        },
        "visualStyles": {
            "*": {"*": {
                "border": [{"show": True, "color": {"solid": {"color": WF_LINE}}, "radius": 6}],
                "dropShadow": [{"show": False}],
            }},
            "page": {"*": {"background": [{"color": {"solid": {"color": WF_CREAM}}, "transparency": 0}],
                           "outspace": [{"color": {"solid": {"color": WF_CREAM}}, "transparency": 0}]}},
        },
    }


def build_report() -> None:
    if RP.exists():
        shutil.rmtree(RP)
    d = RP / "definition"
    write_json(RP / ".platform", {
        "$schema": PLATFORM_SCHEMA,
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid("report")},
    })
    write_json(RP / "definition.pbir", {"version": "4.0",
                                        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(d / "version.json", {"$schema": VERSION_SCHEMA, "version": "2.0.0"})
    shared = RP / "StaticResources" / "SharedResources" / "BaseThemes"
    shared.mkdir(parents=True, exist_ok=True)
    shutil.copy(BASE_THEME_FILE, shared / f"{BASE_THEME}.json")
    write_json(RP / "StaticResources" / "RegisteredResources" / CUSTOM_THEME, theme())
    write_json(d / "report.json", {
        "$schema": REPORT_SCHEMA,
        "themeCollection": {
            "baseTheme": {"name": BASE_THEME, "reportVersionAtImport": VERSION_AT_IMPORT, "type": "SharedResources"},
            "customTheme": {"name": CUSTOM_THEME, "reportVersionAtImport": VERSION_AT_IMPORT,
                            "type": "RegisteredResources"},
        },
        "objects": {"section": [{"properties": {"verticalAlignment": s("Top")}}]},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": CUSTOM_THEME, "path": CUSTOM_THEME, "type": "CustomTheme"}]},
        ],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True},
    })
    pages = [model_explorer(), overview()]
    write_json(d / "pages" / "pages.json", {"$schema": PAGES_SCHEMA, "pageOrder": [p.name for p in pages],
                                            "activePageName": pages[0].name})
    for p in pages:
        pd = d / "pages" / p.name
        write_json(pd / "page.json", {
            "$schema": PAGE_SCHEMA, "name": p.name, "displayName": p.display, "displayOption": "FitToPage",
            "height": p.height, "width": p.width,
            "objects": {"background": [{"properties": {"color": color(p.bg), "transparency": num(0)}}],
                        "outspace": [{"properties": {"color": color(p.bg), "transparency": num(0)}}]},
        })
        for v in p.visuals:
            write_json(pd / "visuals" / v["name"] / "visual.json", v)


def build_pbip() -> None:
    write_json(ROOT / f"{NAME}.pbip", {"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                       "settings": {"enableAutoRecovery": True}})
    write(ROOT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")


if __name__ == "__main__":
    build_semantic_model()
    build_report()
    build_pbip()
    print("built", ROOT / f"{NAME}.pbip")
