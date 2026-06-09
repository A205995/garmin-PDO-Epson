#!/usr/bin/env python3
import json
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
HTML = ROOT / "index.html"
OUT_DIR = ROOT / "outputs"
OUT_FILE = OUT_DIR / "garmin_pdo_epson_robot_database.xlsx"


def extract_app_data():
    js = r"""
const fs = require("fs");
const html = fs.readFileSync(process.argv[1], "utf8");
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];
const cut = script.split("function setTab")[0];
eval(cut + `
process.stdout.write(JSON.stringify({
  exportedAt: new Date().toISOString(),
  models,
  purchasedModels,
  productCatalog
}, null, 2));
`);
"""
    result = subprocess.run(
        ["node", "-e", js, str(HTML)],
        check=True,
        text=True,
        capture_output=True,
    )
    return json.loads(result.stdout)


def value(row, key):
    val = row.get(key, "")
    if val is None:
        return ""
    if isinstance(val, list):
        if val and isinstance(val[0], dict):
            return json.dumps(val, ensure_ascii=False)
        return " / ".join(str(x) for x in val)
    if isinstance(val, dict):
        return json.dumps(val, ensure_ascii=False)
    return val


def col_name(index):
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(row_idx, col_idx, val, style=0):
    ref = f"{col_name(col_idx)}{row_idx}"
    style_attr = f' s="{style}"' if style else ""
    if val is None or val == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(val, bool):
        return f'<c r="{ref}" t="b"{style_attr}><v>{1 if val else 0}</v></c>'
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return f'<c r="{ref}"{style_attr}><v>{val}</v></c>'
    text = escape(str(val), {'"': "&quot;"})
    preserve = ' xml:space="preserve"' if text.strip() != text else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t{preserve}>{text}</t></is></c>'


def worksheet_xml(rows, widths=None):
    max_cols = max((len(r) for r in rows), default=1)
    max_rows = max(len(rows), 1)
    dim = f"A1:{col_name(max_cols)}{max_rows}"
    cols = ""
    if widths:
        parts = []
        for idx, width in enumerate(widths, start=1):
            parts.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
        cols = f"<cols>{''.join(parts)}</cols>"

    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        style = 1 if r_idx == 1 else 0
        cells = "".join(cell_xml(r_idx, c_idx, val, style) for c_idx, val in enumerate(row, start=1))
        sheet_rows.append(f'<row r="{r_idx}">{cells}</row>')

    autofilter = f'<autoFilter ref="{dim}"/>' if len(rows) > 1 and max_cols > 1 else ""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dim}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  {cols}
  <sheetData>{''.join(sheet_rows)}</sheetData>
  {autofilter}
</worksheet>'''


def build_rows(data):
    models = data["models"]
    purchased = data["purchasedModels"]
    catalog = data["productCatalog"]

    robot_cols = [
        "model", "series", "kind", "type", "env", "baseModel", "verified", "isNew",
        "reach", "z", "ratedLoad", "maxLoad", "controller", "mount",
        "j4Rated", "j4Max", "j5Max", "j6Max", "eccRated", "eccMax",
        "j3Force", "j3Speed", "j4Speed", "speedXY", "repeat", "ct",
        "weight", "power", "voltage", "source", "note"
    ]
    robot_headers = [
        "型號", "系列", "軸數類型", "型式", "環境", "基準型號", "已驗證", "官方列表",
        "可達距離 mm", "Z 行程 mm", "額定負載 kg", "最大負載 kg", "控制器", "安裝方式",
        "J4 額定慣量 kgm2", "J4 最大慣量 kgm2", "J5 最大慣量", "J6 最大慣量", "偏心率額定 mm", "偏心率最大 mm",
        "J3 下壓力 N", "J3 速度 mm/s", "J4 速度 deg/s", "水平速度 mm/s", "重複精度", "循環時間 s",
        "本體重量", "電源容量", "電源電壓", "資料來源", "備註"
    ]
    robot_rows = [robot_headers]
    for m in sorted(models, key=lambda x: (x.get("series", ""), x.get("model", ""))):
        robot_rows.append([value(m, c) for c in robot_cols])

    download_rows = [["型號", "官方型號頁", "系列手冊", "基準型號", "下載連結策略"]]
    for m in sorted(models, key=lambda x: x.get("model", "")):
        download_rows.append([
            value(m, "model"),
            value(m, "official"),
            value(m, "manual"),
            value(m, "baseModel"),
            "一律導向官方網頁或系列手冊，避免型號與檔案不一致。"
        ])

    purchased_rows = [["照片型號", "對應工具型號", "前購未稅價", "備註"]]
    for p in purchased:
        purchased_rows.append([value(p, "raw"), value(p, "model"), value(p, "price"), value(p, "note")])

    official_rows = [["官方列表型號", "系列", "軸數類型", "名稱", "新品"]]
    for item in catalog:
        official_rows.append([
            value(item, "model"),
            value(item, "series"),
            value(item, "kind"),
            value(item, "name"),
            value(item, "isNew"),
        ])

    source_map = {}
    for m in models:
        src = value(m, "source")
        if src:
            source_map.setdefault(src, []).append(m["model"])
        manual = value(m, "manual")
        if manual:
            source_map.setdefault(manual, []).append(m["model"])
    source_rows = [["來源", "涵蓋型號數", "代表型號"]]
    for src, linked in sorted(source_map.items(), key=lambda x: x[0]):
        source_rows.append([src, len(linked), " / ".join(linked[:20])])

    summary_rows = [
        ["Garmin PDO Epson 選型資料庫"],
        ["匯出時間", datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")],
        ["來源檔案", str(HTML)],
        ["型號筆數", len(models)],
        ["已購資料筆數", len(purchased)],
        ["官方列表型號筆數", len(catalog)],
        [],
        ["工作表", "用途"],
        ["Robots", "型號主資料，供型號查詢、條件推薦、工程確認使用。"],
        ["Downloads", "官方型號頁與系列手冊連結。"],
        ["Purchased", "Garmin 已購資料，含對應工具型號與價格。"],
        ["Official_List", "官方頁 8 款列表型號。"],
        ["Sources", "資料來源與涵蓋型號索引。"],
    ]

    return [
        ("Summary", summary_rows, [32, 95]),
        ("Robots", robot_rows, [18, 10, 12, 12, 10, 18, 10, 10, 13, 13, 13, 13, 18, 24, 18, 18, 15, 15, 16, 16, 13, 15, 15, 15, 32, 12, 14, 12, 16, 40, 60]),
        ("Downloads", download_rows, [18, 54, 54, 18, 54]),
        ("Purchased", purchased_rows, [32, 22, 14, 60]),
        ("Official_List", official_rows, [18, 10, 12, 34, 10]),
        ("Sources", source_rows, [72, 14, 90]),
    ]


def write_xlsx(sheets):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    sheet_entries = []
    rel_entries = []
    for idx, (name, _rows, _widths) in enumerate(sheets, start=1):
        sheet_entries.append(f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>')
        rel_entries.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    rel_entries.append(f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')

    content_overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for idx in range(1, len(sheets) + 1):
        content_overrides.append(f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')

    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF111111"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1"/></cellXfs>
</styleSheet>'''

    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{''.join(sheet_entries)}</sheets>
</workbook>'''

    workbook_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(rel_entries)}</Relationships>'''

    content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(content_overrides)}
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Garmin PDO Epson Robot Database</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    app = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheets)}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="{len(sheets)}" baseType="lpstr">{''.join(f'<vt:lpstr>{escape(name)}</vt:lpstr>' for name, _, _ in sheets)}</vt:vector></TitlesOfParts>
</Properties>'''

    with zipfile.ZipFile(OUT_FILE, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", styles)
        for idx, (_name, rows, widths) in enumerate(sheets, start=1):
            z.writestr(f"xl/worksheets/sheet{idx}.xml", worksheet_xml(rows, widths))


def main():
    if not HTML.exists():
        print(f"Missing {HTML}", file=sys.stderr)
        return 1
    data = extract_app_data()
    sheets = build_rows(data)
    write_xlsx(sheets)
    print(OUT_FILE)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
