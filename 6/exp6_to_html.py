# -*- coding: utf-8 -*-
# =========================================================
# 法拉第實驗 → 法拉第實驗 → python code for HTML 分析報告 from Exp6.xlsx
# 功能：
# 1. 讀取 Exp6.xlsx
# 2. 解析 A、B、C、D、E 五個工作表
# 3. 擷取 NDAD、頻率、距離、角度、理論值、實驗值、誤差
# 4. 產生 Plotly 互動式圖表
# 5. 產生 HTML 表格
# 6. 自動抽出 E 工作表中的內嵌圖片（波形圖）
# 7. 最後輸出成 exp6_report.html
# =========================================================

from pathlib import Path              # 處理檔案與路徑
import zipfile                        # xlsx 檔本質上是 zip 壓縮格式
import xml.etree.ElementTree as ET    # 解析 Excel 內部 XML
import re                             # 正則表達式，用來解析儲存格座標
import json                           # 產生 Plotly 所需的 JSON 資料
import base64                         # 圖片轉成 base64，方便直接嵌入 HTML
import math                           # 數學模組（保留作擴充用）


# =========================================================
# Excel XML 命名空間
# 不同節點類型要搭配對應 namespace 才能正確找到
# =========================================================
NS_MAIN = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
NS_A = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def col_to_num(col_letters: str) -> int:
    """
    將 Excel 欄位字母轉成數字。
    例如：
    A  -> 1
    B  -> 2
    Z  -> 26
    AA -> 27
    """
    n = 0
    for ch in col_letters:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("a:si", NS_MAIN):
        texts = [t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
        out.append("".join(texts))
    return out


def workbook_sheet_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    """
    讀取 workbook.xml 與 workbook.xml.rels，
    找出每個工作表名稱（如 A、B、C）對應到哪個 XML 檔案。

    回傳格式範例：
    {
        "A": "xl/worksheets/sheet1.xml",
        "B": "xl/worksheets/sheet2.xml"
    }
    """
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    mapping = {}
    for s in wb.find("a:sheets", NS_MAIN):
        name = s.attrib["name"]
        rid = s.attrib[f"{{{NS_REL_OFFICE}}}id"]
        mapping[name] = "xl/" + rid_to_target[rid]
    return mapping


def parse_sheet_to_cells(zf: zipfile.ZipFile, sheet_xml_path: str, shared_strings: list[str]) -> dict[tuple[int, int], dict]:
    """
    將某一張工作表 XML 轉成 Python 字典。

    key   = (row, col)
    value = {
        "value": 儲存格顯示值,
        "formula": 儲存格公式（若有）
    }

    如此後面即可用 (列, 欄) 或 A2、D3 這類位置快速取值。
    """
    xml = ET.fromstring(zf.read(sheet_xml_path))
    cells = {}
    for c in xml.findall(".//a:c", NS_MAIN):
        ref = c.attrib["r"]                       # 儲存格位置，例如 D3
        m = re.match(r"([A-Z]+)(\d+)", ref)
        if not m:
            continue
        col_letters, row_s = m.groups()
        row = int(row_s)
        col = col_to_num(col_letters)

        cell_type = c.attrib.get("t")            # 儲存格型態，例如 s = shared string
        formula_el = c.find("a:f", NS_MAIN)      # 公式節點
        value_el = c.find("a:v", NS_MAIN)        # 值節點
        inline_el = c.find("a:is", NS_MAIN)      # 內嵌文字節點

        value = None
        if value_el is not None:
            raw = value_el.text
            if cell_type == "s" and raw is not None:
                # 若是 shared string，需依索引到 shared_strings 取出真正文字
                value = shared_strings[int(raw)]
            else:
                value = raw
        elif inline_el is not None:
            # 若是 inline string，直接把所有文字節點串起來
            texts = [t.text or "" for t in inline_el.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
            value = "".join(texts)

        cells[(row, col)] = {
            "value": value,
            "formula": formula_el.text if formula_el is not None else None,
        }
    return cells


def cell_value(cells: dict, ref: str):
    """
    用 Excel 座標（例如 A2、D3）讀取儲存格內容。
    若找不到則回傳 None。
    """
    m = re.match(r"([A-Z]+)(\d+)", ref)
    col = col_to_num(m.group(1))
    row = int(m.group(2))
    return cells.get((row, col), {}).get("value")


def cell_num(cells: dict, ref: str):
    """
    讀取儲存格內容並轉成 float。
    若儲存格是空白或無法轉數字，則回傳 None。
    """
    v = cell_value(cells, ref)
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def build_matrix(cells: dict, row_start: int, row_end: int, col_start: int, col_end: int) -> list[list[str]]:
    """
    將某個矩形區域的儲存格資料轉成二維矩陣（list of lists）。
    這個函式保留作為後續擴充，目前主流程未直接使用。
    """
    matrix = []
    for r in range(row_start, row_end + 1):
        row_vals = []
        for c in range(col_start, col_end + 1):
            row_vals.append(cells.get((r, c), {}).get("value"))
        matrix.append(row_vals)
    return matrix


def parse_images_for_sheet(zf: zipfile.ZipFile, sheet_xml_path: str) -> list[dict]:
    """
    從指定工作表中找出內嵌圖片，並轉成可直接放入 HTML 的 base64 資料。

    回傳格式：
    [
        {
            "row": 圖片錨定起始列,
            "col": 圖片錨定起始欄,
            "mime": 圖片 MIME 類型,
            "data_uri": base64 圖片字串
        },
        ...
    ]
    """
    # 先找該工作表的關聯檔（rels）
    sheet_name = Path(sheet_xml_path).name
    rel_path = f"xl/worksheets/_rels/{sheet_name}.rels"
    if rel_path not in zf.namelist():
        return []

    rels_root = ET.fromstring(zf.read(rel_path))
    drawing_target = None

    # 找到該工作表連到哪個 drawing 檔
    for rel in rels_root:
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/drawing"):
            drawing_target = rel.attrib["Target"]
            break
    if not drawing_target:
        return []

    # 進一步找到 drawing 檔與其關聯檔
    drawing_path = "xl/" + drawing_target.replace("../", "")
    drawing_rels_path = "xl/drawings/_rels/" + Path(drawing_path).name + ".rels"
    if drawing_path not in zf.namelist() or drawing_rels_path not in zf.namelist():
        return []

    drawing_root = ET.fromstring(zf.read(drawing_path))
    drawing_rels_root = ET.fromstring(zf.read(drawing_rels_path))
    rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in drawing_rels_root}

    images = []
    for anchor in drawing_root.findall("xdr:twoCellAnchor", NS_XDR):
        from_el = anchor.find("xdr:from", NS_XDR)
        pic = anchor.find("xdr:pic", NS_XDR)
        if from_el is None or pic is None:
            continue

        # 取得圖片錨定的列、欄位置（方便排序）
        row = int(from_el.findtext("xdr:row", default="0", namespaces=NS_XDR))
        col = int(from_el.findtext("xdr:col", default="0", namespaces=NS_XDR))

        # 找到圖片的關聯 id
        blip = pic.find(".//a:blip", {**NS_XDR, **NS_A, "r": NS_REL_OFFICE})
        if blip is None:
            continue
        rid = blip.attrib.get(f"{{{NS_REL_OFFICE}}}embed")
        if not rid or rid not in rid_to_target:
            continue

        # 找到圖片實體檔案路徑
        target = rid_to_target[rid].replace("../", "")
        media_path = "xl/" + target
        if media_path not in zf.namelist():
            continue

        # 依副檔名判斷 MIME type
        ext = Path(media_path).suffix.lower().replace(".", "")
        mime = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
            "webp": "image/webp",
        }.get(ext, "application/octet-stream")

        # 讀圖片後轉成 base64，方便直接嵌入 HTML
        b64 = base64.b64encode(zf.read(media_path)).decode("ascii")
        images.append({
            "row": row + 1,
            "col": col + 1,
            "mime": mime,
            "data_uri": f"data:{mime};base64,{b64}",
        })

    # 依圖片在表中的位置排序
    images.sort(key=lambda x: (x["row"], x["col"]))
    return images


def format_num(v, digits=3):
    """
    將數值格式化成指定小數位數字串。
    若值是空白則回傳空字串。
    若是可轉成數字的字串，也會先轉數字再格式化。
    """
    if v is None:
        return ""
    if isinstance(v, str):
        try:
            v = float(v)
        except Exception:
            return v
    return f"{v:.{digits}f}"


def format_pct(v):
    """
    將數值格式化成百分比字串（保留兩位小數）。
    """
    if v is None:
        return ""
    return f"{v:.2f}%"


def table_html(headers, rows, formats=None):
    """
    將表頭與資料列轉成 HTML 表格。

    headers: 表頭 list
    rows:    二維資料 list
    formats: 可選，字典形式，指定某些欄位的顯示格式
             例如 {0: lambda v: format_num(v, 2), 3: format_pct}
    """
    html = ['<table class="table"><thead><tr>']
    for h in headers:
        html.append(f"<th>{h}</th>")
    html.append("</tr></thead><tbody>")

    formats = formats or {}
    for row in rows:
        html.append("<tr>")
        for idx, val in enumerate(row):
            fmt = formats.get(idx)
            if callable(fmt):
                shown = fmt(val)
            else:
                shown = "" if val is None else str(val)
            html.append(f"<td>{shown}</td>")
        html.append("</tr>")
    html.append("</tbody></table>")
    return "".join(html)


def make_plotly_div(div_id, traces, layout):
    """
    產生一段 Plotly HTML。
    div_id : 圖表的 div id
    traces : Plotly traces 資料
    layout : Plotly 版面設定
    """
    return f"""
<div id="{div_id}" class="chart"></div>
<script>
Plotly.newPlot({json.dumps(div_id)}, {json.dumps(traces, ensure_ascii=False)}, {json.dumps(layout, ensure_ascii=False)}, {{responsive:true}});
</script>
"""


def main():
    """
    主程式流程：
    1. 開啟 Exp6.xlsx
    2. 讀出五個工作表 A~E
    3. 取出各區塊需要的數據
    4. 產生摘要、表格、互動圖表、圖片
    5. 輸出 exp6_report.html
    """
    xlsx_path = Path("Exp6.xlsx")            # 輸入的 Excel 檔案
    out_html = Path("exp6_report.html")      # 輸出的 HTML 檔案

    # 直接用 zip 方式讀 Excel 內部 XML
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_targets = workbook_sheet_targets(zf)

        # 逐一解析 A~E 工作表
        cells_A = parse_sheet_to_cells(zf, sheet_targets["A"], shared_strings)
        cells_B = parse_sheet_to_cells(zf, sheet_targets["B"], shared_strings)
        cells_C = parse_sheet_to_cells(zf, sheet_targets["C"], shared_strings)
        cells_D = parse_sheet_to_cells(zf, sheet_targets["D"], shared_strings)
        cells_E = parse_sheet_to_cells(zf, sheet_targets["E"], shared_strings)

        # 解析 E 表中的內嵌圖片（波形圖）
        images_E = parse_images_for_sheet(zf, sheet_targets["E"])

    # =====================================================
    # A 工作表：基本參數與 NDAD
    # =====================================================
    radius = cell_num(cells_A, "A2")         # 場線圈平均半徑 R
    turns_field = cell_num(cells_A, "C2")    # 場線圈匝數 NF
    series_r = cell_num(cells_A, "D2")       # 串聯電阻 r
    mu0 = cell_num(cells_A, "F2")            # 真空磁導率 μ0
    ndad_400 = cell_num(cells_A, "D3")       # 400 匝線圈 NDAD
    ndad_2000 = cell_num(cells_A, "D4")      # 2000 匝線圈 NDAD
    ndad_ratio = (ndad_2000 / ndad_400) if ndad_400 else None  # 比值

    # =====================================================
    # B 工作表：頻率與感應電動勢關係
    # =====================================================
    freq_khz = [cell_num(cells_B, ref) for ref in ["C7", "D7", "E7", "F7", "G7", "H7"]]
    exp_400 = [cell_num(cells_B, ref) for ref in ["C8", "D8", "E8", "F8", "G8", "H8"]]
    th_400  = [cell_num(cells_B, ref) for ref in ["C9", "D9", "E9", "F9", "G9", "H9"]]
    err_400 = [cell_num(cells_B, ref) for ref in ["C10", "D10", "E10", "F10", "G10", "H10"]]

    exp_2000 = [cell_num(cells_B, ref) for ref in ["C14", "D14", "E14", "F14", "G14", "H14"]]
    th_2000  = [cell_num(cells_B, ref) for ref in ["C15", "D15", "E15", "F15", "G15", "H15"]]
    err_2000 = [cell_num(cells_B, ref) for ref in ["C16", "D16", "E16", "F16", "G16", "H16"]]

    # 計算平均誤差
    avg_err_400 = sum(x for x in err_400 if x is not None) / len([x for x in err_400 if x is not None])
    avg_err_2000 = sum(x for x in err_2000 if x is not None) / len([x for x in err_2000 if x is not None])

    # =====================================================
    # C 工作表：距離與感應電動勢關係
    # =====================================================
    dist_cm = [cell_num(cells_C, ref) for ref in ["C7", "D7", "E7", "F7", "G7", "H7"]]
    exp_dist = [cell_num(cells_C, ref) for ref in ["C8", "D8", "E8", "F8", "G8", "H8"]]
    th_dist  = [cell_num(cells_C, ref) for ref in ["C9", "D9", "E9", "F9", "G9", "H9"]]
    err_dist = [cell_num(cells_C, ref) for ref in ["C10", "D10", "E10", "F10", "G10", "H10"]]
    avg_err_dist = sum(x for x in err_dist if x is not None) / len([x for x in err_dist if x is not None])

    # =====================================================
    # D 工作表：角度與感應電動勢關係
    # =====================================================
    theta_deg = [cell_num(cells_D, ref) for ref in ["C7", "D7", "E7", "F7", "G7", "H7", "I7"]]
    cos_theta = [cell_num(cells_D, ref) for ref in ["C8", "D8", "E8", "F8", "G8", "H8", "I8"]]
    exp_ang   = [cell_num(cells_D, ref) for ref in ["C9", "D9", "E9", "F9", "G9", "H9", "I9"]]
    th_ang    = [cell_num(cells_D, ref) for ref in ["C10", "D10", "E10", "F10", "G10", "H10", "I10"]]
    err_ang   = [cell_num(cells_D, ref) for ref in ["C11", "D11", "E11", "F11", "G11", "H11", "I11"]]
    valid_err_ang = [x for x in err_ang if x is not None]
    avg_err_ang = sum(valid_err_ang) / len(valid_err_ang)

    # =====================================================
    # 摘要資料列
    # =====================================================
    summary_rows = [
        ["場線圈平均半徑 R (m)", radius],
        ["場線圈匝數 NF", turns_field],
        ["串聯電阻 r (Ω)", series_r],
        ["磁導率 μ0 (T·m/A)", mu0],
        ["NDAD (400匝)", ndad_400],
        ["NDAD (2000匝)", ndad_2000],
        ["NDAD 比值 (2000/400)", ndad_ratio],
        ["平均誤差：頻率實驗 400匝", avg_err_400],
        ["平均誤差：頻率實驗 2000匝", avg_err_2000],
        ["平均誤差：距離實驗", avg_err_dist],
        ["平均誤差：角度實驗", avg_err_ang],
    ]

    # HTML 中「程式目的」說明區塊
    purpose_html = """
    <p>本程式用來將本組法拉第實驗資料 <code>Exp6.xlsx</code> 自動整理為一份可直接在瀏覽器中閱讀的 HTML 報告。程式不依賴 openpyxl 讀取工作表內容，而是直接解析 xlsx 內部 XML，因此即使原始活頁簿含有圖表或部分非標準 XML，也能穩定讀取主要數據。</p>
    <p>輸出的報告包含四個核心區塊：A 表的 NDAD 與基本參數摘要、B 表的頻率與感應電動勢分析、C 表的距離與感應電動勢分析、D 表的角度與 cosθ 關係分析，並額外整理 E 表中的內嵌波形圖片。</p>
    <p>互動圖表使用 Plotly 製作，可在瀏覽器中縮放、平移、滑鼠懸停查看數值，有助於更清楚比較實驗值與理論值之間的差異。</p>
    """

    # 摘要表轉成 HTML
    summary_html = "<table class='table summary'><tbody>" + "".join(
        f"<tr><th>{k}</th><td>{format_num(v, 6) if isinstance(v, (float, int)) else v}</td></tr>"
        for k, v in summary_rows
    ) + "</tbody></table>"

    # =====================================================
    # 各區塊表格資料
    # =====================================================
    table_b_rows = [
        [freq_khz[i], exp_400[i], th_400[i], err_400[i], exp_2000[i], th_2000[i], err_2000[i]]
        for i in range(len(freq_khz))
    ]
    table_c_rows = [
        [dist_cm[i], exp_dist[i], th_dist[i], err_dist[i]]
        for i in range(len(dist_cm))
    ]
    table_d_rows = [
        [theta_deg[i], cos_theta[i], exp_ang[i], th_ang[i], err_ang[i]]
        for i in range(len(theta_deg))
    ]

    # =====================================================
    # 產生 Plotly 圖表
    # =====================================================
    plot_b = make_plotly_div(
        "plot_b",
        [
            {"x": freq_khz, "y": exp_400, "type": "scatter", "mode": "lines+markers", "name": "400匝 實驗值"},
            {"x": freq_khz, "y": th_400, "type": "scatter", "mode": "lines+markers", "name": "400匝 理論值"},
            {"x": freq_khz, "y": exp_2000, "type": "scatter", "mode": "lines+markers", "name": "2000匝 實驗值"},
            {"x": freq_khz, "y": th_2000, "type": "scatter", "mode": "lines+markers", "name": "2000匝 理論值"},
        ],
        {
            "title": "實驗B：頻率 f 與感應電動勢 |ε<sub>emf</sub>|",
            "xaxis": {"title": "頻率 f (kHz)"},
            "yaxis": {"title": "感應電動勢 |ε<sub>emf</sub>| (mV)"},
        },
    )

    plot_c = make_plotly_div(
        "plot_c",
        [
            {"x": dist_cm, "y": exp_dist, "type": "scatter", "mode": "lines+markers", "name": "實驗值"},
            {"x": dist_cm, "y": th_dist, "type": "scatter", "mode": "lines+markers", "name": "理論值"},
        ],
        {
            "title": "實驗C：距離 X 與感應電動勢 |ε<sub>emf</sub>|",
            "xaxis": {"title": "距離 X (cm)"},
            "yaxis": {"title": "感應電動勢 |ε<sub>emf</sub>| (mV)"},
        },
    )

    plot_d = make_plotly_div(
        "plot_d",
        [
            {"x": cos_theta, "y": exp_ang, "type": "scatter", "mode": "lines+markers", "name": "實驗值"},
            {"x": cos_theta, "y": th_ang, "type": "scatter", "mode": "lines+markers", "name": "理論值"},
        ],
        {
            "title": "實驗D：cosθ 與感應電動勢 |ε<sub>emf</sub>|",
            "xaxis": {"title": "cosθ"},
            "yaxis": {"title": "感應電動勢 |ε<sub>emf</sub>| (mV)"},
        },
    )

    # =====================================================
    # 整理 E 工作表中的圖片卡片
    # =====================================================
    image_cards = []
    if images_E:
        labels = ["三角波", "正弦波", "方波"]
        for idx, img in enumerate(images_E, start=1):
            label = labels[idx - 1] if idx <= len(labels) else f"波形圖 {idx}"
            image_cards.append(f"""
            <div class="img-card">
                <div class="img-title">{label}</div>
                <img src="{img['data_uri']}" alt="{label}">
            </div>
            """)

    # =====================================================
    # 組合完整 HTML 頁面
    # =====================================================
    html = f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>法拉第實驗 Python 輔助分析報告</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
body {{
    font-family: Arial, "Microsoft JhengHei", sans-serif;
    margin: 24px;
    line-height: 1.65;
    color: #222;
    background: #fff;
}}
h1, h2, h3 {{
    margin: 0.4em 0;
}}
.small {{ color:#666; font-size:13px; }}
.card {{
    background: #fafafa;
    border: 1px solid #ddd;
    border-radius: 12px;
    padding: 16px 18px;
    margin: 16px 0;
}}
.table {{
    border-collapse: collapse;
    width: 100%;
    font-size: 14px;
    background: #fff;
}}
.table th, .table td {{
    border-bottom: 1px solid #e5e5e5;
    padding: 8px 10px;
    text-align: right;
}}
.table th {{
    background: #f3f6fb;
    text-align: center;
}}
.table td:first-child, .table th:first-child {{
    text-align: center;
}}
.summary th {{
    text-align:left;
    width: 42%;
}}
.summary td {{
    text-align:right;
}}
.grid {{
    display:grid;
    grid-template-columns: 1fr 1fr;
    gap: 18px;
}}
.chart {{
    min-height: 430px;
    border: 1px solid #ddd;
    border-radius: 12px;
    background:#fff;
}}
.img-grid {{
    display:grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    gap: 16px;
}}
.img-card {{
    border:1px solid #ddd;
    border-radius: 12px;
    padding: 10px;
    background:#fff;
}}
.img-card img {{
    width:100%;
    height:auto;
    display:block;
    border-radius:8px;
}}
.img-title {{
    font-weight:700;
    margin-bottom:8px;
}}
.code {{
    background:#f6f8fa;
    padding:2px 6px;
    border-radius:6px;
}}
@media (max-width: 960px) {{
    .grid {{ grid-template-columns: 1fr; }}
}}
</style>
</head>
<body>
<h1>法拉第實驗 — Python 程式輔助 HTML 分析報告</h1>
<p class="small">學號：E24144773 姓名：彭靖翔</p>

<div class="card">
  <h2>程式目的</h2>
  {purpose_html}
</div>

<div class="card">
  <h2>摘要</h2>
  {summary_html}
  <p>從結果可見：頻率增加時感應電動勢上升、距離增加時感應電動勢下降、感應電動勢與 <span class="code">cosθ</span> 呈近似線性關係。這些結果與法拉第定律及磁通量公式相符。</p>
</div>

<div class="card">
  <h2>互動圖表</h2>
  <div class="grid">
    {plot_b}
    {plot_c}
  </div>
  <div style="margin-top:18px;">
    {plot_d}
  </div>
</div>

<div class="card">
  <h2>實驗B：頻率分析資料表</h2>
  {table_html(
      ["頻率 f (kHz)", "400匝 實驗值", "400匝 理論值", "400匝 誤差", "2000匝 實驗值", "2000匝 理論值", "2000匝 誤差"],
      table_b_rows,
      formats={0: lambda v: format_num(v, 2), 1: lambda v: format_num(v, 3), 2: lambda v: format_num(v, 3), 3: format_pct, 4: lambda v: format_num(v, 3), 5: lambda v: format_num(v, 3), 6: format_pct}
  )}
</div>

<div class="card">
  <h2>實驗C：距離分析資料表</h2>
  {table_html(
      ["距離 X (cm)", "實驗值", "理論值", "百分誤差"],
      table_c_rows,
      formats={0: lambda v: format_num(v, 0), 1: lambda v: format_num(v, 3), 2: lambda v: format_num(v, 3), 3: format_pct}
  )}
</div>

<div class="card">
  <h2>實驗D：角度分析資料表</h2>
  {table_html(
      ["θ (°)", "cosθ", "實驗值", "理論值", "百分誤差"],
      table_d_rows,
      formats={0: lambda v: format_num(v, 0), 1: lambda v: format_num(v, 6), 2: lambda v: format_num(v, 3), 3: lambda v: format_num(v, 3), 4: format_pct}
  )}
</div>

<div class="card">
  <h2>實驗E：波形圖片</h2>
  <p>以下圖片由活頁簿內嵌圖檔自動抽出並嵌入 HTML，可直接在瀏覽器查看三角波、正弦波與方波的輸入/輸出波形。</p>
  <div class="img-grid">
    {"".join(image_cards) if image_cards else "<p>此工作表未找到內嵌圖片。</p>"}
  </div>
</div>

</body>
</html>
"""

    # 將 HTML 寫入檔案
    out_html.write_text(html, encoding="utf-8")
    print(f"已輸出: {out_html.resolve()}")


# =========================================================
# 程式進入點
# =========================================================
if __name__ == "__main__":
    main()
