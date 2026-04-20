# -*- coding: utf-8 -*-
# ==========================================
# 實驗八：磁滯曲線 → Excel 轉 HTML 分析報告
# 本程式功能：
# 1. 直接解析 exp8.xlsx（不使用 openpyxl）
# 2. 擷取電流 I、角度 θ、tanθ
# 3. 自動找交點（x軸→Hc，y軸→Br）
# 4. 產生互動式 Plotly 圖表
# 5. 輸出 exp8_report.html
# ==========================================

from pathlib import Path          # 處理檔案路徑
import json                       # 將資料轉為 JSON（Plotly 使用）
import re                         # 正規表達式（解析儲存格位置）
import zipfile                    # xlsx 本質為 zip 檔
import xml.etree.ElementTree as ET  # 解析 Excel XML 結構

import numpy as np                # 數值計算
import pandas as pd               # 表格資料處理

# =========================
# 檔案路徑設定
# =========================
base = Path('.')                  # 目前資料夾
xlsx_path = base / 'exp8.xlsx'    # 輸入 Excel
html_path = base / 'exp8_report.html'  # 輸出 HTML

# Excel XML namespace（固定）
NS = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL_NS = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}

# =========================
# 欄位字母 → 數字（A=1）
# =========================
def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

# =========================
# 讀取 Excel 文字池（sharedStrings）
# =========================
def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    shared = []
    for si in root.findall('a:si', NS):
        texts = []
        for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
            texts.append(t.text or '')
        shared.append(''.join(texts))
    return shared

# =========================
# 取得工作表名稱與檔案位置
# =========================
def get_sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
    rel_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    rel_map = {}
     # 建立 ID → 檔案 mapping
    for rel in rel_root.findall('r:Relationship', REL_NS):
        rid = rel.attrib['Id']
        target = rel.attrib['Target']
        if not target.startswith('xl/'):
            target = 'xl/' + target
        rel_map[rid] = target
    sheets = {}
    for s in wb_root.findall('a:sheets/a:sheet', NS):
        name = s.attrib['name']
        rid = s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        sheets[name] = rel_map[rid]
    return sheets

# =========================
# 解析單一儲存格
# =========================
def cell_value(cell, shared_strings: list[str]):
    t = cell.attrib.get('t')
    if t == 's':
        v = cell.find('a:v', NS)
        return shared_strings[int(v.text)] if v is not None else ''
    if t == 'inlineStr':
        node = cell.find('a:is', NS)
        if node is None:
            return ''
        texts = []
        for tnode in node.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
            texts.append(tnode.text or '')
        return ''.join(texts)
    v = cell.find('a:v', NS)
    return v.text if v is not None else ''

# =========================
# 將整張 Excel sheet 轉為 DataFrame
# =========================
def read_sheet_matrix(zf: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> pd.DataFrame:
    root = ET.fromstring(zf.read(sheet_path))
    rows = []
    for row in root.findall('.//a:sheetData/a:row', NS):
        r = int(row.attrib['r'])
        row_dict = {'__row__': r}
        for c in row.findall('a:c', NS):
            ref = c.attrib['r']
            m = re.match(r'([A-Z]+)(\d+)', ref)
            if not m:
                continue
            col_num = col_to_num(m.group(1))
            row_dict[col_num] = cell_value(c, shared_strings)
        rows.append(row_dict)
    if not rows:
        return pd.DataFrame()
    max_col = 0
    for row in rows:
        ints = [k for k in row if isinstance(k, int)]
        if ints:
            max_col = max(max_col, max(ints))
    data, idx = [], []
    for row in rows:
        idx.append(row['__row__'])
        data.append([row.get(c, '') for c in range(1, max_col + 1)])
    df = pd.DataFrame(data, index=idx, columns=[f'C{c}' for c in range(1, max_col + 1)])
    df.index.name = 'excel_row'
    return df

# =========================
# 安全轉 float（避免錯誤）
# =========================
def to_float(x):
    if x is None:
        return np.nan
    s = str(x).strip()
    if s == '':
        return np.nan
    try:
        return float(s)
    except Exception:
        return np.nan

# =========================
# DataFrame → HTML 表格
# =========================
def df_to_html(df: pd.DataFrame) -> str:
    show = df.copy()
    for col in show.columns:
        if pd.api.types.is_numeric_dtype(show[col]):
            show[col] = show[col].map(lambda x: f'{x:.6f}' if pd.notna(x) else '')
    return show.to_html(index=False, border=0, classes='table', justify='center')

# =========================
# 找交點（磁滯核心🔥）
# x軸交點 → tanθ=0 → Hc
# y軸交點 → I=0 → Br
# =========================
def find_axis_crossings(df: pd.DataFrame, path_label: str):
    pts = []
    x = df['I_A'].to_numpy(float)
    y = df['tan_theta'].to_numpy(float)
    for i in range(len(df) - 1):
        x1, y1 = x[i], y[i]
        x2, y2 = x[i + 1], y[i + 1]
        if np.isfinite(y1) and np.isfinite(y2):
            if y1 == 0:
                pts.append({'x': x1, 'y': 0.0, 'kind': 'x軸交點', 'path': path_label})
            elif y1 * y2 < 0 and y2 != y1:
                t = -y1 / (y2 - y1)
                xc = x1 + t * (x2 - x1)
                pts.append({'x': xc, 'y': 0.0, 'kind': 'x軸交點', 'path': path_label})
        if np.isfinite(x1) and np.isfinite(x2):
            if x1 == 0:
                pts.append({'x': 0.0, 'y': y1, 'kind': 'y軸交點', 'path': path_label})
            elif x1 * x2 < 0 and x2 != x1:
                t = -x1 / (x2 - x1)
                yc = y1 + t * (y2 - y1)
                pts.append({'x': 0.0, 'y': yc, 'kind': 'y軸交點', 'path': path_label})
    dedup = []
    for p in pts:
        if not any(abs(p['x'] - q['x']) < 1e-6 and abs(p['y'] - q['y']) < 1e-6 and p['kind'] == q['kind'] for q in dedup):
            dedup.append(p)
    return dedup


zf = zipfile.ZipFile(xlsx_path)
shared_strings = load_shared_strings(zf)
sheet_paths = get_sheet_paths(zf)
df_s1 = read_sheet_matrix(zf, sheet_paths['磁滯曲線實驗1'], shared_strings)
df_s2 = read_sheet_matrix(zf, sheet_paths['磁滯曲線實驗2'], shared_strings)

segment_defs = [
    ('A', '0.05A → 飽和', 2, 3, 4),
    ('B', '飽和 → 0A', 6, 7, 8),
    ('C', '-0.05A → -飽和', 10, 11, 12),
    ('D', '-飽和 → 0A', 14, 15, 16),
    ('E', '0.05A → 飽和', 18, 19, 20),
]

exp1_tables = {}
for code_name, display_name, cI, cTheta, cTan in segment_defs:
    sub = pd.DataFrame({
        'I_A': df_s1[f'C{cI}'].map(to_float),
        'theta_deg': df_s1[f'C{cTheta}'].map(to_float),
        'tan_theta': df_s1[f'C{cTan}'].map(to_float),
    }).dropna(how='all')
    sub = sub.dropna(subset=['I_A', 'tan_theta'], how='any').reset_index(drop=True)
    exp1_tables[f'{code_name}_{display_name}'] = sub

loop_order = [
    'A_0.05A → 飽和',
    'B_飽和 → 0A',
    'C_-0.05A → -飽和',
    'D_-飽和 → 0A',
    'E_0.05A → 飽和',
]
loop_df = pd.concat([exp1_tables[name].assign(path=name.split('_')[0]) for name in loop_order], ignore_index=True)

exp2_defs = [
    ('B', '表格B：飽和 → 0A', 1, 2),
    ('C', '表格C：-0.05A → -飽和', 3, 4),
    ('D', '表格D：飽和 → 0A', 5, 6),
    ('E', '表格E：0.05A → 飽和', 7, 8),
]
exp2_tables = {}
for key, title, cI, cTan in exp2_defs:
    sub = pd.DataFrame({
        'I_A': df_s2[f'C{cI}'].map(to_float),
        'tan_theta': df_s2[f'C{cTan}'].map(to_float),
    }).dropna(how='all')
    sub = sub.dropna(subset=['I_A', 'tan_theta'], how='any').reset_index(drop=True)
    exp2_tables[f'{key}_{title}'] = sub

hc_left = to_float(df_s2.loc[6, 'C5'])
hc_right = to_float(df_s2.loc[6, 'C8'])
br_left = to_float(df_s2.loc[7, 'C5'])
br_right = to_float(df_s2.loc[7, 'C8'])

summary_rows = [
    ('實驗1總資料點數', f'{len(loop_df)}'),
    ('實驗1最大 tanθ', f"{loop_df['tan_theta'].max():.4f}"),
    ('實驗1最小 tanθ', f"{loop_df['tan_theta'].min():.4f}"),
    ('實驗1最大電流 I (A)', f"{loop_df['I_A'].max():.4f}"),
    ('實驗1最小電流 I (A)', f"{loop_df['I_A'].min():.4f}"),
    ('左側抗磁力 Hc (N/L)', f'{hc_left:.6f}' if pd.notna(hc_left) else '—'),
    ('右側抗磁力 Hc (N/L)', f'{hc_right:.6f}' if pd.notna(hc_right) else '—'),
    ('左側殘磁量 Br (αμ0)', f'{br_left:.4f}' if pd.notna(br_left) else '—'),
    ('右側殘磁量 Br (αμ0)', f'{br_right:.4f}' if pd.notna(br_right) else '—'),
]
summary_html = "<table class='summary'>" + ''.join(f'<tr><th>{k}</th><td>{v}</td></tr>' for k, v in summary_rows) + '</table>'

purpose_html = """
<p>本程式用於將實驗八「磁滯曲線」之 <code>exp8.xlsx</code> 自動整理為可於瀏覽器檢視的 HTML 報告。透過 <code>zipfile + XML</code> 解析資料並繪製圖形，特別採用滑鼠互動式介面，使用者可將滑鼠移至資料點即時顯示數值與位置資訊，提升判讀效率與準確性。</p>"""

plotly_cdn = 'https://cdn.plot.ly/plotly-2.32.0.min.js'

fig1_traces = []
axis_crossings = []
for name in loop_order:
    code = name.split('_')[0]
    dfp = exp1_tables[name]
    fig1_traces.append({
        'x': dfp['I_A'].tolist(),
        'y': dfp['tan_theta'].tolist(),
        'type': 'scatter',
        'mode': 'lines+markers',
        'name': f'{code}段',
        'line': {'width': 3},
        'marker': {'size': 8},
    })
    axis_crossings.extend(find_axis_crossings(dfp, code))

x_cross = [p for p in axis_crossings if p['kind'] == 'x軸交點']
y_cross = [p for p in axis_crossings if p['kind'] == 'y軸交點']
if x_cross:
    fig1_traces.append({
        'x': [p['x'] for p in x_cross],
        'y': [p['y'] for p in x_cross],
        'type': 'scatter',
        'mode': 'markers+text',
        'name': 'x軸交點',
        'text': [f"{p['path']}段\\n({p['x']:.4f}, 0)" for p in x_cross],
        'textposition': 'top center',
        'marker': {'size': 13, 'symbol': 'diamond', 'line': {'width': 2, 'color': '#8a2be2'}},
    })
if y_cross:
    fig1_traces.append({
        'x': [p['x'] for p in y_cross],
        'y': [p['y'] for p in y_cross],
        'type': 'scatter',
        'mode': 'markers+text',
        'name': 'y軸交點',
        'text': [f"{p['path']}段\\n(0, {p['y']:.4f})" for p in y_cross],
        'textposition': 'bottom center',
        'marker': {'size': 14, 'symbol': 'circle-open', 'line': {'width': 3, 'color': '#d9485f'}},
    })

fig1_shapes = []
fig1_annotations = []
for p in x_cross:
    fig1_shapes.append({'type': 'line', 'xref': 'x', 'yref': 'paper', 'x0': p['x'], 'x1': p['x'], 'y0': 0, 'y1': 1, 'line': {'dash': 'dot', 'width': 1.5, 'color': '#8a2be2'}})
    fig1_annotations.append({'x': p['x'], 'y': 0, 'xref': 'x', 'yref': 'y', 'text': f"x交點={p['x']:.4f}", 'showarrow': True, 'arrowhead': 2, 'ax': 0, 'ay': -28, 'font': {'color': '#8a2be2', 'size': 12}})
for p in y_cross:
    fig1_shapes.append({'type': 'line', 'xref': 'paper', 'yref': 'y', 'x0': 0, 'x1': 1, 'y0': p['y'], 'y1': p['y'], 'line': {'dash': 'dot', 'width': 1.5, 'color': '#d9485f'}})
    fig1_annotations.append({'x': 0, 'y': p['y'], 'xref': 'x', 'yref': 'y', 'text': f"y交點={p['y']:.4f}", 'showarrow': True, 'arrowhead': 2, 'ax': 52, 'ay': 0, 'font': {'color': '#d9485f', 'size': 12}})

fig1_layout = {
    'title': '圖1：磁滯曲線實驗1各路徑 I 與 tanθ 的關係（含交點）',
    'xaxis': {'title': '電流 I (A)', 'zeroline': True, 'zerolinewidth': 2, 'gridcolor': '#d9e2f2'},
    'yaxis': {'title': 'tanθ', 'zeroline': True, 'zerolinewidth': 2, 'gridcolor': '#d9e2f2'},
    'paper_bgcolor': 'rgba(0,0,0,0)',
    'plot_bgcolor': '#ffffff',
    'legend': {'orientation': 'h', 'y': 1.12, 'x': 0},
    'margin': {'l': 60, 'r': 30, 't': 80, 'b': 60},
    'shapes': fig1_shapes,
    'annotations': fig1_annotations,
}

fig3_traces = []
for key in exp2_tables:
    dfp = exp2_tables[key]
    label = key.split('_')[0]
    fig3_traces.append({
        'x': dfp['I_A'].tolist(),
        'y': dfp['tan_theta'].tolist(),
        'type': 'scatter',
        'mode': 'lines+markers',
        'name': label,
        'line': {'width': 3},
        'marker': {'size': 8},
    })

shapes = []
annotations = []
if pd.notna(hc_left):
    shapes.append({'type': 'line', 'xref': 'x', 'yref': 'paper', 'x0': hc_left, 'x1': hc_left, 'y0': 0, 'y1': 1, 'line': {'dash': 'dash', 'width': 2, 'color': '#7c4dff'}})
    annotations.append({'x': hc_left, 'y': 1, 'xref': 'x', 'yref': 'paper', 'text': f'Hc左={hc_left:.4f}', 'showarrow': False, 'yshift': 12, 'font': {'color': '#7c4dff'}})
if pd.notna(hc_right):
    shapes.append({'type': 'line', 'xref': 'x', 'yref': 'paper', 'x0': hc_right, 'x1': hc_right, 'y0': 0, 'y1': 1, 'line': {'dash': 'dash', 'width': 2, 'color': '#7c4dff'}})
    annotations.append({'x': hc_right, 'y': 1, 'xref': 'x', 'yref': 'paper', 'text': f'Hc右={hc_right:.4f}', 'showarrow': False, 'yshift': 30, 'font': {'color': '#7c4dff'}})
if pd.notna(br_left):
    shapes.append({'type': 'line', 'xref': 'paper', 'yref': 'y', 'x0': 0, 'x1': 1, 'y0': br_left, 'y1': br_left, 'line': {'dash': 'dot', 'width': 2, 'color': '#ff5c8a'}})
    annotations.append({'x': 1, 'y': br_left, 'xref': 'paper', 'yref': 'y', 'text': f'Br左={br_left:.2f}', 'showarrow': False, 'xshift': 42, 'font': {'color': '#ff5c8a'}})
if pd.notna(br_right):
    shapes.append({'type': 'line', 'xref': 'paper', 'yref': 'y', 'x0': 0, 'x1': 1, 'y0': br_right, 'y1': br_right, 'line': {'dash': 'dot', 'width': 2, 'color': '#ff5c8a'}})
    annotations.append({'x': 1, 'y': br_right, 'xref': 'paper', 'yref': 'y', 'text': f'Br右={br_right:.2f}', 'showarrow': False, 'xshift': 42, 'font': {'color': '#ff5c8a'}})

fig3_layout = {
    'title': '圖2：實驗2過零附近資料與 Hc / Br 位置',
    'xaxis': {'title': '電流 I (A)', 'zeroline': True, 'zerolinewidth': 2, 'gridcolor': '#d9e2f2'},
    'yaxis': {'title': 'tanθ', 'zeroline': True, 'zerolinewidth': 2, 'gridcolor': '#d9e2f2'},
    'paper_bgcolor': 'rgba(0,0,0,0)',
    'plot_bgcolor': '#ffffff',
    'legend': {'orientation': 'h', 'y': 1.12, 'x': 0},
    'margin': {'l': 60, 'r': 40, 't': 80, 'b': 60},
    'shapes': shapes,
    'annotations': annotations,
}

html_template = """<!doctype html>
<html lang='zh-Hant'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>實驗八：磁滯曲線 HTML 分析報告</title>
<script src='__PLOTLY_CDN__'></script>
<style>
:root {
    --bg1:#f4f7ff;
    --bg2:#eef8ff;
    --card:#ffffffcc;
    --line:#d9e3f0;
    --text:#1f2937;
    --muted:#5f6b7a;
    --title:#19335a;
    --accent:#4f7cff;
    --accent2:#7c4dff;
}
*{box-sizing:border-box}
body {
    font-family: "Segoe UI", "Microsoft JhengHei", Arial, sans-serif;
    margin: 0;
    line-height: 1.75;
    color: var(--text);
    background: linear-gradient(135deg, var(--bg1), var(--bg2));
}
.container {
    max-width: 1280px;
    margin: 0 auto;
    padding: 32px 22px 60px;
}
.hero {
    background: linear-gradient(135deg, #1d4ed8, #6d28d9 65%, #0ea5e9);
    color: #fff;
    border-radius: 28px;
    padding: 30px 32px;
    box-shadow: 0 20px 50px rgba(49, 75, 132, .18);
    margin-bottom: 24px;
}
.hero h1 { margin: 0 0 8px; font-size: 34px; letter-spacing: .5px; }
.hero p { margin: 4px 0; font-size: 18px; opacity: .96; }
.grid {
    display:grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
}
.card {
    background: var(--card);
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255,255,255,.7);
    border-radius: 24px;
    padding: 22px 24px;
    margin: 18px 0;
    box-shadow: 0 12px 30px rgba(74, 93, 143, .10);
}
.card h2 {
    margin: 0 0 12px;
    color: var(--title);
    font-size: 24px;
}
.card h3 {
    margin: 22px 0 10px;
    color: #23406b;
    font-size: 18px;
}
.badge {
    display:inline-block;
    background: #e9f0ff;
    color: var(--accent2);
    border: 1px solid #d9e6ff;
    padding: 5px 10px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 700;
    margin-bottom: 10px;
}
.table, .summary {
    width:100%;
    border-collapse: collapse;
    overflow:hidden;
    background:#fff;
    border-radius: 18px;
}
.table th, .table td, .summary th, .summary td {
    border-bottom: 1px solid #ebf0f7;
    padding: 10px 12px;
}
.table th, .summary th {
    background: linear-gradient(135deg, #edf3ff, #f7faff);
    color:#24446f;
}
.table td { text-align:right; }
.table th { text-align:center; }
.table td:first-child, .table th:first-child { text-align:center; }
.summary th { text-align:left; width:46%; }
.summary td { text-align:right; }
.chart {
    min-height: 480px;
    border-radius: 20px;
    background: linear-gradient(180deg, #ffffff, #fbfdff);
    border: 1px solid var(--line);
    box-shadow: inset 0 0 0 1px rgba(255,255,255,.5);
}
.note {
    color: var(--muted);
    font-size: 14px;
}
@media (max-width: 960px) {
    .grid { grid-template-columns: 1fr; }
    .hero h1 { font-size: 28px; }
}
</style>
</head>
<body>
<div class='container'>
  <section class='hero'>
    <h1>實驗八：磁滯曲線 — HTML 分析報告</h1>
   <p style="font-size: 24px;"><strong>學號：</strong>E24144773　<strong>姓名：</strong>彭靖翔</p>
  </section>

  <div class='card'>
    <div class='badge' style="font-size: 28px;">程式目的</div>
    __PURPOSE_HTML__
  </div>

  <div class='card'>
    <div class='badge' style="font-size: 30px;">摘要</div>
    __SUMMARY_HTML__
    <p>由圖形可觀察到量測資料在磁化與退磁過程中並不重合，形成典型的磁滯回線；此外，實驗2所整理的過零附近資料可用來估計抗磁力 Hc，而 I=0 時的 tanθ 可對應殘磁量 Br 的概念。</p>
  </div>

  <div class='card'>
    <h2>互動圖表</h2>
    <div class='grid'>
      <div id='plot1' class='chart'></div>
      <div id='plot3' class='chart'></div>
    </div>
  </div>

  <div class='card'>
    <h2>資料表</h2>
    <h3>表1：磁滯曲線實驗1－A段</h3>
    __TABLE_A__
    <h3>表2：磁滯曲線實驗1－B段</h3>
    __TABLE_B__
    <h3>表3：磁滯曲線實驗1－C段</h3>
    __TABLE_C__
    <h3>表4：磁滯曲線實驗1－D段</h3>
    __TABLE_D__
    <h3>表5：磁滯曲線實驗1－E段</h3>
    __TABLE_E__
    <h3>表6：磁滯曲線實驗2整理</h3>
    __TABLE_EXP2_ALL__
  </div>
</div>
<script>
Plotly.newPlot('plot1', __FIG1_TRACES__, __FIG1_LAYOUT__, {responsive:true, displaylogo:false});
Plotly.newPlot('plot3', __FIG3_TRACES__, __FIG3_LAYOUT__, {responsive:true, displaylogo:false});
</script>
</body>
</html>
"""

exp2_all_df = pd.concat([exp2_tables[k].assign(section=k.split('_')[0]) for k in exp2_tables], ignore_index=True)
html = (
    html_template
    .replace('__PLOTLY_CDN__', plotly_cdn)
    .replace('__PURPOSE_HTML__', purpose_html)
    .replace('__SUMMARY_HTML__', summary_html)
    .replace('__TABLE_A__', df_to_html(exp1_tables['A_0.05A → 飽和']))
    .replace('__TABLE_B__', df_to_html(exp1_tables['B_飽和 → 0A']))
    .replace('__TABLE_C__', df_to_html(exp1_tables['C_-0.05A → -飽和']))
    .replace('__TABLE_D__', df_to_html(exp1_tables['D_-飽和 → 0A']))
    .replace('__TABLE_E__', df_to_html(exp1_tables['E_0.05A → 飽和']))
    .replace('__TABLE_EXP2_ALL__', df_to_html(exp2_all_df))
    .replace('__FIG1_TRACES__', json.dumps(fig1_traces, ensure_ascii=False))
    .replace('__FIG1_LAYOUT__', json.dumps(fig1_layout, ensure_ascii=False))
    .replace('__FIG3_TRACES__', json.dumps(fig3_traces, ensure_ascii=False))
    .replace('__FIG3_LAYOUT__', json.dumps(fig3_layout, ensure_ascii=False))
)
html_path.write_text(html, encoding='utf-8')
print(f'已輸出: {html_path.resolve()}')
