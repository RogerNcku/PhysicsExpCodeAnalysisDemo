# -*- coding: utf-8 -*-
# =========================================================
# 實驗九：折射率、色散、光譜儀 → Excel 轉互動式 HTML 報告
# 功能：
# 1. 只讀取一次 exp9.xlsx（xlsx 以 zipfile + XML 解析，不依賴 openpyxl）
# 2. 自動擷取 Sellmeier 係數、8 條汞燈光譜線、5 次量測值、平均 n、標準差、理論值、誤差
# 3. 產生更漂亮、多色、滑鼠提示字體較大的 Plotly 互動式 HTML
# 4. 輸出 exp9_report.html
# =========================================================

from pathlib import Path
import json
import math
import re
import statistics
import zipfile
import xml.etree.ElementTree as ET

BASE = Path(__file__).resolve().parent
XLSX_PATH = BASE / "exp9.xlsx"
HTML_PATH = BASE / "exp9_report.html"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

COLORS = {
    "red": "#ff4d6d",
    "orange": "#ff9f1c",
    "yellow": "#ffd166",
    "green": "#06d6a0",
    "cyan": "#00b4d8",
    "blue": "#4d96ff",
    "purple": "#7b2ff7",
    "pink": "#f72585",
    "dark": "#14213d",
}

LINE_COLORS = [
    "#ff4d6d", "#ff9f1c", "#f4d35e", "#06d6a0",
    "#00b4d8", "#4d96ff", "#7b2ff7", "#f72585"
]

# ---------------------------------------------------------
# Excel XML 基本工具
# ---------------------------------------------------------
def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    shared = []
    for si in root.findall("a:si", NS):
        texts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            texts.append(t.text or "")
        shared.append("".join(texts))
    return shared


def get_sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {}
    for rel in rel_root.findall("r:Relationship", REL_NS):
        rid = rel.attrib["Id"]
        target = rel.attrib["Target"]
        if not target.startswith("xl/"):
            target = "xl/" + target
        rel_map[rid] = target
    sheets = {}
    for s in wb_root.findall("a:sheets/a:sheet", NS):
        name = s.attrib["name"]
        rid = s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        sheets[name] = rel_map[rid]
    return sheets


def cell_value(cell, shared_strings: list[str]):
    t = cell.attrib.get("t")
    if t == "s":
        v = cell.find("a:v", NS)
        return shared_strings[int(v.text)] if v is not None and v.text else ""
    if t == "inlineStr":
        node = cell.find("a:is", NS)
        if node is None:
            return ""
        return "".join(t.text or "" for t in node.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
    v = cell.find("a:v", NS)
    return v.text if v is not None else ""


def read_sheet_matrix(zf: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> dict[tuple[int, int], str]:
    root = ET.fromstring(zf.read(sheet_path))
    cells = {}
    for row in root.findall(".//a:sheetData/a:row", NS):
        r = int(row.attrib["r"])
        for c in row.findall("a:c", NS):
            ref = c.attrib.get("r", "")
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if not m:
                continue
            col = col_to_num(m.group(1))
            cells[(r, col)] = cell_value(c, shared_strings)
    return cells


def to_float(x):
    if x is None:
        return None
    s = str(x).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def fmt(x, digits=4):
    if x is None:
        return "—"
    try:
        return f"{float(x):.{digits}f}"
    except Exception:
        return str(x)


def html_escape(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# ---------------------------------------------------------
# 實驗資料解析
# ---------------------------------------------------------
def parse_exp9_xlsx(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = load_shared_strings(zf)
        sheet_paths = get_sheet_paths(zf)

        # 主要資料使用「編號 2~8最小偏向角」，此工作表包含編號1~8與右側總表。
        main_sheet_name = "編號 2~8最小偏向角"
        if main_sheet_name not in sheet_paths:
            # 若日後工作表名稱不同，改找含「最小偏向角」的工作表。
            candidates = [name for name in sheet_paths if "最小偏向角" in name]
            if not candidates:
                raise RuntimeError("找不到最小偏向角資料工作表。")
            main_sheet_name = candidates[-1]

        cells = read_sheet_matrix(zf, sheet_paths[main_sheet_name], shared_strings)

    # Sellmeier 係數通常在 row 3, col B~G 或 A~F，本檔在 B~G。
    coeff = {
        "B1": to_float(cells.get((3, 2))),
        "B2": to_float(cells.get((3, 3))),
        "B3": to_float(cells.get((3, 4))),
        "C1": to_float(cells.get((3, 5))),
        "C2": to_float(cells.get((3, 6))),
        "C3": to_float(cells.get((3, 7))),
    }

    # 自動尋找每個「編號」開頭列。格式：col B = 編號 1, col C = wavelength。
    starts = []
    for (r, c), value in sorted(cells.items()):
        if c == 2 and isinstance(value, str) and value.strip().startswith("編號"):
            label = value.strip()
            number_match = re.search(r"\d+", label)
            number = int(number_match.group()) if number_match else len(starts) + 1
            wavelength = to_float(cells.get((r, 3)))
            if wavelength is not None:
                starts.append((r, number, wavelength))

    records = []
    summary = []
    for start_row, number, wavelength in starts:
        measurements = []
        for rr in range(start_row, start_row + 5):
            d1 = to_float(cells.get((rr, 4)))
            d2 = to_float(cells.get((rr, 5)))
            D = to_float(cells.get((rr, 6)))
            n = to_float(cells.get((rr, 7)))
            if d1 is None or d2 is None or D is None or n is None:
                continue
            row = {
                "line_no": number,
                "lambda_A": wavelength,
                "trial": len(measurements) + 1,
                "d1_deg": d1,
                "d2_deg": d2,
                "D_deg": D,
                "n": n,
            }
            measurements.append(row)
            records.append(row)

        mean_n = to_float(cells.get((start_row + 5, 4)))
        std_n = to_float(cells.get((start_row + 5, 7)))
        theory_n = to_float(cells.get((start_row + 6, 4)))
        error_pct = to_float(cells.get((start_row + 6, 7)))

        # 若 cached formula value 缺少，則用 Python 補算。
        if mean_n is None and measurements:
            mean_n = statistics.mean(m["n"] for m in measurements)
        if std_n is None and len(measurements) > 1:
            std_n = statistics.pstdev(m["n"] for m in measurements)
        if theory_n is None:
            theory_n = sellmeier_n(wavelength, coeff)
        if error_pct is None and mean_n is not None and theory_n is not None:
            error_pct = abs((mean_n - theory_n) / theory_n) * 100

        summary.append({
            "line_no": number,
            "lambda_A": wavelength,
            "lambda_nm": wavelength / 10,
            "mean_D_deg": statistics.mean(m["D_deg"] for m in measurements) if measurements else None,
            "mean_n": mean_n,
            "std_n": std_n,
            "theory_n": theory_n,
            "error_pct": error_pct,
            "count": len(measurements),
        })

    summary.sort(key=lambda x: x["lambda_A"], reverse=True)
    return coeff, summary, records, main_sheet_name


def sellmeier_n(lambda_A: float, coeff: dict) -> float | None:
    if lambda_A is None:
        return None
    try:
        lam_um = lambda_A * 0.0001  # Å → μm
        lam2 = lam_um ** 2
        n2 = 1
        n2 += coeff["B1"] * lam2 / (lam2 - coeff["C1"])
        n2 += coeff["B2"] * lam2 / (lam2 - coeff["C2"])
        n2 += coeff["B3"] * lam2 / (lam2 - coeff["C3"])
        return math.sqrt(n2)
    except Exception:
        return None

# ---------------------------------------------------------
# HTML 表格與圖表資料
# ---------------------------------------------------------
def table_html(headers, rows, numeric_digits=4):
    thead = "<thead><tr>" + "".join(f"<th>{html_escape(h)}</th>" for h in headers) + "</tr></thead>"
    body_rows = []
    for row in rows:
        tds = []
        for item in row:
            if isinstance(item, float):
                value = fmt(item, numeric_digits)
            else:
                value = html_escape(item)
            tds.append(f"<td>{value}</td>")
        body_rows.append("<tr>" + "".join(tds) + "</tr>")
    return "<table class='data-table'>" + thead + "<tbody>" + "".join(body_rows) + "</tbody></table>"


def make_plot_traces(summary, records, coeff):
    wavelengths = [s["lambda_A"] for s in summary]
    mean_n = [s["mean_n"] for s in summary]
    theory_n = [s["theory_n"] for s in summary]
    error_pct = [s["error_pct"] for s in summary]
    std_n = [s["std_n"] for s in summary]
    mean_D = [s["mean_D_deg"] for s in summary]
    line_no = [s["line_no"] for s in summary]

    custom_summary = [[
        s["line_no"], s["lambda_A"], s["lambda_nm"], s["mean_n"], s["theory_n"], s["error_pct"], s["std_n"], s["mean_D_deg"]
    ] for s in summary]

    # 理論 Sellmeier 平滑曲線：從最短到最長波長
    min_lam = min(wavelengths)
    max_lam = max(wavelengths)
    smooth_x = [min_lam + (max_lam - min_lam) * i / 350 for i in range(351)]
    smooth_y = [sellmeier_n(x, coeff) for x in smooth_x]

    fig1 = {
        "traces": [
            {
                "x": smooth_x,
                "y": smooth_y,
                "type": "scatter",
                "mode": "lines",
                "name": "Sellmeier 理論曲線",
                "line": {"width": 5, "color": COLORS["purple"], "shape": "spline"},
                "hovertemplate": "<b>Sellmeier 理論</b><br>λ = %{x:.0f} Å<br>n = %{y:.5f}<extra></extra>",
            },
            {
                "x": wavelengths,
                "y": mean_n,
                "type": "scatter",
                "mode": "markers+lines",
                "name": "實驗平均折射率 n",
                "customdata": custom_summary,
                "line": {"width": 4, "color": COLORS["cyan"], "shape": "spline"},
                "marker": {
                    "size": 16,
                    "color": error_pct,
                    "colorscale": "Turbo",
                    "showscale": True,
                    "colorbar": {"title": "誤差 %"},
                    "line": {"width": 2, "color": "white"},
                },
                "hovertemplate": (
                    "<b>編號 %{customdata[0]}</b><br>"
                    "波長 λ = %{customdata[1]:.0f} Å（%{customdata[2]:.1f} nm）<br>"
                    "實驗 n = %{customdata[3]:.5f}<br>"
                    "理論 n = %{customdata[4]:.5f}<br>"
                    "誤差 = %{customdata[5]:.3f}%<br>"
                    "標準差 = %{customdata[6]:.5f}<br>"
                    "平均最小偏向角 D = %{customdata[7]:.2f}°<extra></extra>"
                ),
            },
        ],
        "layout": base_layout(
            "圖1：λ–n 色散曲線（實驗值 vs. Sellmeier 理論值）",
            "波長 λ (Å)",
            "折射率 n",
            reverse_x=True,
        ),
    }

    fig2 = {
        "traces": [
            {
                "x": [f"編號{s['line_no']}<br>{int(s['lambda_A'])} Å" for s in summary],
                "y": error_pct,
                "type": "bar",
                "name": "誤差 (%)",
                "marker": {"color": error_pct, "colorscale": "Sunsetdark", "line": {"width": 1.5, "color": "white"}},
                "customdata": custom_summary,
                "hovertemplate": (
                    "<b>編號 %{customdata[0]}</b><br>λ = %{customdata[1]:.0f} Å<br>"
                    "誤差 = %{y:.3f}%<br>實驗 n = %{customdata[3]:.5f}<br>理論 n = %{customdata[4]:.5f}<extra></extra>"
                ),
            },
            {
                "x": [f"編號{s['line_no']}<br>{int(s['lambda_A'])} Å" for s in summary],
                "y": std_n,
                "type": "scatter",
                "mode": "lines+markers",
                "name": "標準差",
                "yaxis": "y2",
                "line": {"width": 4, "color": COLORS["green"], "shape": "spline"},
                "marker": {"size": 13, "color": COLORS["green"], "line": {"width": 2, "color": "white"}},
                "hovertemplate": "<b>%{x}</b><br>標準差 = %{y:.5f}<extra></extra>",
            },
        ],
        "layout": with_y2(base_layout("圖2：各光譜線誤差與標準差比較", "光譜線", "誤差 (%)"), "標準差"),
    }

    fig3_traces = []
    for i, s in enumerate(summary):
        pts = [r for r in records if r["line_no"] == s["line_no"]]
        pts.sort(key=lambda r: r["trial"])
        fig3_traces.append({
            "x": [f"編號{s['line_no']}<br>{int(s['lambda_A'])} Å" for _ in pts],
            "y": [r["n"] for r in pts],
            "type": "box",
            "name": f"{int(s['lambda_A'])} Å",
            "boxpoints": "all",
            "jitter": 0.55,
            "pointpos": 0,
            "marker": {"size": 8, "color": LINE_COLORS[i % len(LINE_COLORS)], "opacity": 0.82},
            "line": {"width": 2, "color": LINE_COLORS[i % len(LINE_COLORS)]},
            "customdata": [[r["trial"], r["d1_deg"], r["d2_deg"], r["D_deg"], r["n"]] for r in pts],
            "hovertemplate": (
                "<b>%{x}</b><br>第 %{customdata[0]} 次量測<br>"
                "d1 = %{customdata[1]:.2f}°<br>d2 = %{customdata[2]:.2f}°<br>"
                "D = %{customdata[3]:.2f}°<br>n = %{customdata[4]:.5f}<extra></extra>"
            ),
        })
    fig3 = {
        "traces": fig3_traces,
        "layout": base_layout("圖3：每條光譜線的 5 次折射率量測分布", "光譜線", "折射率 n"),
    }

    fig4 = {
        "traces": [
            {
                "x": wavelengths,
                "y": mean_D,
                "type": "scatter",
                "mode": "markers+lines",
                "name": "平均最小偏向角 D",
                "customdata": custom_summary,
                "line": {"width": 5, "color": COLORS["orange"], "shape": "spline"},
                "marker": {"size": 15, "color": mean_D, "colorscale": "Plasma", "line": {"width": 2, "color": "white"}},
                "hovertemplate": (
                    "<b>編號 %{customdata[0]}</b><br>λ = %{customdata[1]:.0f} Å<br>"
                    "平均 D = %{customdata[7]:.2f}°<br>平均 n = %{customdata[3]:.5f}<extra></extra>"
                ),
            }
        ],
        "layout": base_layout("圖4：波長與平均最小偏向角 D 的關係", "波長 λ (Å)", "平均最小偏向角 D (°)", reverse_x=True),
    }

    return fig1, fig2, fig3, fig4


def base_layout(title, x_title, y_title, reverse_x=False):
    return {
        "title": {"text": title, "font": {"size": 24, "color": "#14213d"}, "x": 0.03},
        "font": {"family": "Microsoft JhengHei, Noto Sans TC, Arial", "size": 16, "color": "#1f2937"},
        "xaxis": {
            "title": {"text": x_title, "font": {"size": 18}},
            "gridcolor": "rgba(100,116,139,.18)",
            "zerolinecolor": "rgba(15,23,42,.3)",
            "autorange": "reversed" if reverse_x else True,
        },
        "yaxis": {
            "title": {"text": y_title, "font": {"size": 18}},
            "gridcolor": "rgba(100,116,139,.18)",
            "zerolinecolor": "rgba(15,23,42,.3)",
        },
        "plot_bgcolor": "rgba(255,255,255,0.92)",
        "paper_bgcolor": "rgba(0,0,0,0)",
        "legend": {"orientation": "h", "y": 1.12, "x": 0, "font": {"size": 15}},
        "hoverlabel": {
            "bgcolor": "rgba(15,23,42,0.94)",
            "bordercolor": "#93c5fd",
            "font": {"family": "Microsoft JhengHei, Noto Sans TC, Arial", "size": 21, "color": "#ffffff"},
            "align": "left",
        },
        "margin": {"l": 72, "r": 48, "t": 100, "b": 76},
    }


def with_y2(layout, y2_title):
    layout = dict(layout)
    layout["yaxis2"] = {
        "title": {"text": y2_title, "font": {"size": 18, "color": COLORS["green"]}},
        "overlaying": "y",
        "side": "right",
        "showgrid": False,
        "tickfont": {"color": COLORS["green"]},
    }
    layout["margin"] = {"l": 72, "r": 86, "t": 100, "b": 76}
    return layout

# ---------------------------------------------------------
# HTML 報告產生
# ---------------------------------------------------------
def build_html(coeff, summary, records, source_sheet_name):
    fig1, fig2, fig3, fig4 = make_plot_traces(summary, records, coeff)

    mean_errors = statistics.mean(s["error_pct"] for s in summary if s["error_pct"] is not None)
    max_error = max(summary, key=lambda s: s["error_pct"] or -1)
    min_error = min(summary, key=lambda s: s["error_pct"] or 999)
    max_n = max(summary, key=lambda s: s["mean_n"] or -1)
    min_n = min(summary, key=lambda s: s["mean_n"] or 999)

    summary_cards = [
        ("資料來源工作表", source_sheet_name, "自動讀取 exp9.xlsx", "#7b2ff7"),
        ("光譜線數量", f"{len(summary)} 條", "每條 5 次量測", "#4d96ff"),
        ("平均誤差", f"{mean_errors:.3f}%", "實驗 n 與理論 n 比較", "#06d6a0"),
        ("最小誤差", f"編號 {min_error['line_no']}：{min_error['error_pct']:.3f}%", f"λ = {min_error['lambda_A']:.0f} Å", "#00b4d8"),
        ("最大誤差", f"編號 {max_error['line_no']}：{max_error['error_pct']:.3f}%", f"λ = {max_error['lambda_A']:.0f} Å", "#ff4d6d"),
        ("n 範圍", f"{min_n['mean_n']:.4f} → {max_n['mean_n']:.4f}", "短波長折射率較高", "#ff9f1c"),
    ]
    card_html = "".join(
        f"""
        <div class="stat-card" style="--accent:{color}">
          <div class="stat-label">{html_escape(label)}</div>
          <div class="stat-value">{html_escape(value)}</div>
          <div class="stat-note">{html_escape(note)}</div>
        </div>
        """ for label, value, note, color in summary_cards
    )

    summary_rows = []
    for s in summary:
        summary_rows.append([
            f"編號 {s['line_no']}",
            s["lambda_A"],
            s["lambda_nm"],
            s["mean_D_deg"],
            s["mean_n"],
            s["std_n"],
            s["theory_n"],
            s["error_pct"],
        ])
    summary_table = table_html(
        ["光譜線", "λ (Å)", "λ (nm)", "平均 D (°)", "實驗平均 n", "標準差", "理論 n", "誤差 (%)"],
        summary_rows,
        numeric_digits=5,
    )

    record_rows = []
    for r in sorted(records, key=lambda x: (x["line_no"], x["trial"])):
        record_rows.append([
            f"編號 {r['line_no']}",
            r["lambda_A"],
            r["trial"],
            r["d1_deg"],
            r["d2_deg"],
            r["D_deg"],
            r["n"],
        ])
    record_table = table_html(
        ["光譜線", "λ (Å)", "次數", "d1 (°)", "d2 (°)", "D (°)", "n"],
        record_rows,
        numeric_digits=5,
    )

    coeff_rows = [[k, coeff[k]] for k in ["B1", "B2", "B3", "C1", "C2", "C3"]]
    coeff_table = table_html(["Sellmeier 係數", "值"], coeff_rows, numeric_digits=8)

    html = f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>實驗九：折射率、色散、光譜儀 HTML 分析報告</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
:root {{
  --bg0:#07111f;
  --bg1:#0f172a;
  --card:rgba(255,255,255,.88);
  --line:rgba(148,163,184,.35);
  --text:#1f2937;
  --muted:#64748b;
  --title:#0f172a;
}}
* {{ box-sizing:border-box; }}
html {{ scroll-behavior:smooth; }}
body {{
  margin:0;
  font-family:"Microsoft JhengHei", "Noto Sans TC", "Segoe UI", Arial, sans-serif;
  color:var(--text);
  background:
    radial-gradient(circle at 5% 8%, rgba(255,77,109,.35), transparent 28%),
    radial-gradient(circle at 90% 10%, rgba(77,150,255,.38), transparent 26%),
    radial-gradient(circle at 70% 75%, rgba(6,214,160,.25), transparent 28%),
    linear-gradient(135deg, #ecfeff 0%, #eef2ff 35%, #fff1f2 70%, #fffbeb 100%);
  min-height:100vh;
}}
.container {{ max-width:1380px; margin:0 auto; padding:28px 22px 70px; }}
.hero {{
  position:relative;
  overflow:hidden;
  border-radius:34px;
  padding:36px 38px;
  color:#fff;
  background:linear-gradient(135deg, #111827 0%, #312e81 38%, #0369a1 67%, #059669 100%);
  box-shadow:0 28px 70px rgba(30,41,59,.24);
}}
.hero::before {{
  content:""; position:absolute; inset:-80px -40px auto auto; width:360px; height:360px;
  background:conic-gradient(from 90deg, #ff4d6d, #ffd166, #06d6a0, #4d96ff, #7b2ff7, #ff4d6d);
  opacity:.42; filter:blur(10px); border-radius:50%;
}}
.hero h1 {{ position:relative; margin:0 0 12px; font-size:42px; letter-spacing:.5px; line-height:1.2; }}
.hero p {{ position:relative; margin:8px 0; font-size:21px; opacity:.96; }}
.nav {{ position:relative; display:flex; flex-wrap:wrap; gap:10px; margin-top:22px; }}
.nav a {{
  color:#fff; text-decoration:none; font-weight:700;
  background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.32);
  padding:9px 14px; border-radius:999px; backdrop-filter:blur(8px);
}}
.nav a:hover {{ background:rgba(255,255,255,.25); }}
.card {{
  background:var(--card);
  backdrop-filter:blur(12px);
  border:1px solid rgba(255,255,255,.7);
  border-radius:28px;
  padding:24px 26px;
  margin:22px 0;
  box-shadow:0 18px 45px rgba(30,41,59,.13);
}}
.card h2 {{ margin:0 0 15px; color:var(--title); font-size:27px; }}
.card h3 {{ margin:26px 0 12px; color:#1e3a8a; font-size:20px; }}
.badge {{
  display:inline-flex; align-items:center; gap:8px;
  border-radius:999px; padding:7px 14px; margin-bottom:12px;
  background:linear-gradient(135deg, rgba(124,47,247,.14), rgba(0,180,216,.14));
  border:1px solid rgba(99,102,241,.25); color:#4338ca; font-weight:900;
}}
.stat-grid {{ display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin-top:16px; }}
.stat-card {{
  border-radius:24px; padding:18px 18px;
  background:linear-gradient(180deg, rgba(255,255,255,.96), rgba(248,250,252,.88));
  border:1px solid rgba(148,163,184,.25);
  box-shadow:inset 0 4px 0 var(--accent), 0 10px 24px rgba(15,23,42,.08);
}}
.stat-label {{ font-size:14px; color:var(--muted); font-weight:800; }}
.stat-value {{ font-size:26px; color:#0f172a; font-weight:950; margin:5px 0; }}
.stat-note {{ font-size:14px; color:#64748b; }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:22px; }}
.chart {{
  min-height:570px; border-radius:24px; overflow:hidden;
  background:rgba(255,255,255,.92); border:1px solid rgba(148,163,184,.28);
  box-shadow:0 14px 34px rgba(15,23,42,.10);
}}
.chart-wide {{ min-height:610px; }}
.formula {{
  font-size:18px; padding:18px 20px; border-radius:22px;
  background:linear-gradient(135deg, rgba(255,255,255,.96), rgba(239,246,255,.95));
  border:1px solid rgba(59,130,246,.22); color:#1e293b;
}}
.formula code {{ color:#be123c; font-weight:900; }}
.table-wrap {{ overflow:auto; border-radius:22px; border:1px solid rgba(148,163,184,.28); background:#fff; }}
.data-table {{ width:100%; border-collapse:collapse; min-width:900px; font-size:15px; }}
.data-table th {{
  position:sticky; top:0; z-index:1;
  background:linear-gradient(135deg, #1d4ed8, #7c3aed);
  color:#fff; padding:12px 14px; text-align:center; white-space:nowrap;
}}
.data-table td {{ padding:10px 13px; border-bottom:1px solid #e5e7eb; text-align:right; white-space:nowrap; }}
.data-table td:first-child {{ text-align:center; font-weight:800; color:#1e40af; }}
.data-table tbody tr:nth-child(even) {{ background:#f8fafc; }}
.data-table tbody tr:hover {{ background:#eef6ff; }}
.note {{ color:#475569; font-size:16px; }}
.footer {{ text-align:center; color:#64748b; padding:18px; }}
@media (max-width:1050px) {{ .grid,.stat-grid {{ grid-template-columns:1fr; }} .hero h1 {{ font-size:32px; }} }}
</style>
</head>
<body>
<div class="container">
  <section class="hero">
    <h1>實驗九：折射率、色散、光譜儀<br>Interactive HTML Analysis</h1>
    <p><strong>學號：</strong>E24144773　<strong>姓名：</strong>彭靖翔</p>
    <p>由 <strong>exp9.xlsx</strong> 自動擷取最小偏向角、折射率、Sellmeier 理論值與誤差，並產生可互動縮放、滑鼠懸停查看大字提示的報告。</p>
    <div class="nav">
      <a href="#purpose">程式目的</a><a href="#summary">摘要</a><a href="#charts">互動圖表</a><a href="#tables">資料表</a>
    </div>
  </section>

  <section id="purpose" class="card">
    <div class="badge">程式目的</div>
    <p class="note">本程式會讀取一次 <code>exp9.xlsx</code>，直接解析工作表內的 8 條汞燈光譜線資料，包含每條光譜線 5 次量測的 d1、d2、最小偏向角 D 與折射率 n，並整理出平均折射率、標準差、Sellmeier 理論值與百分誤差。輸出的 HTML 使用 Plotly 互動圖表，滑鼠移動到資料點時會以較大的提示框顯示完整資訊，方便在報告展示或口頭說明時直接使用。</p>
    <div class="formula">
      折射率計算：<code>n = sin((D + A) / 2) / sin(A / 2)</code>，本實驗三稜鏡頂角 <code>A = 60°</code>。<br>
      理論折射率：使用 SF-11 的 Sellmeier equation，由波長 λ 計算理論 n，再與實驗平均值比較誤差。
    </div>
  </section>

  <section id="summary" class="card">
    <div class="badge">摘要與重點結果</div>
    <div class="stat-grid">{card_html}</div>
    <p class="note">整體趨勢顯示，波長較短的紫光、藍紫光折射率較高，而波長較長的紅光折射率較低，符合正常色散現象。互動圖中點的顏色也會隨誤差或數值大小變化，便於快速判讀。</p>
  </section>

  <section id="charts" class="card">
    <div class="badge">互動圖表</div>
    <div class="grid">
      <div id="plot1" class="chart chart-wide"></div>
      <div id="plot2" class="chart chart-wide"></div>
    </div>
    <div class="grid" style="margin-top:22px;">
      <div id="plot3" class="chart chart-wide"></div>
      <div id="plot4" class="chart chart-wide"></div>
    </div>
  </section>

  <section id="tables" class="card">
    <div class="badge">資料表</div>
    <h3>表1：8 條光譜線整理結果</h3>
    <div class="table-wrap">{summary_table}</div>
    <h3>表2：40 筆原始量測資料</h3>
    <div class="table-wrap">{record_table}</div>
    <h3>表3：Sellmeier 係數</h3>
    <div class="table-wrap">{coeff_table}</div>
  </section>

  <div class="footer">Generated by exp9_to_html.py｜滑鼠提示字體已放大，圖表可縮放、拖曳與下載。</div>
</div>
<script>
const CONFIG = {{responsive:true, displaylogo:false, modeBarButtonsToRemove:["lasso2d", "select2d"]}};
Plotly.newPlot("plot1", {json.dumps(fig1['traces'], ensure_ascii=False)}, {json.dumps(fig1['layout'], ensure_ascii=False)}, CONFIG);
Plotly.newPlot("plot2", {json.dumps(fig2['traces'], ensure_ascii=False)}, {json.dumps(fig2['layout'], ensure_ascii=False)}, CONFIG);
Plotly.newPlot("plot3", {json.dumps(fig3['traces'], ensure_ascii=False)}, {json.dumps(fig3['layout'], ensure_ascii=False)}, CONFIG);
Plotly.newPlot("plot4", {json.dumps(fig4['traces'], ensure_ascii=False)}, {json.dumps(fig4['layout'], ensure_ascii=False)}, CONFIG);
</script>
</body>
</html>
"""
    return html


def main():
    if not XLSX_PATH.exists():
        raise FileNotFoundError(f"找不到 {XLSX_PATH}，請確認 exp9.xlsx 與本程式在同一個資料夾。")
    coeff, summary, records, source_sheet_name = parse_exp9_xlsx(XLSX_PATH)
    html = build_html(coeff, summary, records, source_sheet_name)
    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"已輸出：{HTML_PATH}")
    print(f"共整理 {len(summary)} 條光譜線、{len(records)} 筆量測資料。")

if __name__ == "__main__":
    main()
