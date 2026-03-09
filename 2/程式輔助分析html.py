# -*- coding: utf-8 -*-
import re
import numpy as np
import pandas as pd
from scipy.interpolate import CubicSpline
import plotly.graph_objects as go

# =========================================================
# 互動圖版本（Plotly）：
# - 滑鼠移到點上才顯示 (頻率, V, dV/df)
# - 同時畫出 Spline 平滑曲線 + 原始資料點
# - 讀 Excel 時會自動找到表頭列（因為前面有 τ 理論值等說明列）
# =========================================================

# ===== Excel 檔名/路徑=====
xlsx_path = r"RC電路對交流訊號的反應.xlsx"
sheet_name = "工作表1"  # 不確定可改 None（掃描全部工作表）

# ---------------------------------------------------------
# 工具函式：字串正規化（去空白）
# ---------------------------------------------------------
def normalize(s):
    s = "" if s is None else str(s)
    return re.sub(r"\s+", "", s)

# ---------------------------------------------------------
# 找表頭列：同一列同時含「頻率」「Vc」「Vr」就判定為表頭
# ---------------------------------------------------------
def find_header_row(df_raw):
    freq_keys = ["頻率(Hz)", "頻率（Hz）", "頻率Hz", "頻率"]
    vc_keys   = ["Vc(mV)", "V_c(mV)", "Vc", "V_c", "電容"]
    vr_keys   = ["Vr(mV)", "V_r(mV)", "Vr", "V_r", "電阻"]

    for r in range(min(len(df_raw), 200)):  # 最多掃前 200 列
        row = [normalize(x) for x in df_raw.iloc[r].tolist()]

        def has_any(keys):
            keys = [normalize(k) for k in keys]
            return any(any(k in cell for k in keys) for cell in row)

        if has_any(freq_keys) and has_any(vc_keys) and has_any(vr_keys):
            return r
    return None

# ---------------------------------------------------------
# 依關鍵字找到對應欄位名稱
# ---------------------------------------------------------
def get_col_by_keywords(columns, keywords):
    cols_norm = [normalize(c) for c in columns]
    for i, cn in enumerate(cols_norm):
        for k in keywords:
            if normalize(k) in cn:
                return columns[i]
    return None

# ---------------------------------------------------------
# Spline 平滑（回傳平滑後的 x_new, y_new）
# ---------------------------------------------------------
def spline_smooth(x, y, n=600):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    # 去掉 NaN/inf
    mask = np.isfinite(x) & np.isfinite(y)
    x, y = x[mask], y[mask]

    # 確保 x 遞增
    idx = np.argsort(x)
    x, y = x[idx], y[idx]

    # CubicSpline 需要 x 嚴格遞增，移除重複 x
    x_unique, uidx = np.unique(x, return_index=True)
    y_unique = y[uidx]

    # 點太少就不做 spline（避免失敗）
    if len(x_unique) < 4:
        return x_unique, y_unique

    cs = CubicSpline(x_unique, y_unique, bc_type="natural")
    x_new = np.linspace(x_unique.min(), x_unique.max(), n)
    y_new = cs(x_new)
    return x_new, y_new

# ---------------------------------------------------------
# 畫互動圖（hover 顯示頻率、V、dV/df）
# ---------------------------------------------------------
def plot_interactive(x, y, title, y_label, out_html, trace_name="V"):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    # 清理資料
    mask = np.isfinite(x) & np.isfinite(y)
    x, y = x[mask], y[mask]

    # 排序
    idx = np.argsort(x)
    x, y = x[idx], y[idx]

    # 數值微分：dV/df（單位：mV/Hz）
    dy_df = np.gradient(y, x)

    # Spline 平滑線
    xs, ys = spline_smooth(x, y, n=800)

    # hover 內容（只在滑鼠移上去才顯示）
    # customdata 內放 (V, dV/df)
    customdata = np.column_stack([y, dy_df])

    hovertemplate = (
    "頻率: %{x:.0f} Hz<br>"
    + trace_name + ": %{customdata[0]:.0f} mV<br>"
    + "d" + trace_name + "/df: %{customdata[1]:.4f} mV/Hz<br>"
    + "<extra></extra>"
    )

    fig = go.Figure()

    # 1) Spline 平滑線（不顯示 hover，避免干擾）
    fig.add_trace(
        go.Scatter(
            x=xs, y=ys,
            mode="lines",
            name="Spline 平滑線",
            hoverinfo="skip",
            line=dict(width=3)
        )
    )

    # 2) 原始資料點（hover 顯示資訊）
    fig.add_trace(
        go.Scatter(
            x=x, y=y,
            mode="markers",
            name="原始資料點",
            marker=dict(size=9),
            customdata=customdata,
            hovertemplate=hovertemplate
        )
    )

    # 圖表版面設定
    fig.update_layout(
        title=title,
        xaxis_title="頻率 (Hz)",
        yaxis_title=y_label,
        template="plotly_white",
        width=1100,
        height=550
    )

    # 存成 HTML（用瀏覽器開就能互動）
    fig.write_html(out_html, include_plotlyjs="cdn")
    print(f"已輸出互動圖：{out_html}")

# =========================================================
# 主程式：讀 Excel → 找表頭 → 取出 freq/Vc/Vr → 畫兩張互動圖
# =========================================================

# 用 header=None 讀，因為前面有說明文字列
df_raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)
df_raw = df_raw.dropna(how="all")

# 找表頭所在列
header_row = find_header_row(df_raw)
if header_row is None:
    raise RuntimeError("找不到表頭列（需要同一列含：頻率(Hz)、Vc(mV)、Vr(mV)）。請確認 Excel 表頭文字。")

# 設定表頭 & 取資料
header = df_raw.iloc[header_row].tolist()
df = df_raw.iloc[header_row + 1:].copy()
df.columns = header

# 找出三個欄位
freq_col = get_col_by_keywords(df.columns, ["頻率(Hz)", "頻率（Hz）", "頻率"])
vc_col   = get_col_by_keywords(df.columns, ["Vc(mV)", "V_c(mV)", "Vc", "V_c"])
vr_col   = get_col_by_keywords(df.columns, ["Vr(mV)", "V_r(mV)", "Vr", "V_r"])

if not all([freq_col, vc_col, vr_col]):
    raise RuntimeError(f"欄位對應失敗：freq={freq_col}, Vc={vc_col}, Vr={vr_col}")

# 整理資料表
data = df[[freq_col, vc_col, vr_col]].copy()
data.columns = ["freq_Hz", "Vc_mV", "Vr_mV"]

# 轉數字（遇到文字就轉 NaN）
for c in data.columns:
    data[c] = pd.to_numeric(data[c], errors="coerce")

# 去除無效列並排序
data = data.dropna(subset=["freq_Hz"]).sort_values("freq_Hz")

# Picture1：Vc
plot_interactive(
    data["freq_Hz"].values,
    data["Vc_mV"].values,
    title="Picture1：電容兩端電壓（滑鼠移上點顯示 頻率、Vc、dVc/df）",
    y_label="Vc (mV)",
    out_html="Picture1_Vc_interactive.html",
    trace_name="Vc"
)

# Picture2：Vr
plot_interactive(
    data["freq_Hz"].values,
    data["Vr_mV"].values,
    title="Picture2：電阻兩端電壓（滑鼠移上點顯示 頻率、Vr、dVr/df）",
    y_label="Vr (mV)",
    out_html="Picture2_Vr_interactive.html",
    trace_name="Vr"
)