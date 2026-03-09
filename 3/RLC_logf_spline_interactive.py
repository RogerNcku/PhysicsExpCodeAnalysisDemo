# -*- coding: utf-8 -*-
# =========================================================
# RLC交流電路：互動式 Spline 平滑曲線圖
# 功能：
# 1. 從 Excel 讀取資料
# 2. 以 log(f)Hz 為橫軸，VR1kΩ / VR2kΩ / VR4kΩ 為縱軸
# 3. 繪製 Spline 平滑曲線
# 4. 滑鼠移到原始資料點時，顯示：
#    - f(kHz)
#    - log(f)Hz
#    - 伏特(V)
#    - dV/d(logf)
#    - dV/df(kHz)
# 5. 自動標記：
#    - 最大正斜率點（上升最陡）
#    - 最大負斜率點（下降最陡）
#    - 共振峰值點（最大電壓點）
# 6. 自動畫出峰值點對應的垂直虛線
# =========================================================

# ---------------------------------------------------------
# 匯入需要的套件
# ---------------------------------------------------------
import re
import numpy as np
import pandas as pd
from scipy.interpolate import CubicSpline
import plotly.graph_objects as go

# ---------------------------------------------------------
# Excel 檔案路徑
# ---------------------------------------------------------
xlsx_path = r"RLC交流電路.xlsx"

# ---------------------------------------------------------
# 指定工作表名稱
# ---------------------------------------------------------
sheet_name = "RLC 交流電路"

# ---------------------------------------------------------
# 第一次讀 Excel：先不指定表頭
# ---------------------------------------------------------
df_raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)

# ---------------------------------------------------------
# 字串正規化函式：移除空白，方便判斷表頭名稱
# ---------------------------------------------------------
def normalize(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""

# ---------------------------------------------------------
# 自動尋找表頭列
# 只要某一列同時包含 f(kHz)、log(f)Hz、VR1kΩ(V)
# 就視為真正的表頭
# ---------------------------------------------------------
header_row = None
for r in range(min(len(df_raw), 20)):
    row = [normalize(x) for x in df_raw.iloc[r].tolist()]
    if any("f(kHz)" in c for c in row) and \
       any("log(f)Hz" in c for c in row) and \
       any("VR1kΩ(V)" in c for c in row):
        header_row = r
        break

# ---------------------------------------------------------
# 找不到表頭就停止
# ---------------------------------------------------------
if header_row is None:
    raise ValueError("找不到表頭列，請檢查 Excel 欄位名稱或工作表名稱。")

# ---------------------------------------------------------
# 重新讀取 Excel，使用找到的表頭列作為欄名
# ---------------------------------------------------------
df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)

# ---------------------------------------------------------
# 只保留需要的欄位，並移除空值
# ---------------------------------------------------------
df = df[["f(kHz)", "log(f)Hz", "VR1kΩ(V)", "VR2kΩ(V)", "VR4kΩ(V)"]].dropna()

# ---------------------------------------------------------
# 取出橫軸資料
# x      : log(f)Hz
# fkhz   : 原始頻率(kHz)
# ---------------------------------------------------------
x = df["log(f)Hz"].to_numpy(float)
fkhz = df["f(kHz)"].to_numpy(float)

# ---------------------------------------------------------
# 三條曲線資料
# ---------------------------------------------------------
series = {
    "4kΩ": df["VR4kΩ(V)"].to_numpy(float),
    "2kΩ": df["VR2kΩ(V)"].to_numpy(float),
    "1kΩ": df["VR1kΩ(V)"].to_numpy(float),
}

# ---------------------------------------------------------
# 各曲線顏色
# ---------------------------------------------------------
colors = {
    "4kΩ": "#f4b000",
    "2kΩ": "#6aa84f",
    "1kΩ": "#e67e22"
}

# ---------------------------------------------------------
# 建立 Plotly 圖表
# ---------------------------------------------------------
fig = go.Figure()

# ---------------------------------------------------------
# 依序處理三條曲線
# ---------------------------------------------------------
for name in ["4kΩ", "2kΩ", "1kΩ"]:

    # 取出該條曲線的電壓資料
    y = series[name]

    # -----------------------------------------------------
    # 依 x 由小到大排序，避免微分與插值出錯
    # -----------------------------------------------------
    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]          # log(f)Hz
    y_sorted = y[sort_idx]          # 電壓
    fkhz_sorted = fkhz[sort_idx]    # 原始頻率(kHz)

    # -----------------------------------------------------
    # 計算兩種斜率
    # slope_logf : dV/d(logf)
    # slope_fkhz : dV/df(kHz)
    # -----------------------------------------------------
    slope_logf = np.gradient(y_sorted, x_sorted)
    slope_fkhz = np.gradient(y_sorted, fkhz_sorted)

    # -----------------------------------------------------
    # 為避免 CubicSpline 因 x 重複而失敗，先取唯一 x
    # -----------------------------------------------------
    xu, idx = np.unique(x_sorted, return_index=True)
    yu = y_sorted[idx]

    # -----------------------------------------------------
    # 產生平滑曲線
    # -----------------------------------------------------
    xs = np.linspace(xu.min(), xu.max(), 800)
    ys = CubicSpline(xu, yu, bc_type="natural")(xs)

    # -----------------------------------------------------
    # 加入平滑曲線
    # -----------------------------------------------------
    fig.add_trace(go.Scatter(
        x=xs,
        y=ys,
        mode="lines",
        name=name,
        line=dict(width=3, color=colors[name]),
        hoverinfo="skip"
    ))

    # -----------------------------------------------------
    # 找最大正斜率點、最大負斜率點、峰值點
    #
    # max_pos_idx : 上升最陡
    # max_neg_idx : 下降最陡
    # peak_idx    : 最大電壓點（共振峰值點）
    # -----------------------------------------------------
    max_pos_idx = np.argmax(slope_logf)
    max_neg_idx = np.argmin(slope_logf)
    peak_idx = np.argmax(y_sorted)

    # -----------------------------------------------------
    # 原始資料點 hover 顯示資料
    # -----------------------------------------------------
    customdata = np.column_stack([
        fkhz_sorted,   # [0] 原始頻率(kHz)
        x_sorted,      # [1] log(f)Hz
        y_sorted,      # [2] 電壓(V)
        slope_logf,    # [3] dV/d(logf)
        slope_fkhz     # [4] dV/df(kHz)
    ])

    # -----------------------------------------------------
    # 加入原始資料點（可 hover）
    # -----------------------------------------------------
    fig.add_trace(go.Scatter(
        x=x_sorted,
        y=y_sorted,
        mode="markers",
        name=name,
        showlegend=False,
        marker=dict(size=10, color=colors[name]),
        customdata=customdata,
        hovertemplate=(
            f"{name}<br>"
            "f(kHz): %{customdata[0]:g}<br>"
            "log(f)Hz: %{customdata[1]:.6f}<br>"
            "伏特(V): %{customdata[2]:.2f}<br>"
            "dV/d(logf): %{customdata[3]:.4f}<br>"
            "dV/df(kHz): %{customdata[4]:.6f}<extra></extra>"
        )
    ))

    # -----------------------------------------------------
    # 標記最大正斜率點
    # -----------------------------------------------------
    fig.add_trace(go.Scatter(
        x=[x_sorted[max_pos_idx]],
        y=[y_sorted[max_pos_idx]],
        mode="markers+text",
        name=f"{name} 最大正斜率",
        marker=dict(size=16, color=colors[name], symbol="star"),
        text=[f"{name} 最大正斜率"],
        textposition="top center",
        customdata=[[
            fkhz_sorted[max_pos_idx],
            x_sorted[max_pos_idx],
            y_sorted[max_pos_idx],
            slope_logf[max_pos_idx],
            slope_fkhz[max_pos_idx]
        ]],
        hovertemplate=(
            f"{name} 最大正斜率點<br>"
            "f(kHz): %{customdata[0]:g}<br>"
            "log(f)Hz: %{customdata[1]:.6f}<br>"
            "伏特(V): %{customdata[2]:.2f}<br>"
            "dV/d(logf): %{customdata[3]:.4f}<br>"
            "dV/df(kHz): %{customdata[4]:.6f}<extra></extra>"
        )
    ))

    # -----------------------------------------------------
    # 標記最大負斜率點
    # -----------------------------------------------------
    fig.add_trace(go.Scatter(
        x=[x_sorted[max_neg_idx]],
        y=[y_sorted[max_neg_idx]],
        mode="markers+text",
        name=f"{name} 最大負斜率",
        marker=dict(size=14, color=colors[name], symbol="diamond"),
        text=[f"{name} 最大負斜率"],
        textposition="bottom center",
        customdata=[[
            fkhz_sorted[max_neg_idx],
            x_sorted[max_neg_idx],
            y_sorted[max_neg_idx],
            slope_logf[max_neg_idx],
            slope_fkhz[max_neg_idx]
        ]],
        hovertemplate=(
            f"{name} 最大負斜率點<br>"
            "f(kHz): %{customdata[0]:g}<br>"
            "log(f)Hz: %{customdata[1]:.6f}<br>"
            "伏特(V): %{customdata[2]:.2f}<br>"
            "dV/d(logf): %{customdata[3]:.4f}<br>"
            "dV/df(kHz): %{customdata[4]:.6f}<extra></extra>"
        )
    ))

    # -----------------------------------------------------
    # 標記共振峰值點（最大電壓點）
    # 使用較大的圓點，並加上文字說明
    # -----------------------------------------------------
    fig.add_trace(go.Scatter(
        x=[x_sorted[peak_idx]],
        y=[y_sorted[peak_idx]],
        mode="markers",
        # mode="markers+text",
        name=f"{name} 峰值點",
        marker=dict(
            size=18,
            color=colors[name],
            symbol="circle-open",
            line=dict(width=3, color=colors[name])
        ),
        text=[f"{name} 峰值"],
        textposition="top right",
        customdata=[[
            fkhz_sorted[peak_idx],
            x_sorted[peak_idx],
            y_sorted[peak_idx],
            slope_logf[peak_idx],
            slope_fkhz[peak_idx]
        ]],
        hovertemplate=(
            f"{name} 共振峰值點<br>"
            "f(kHz): %{customdata[0]:g}<br>"
            "log(f)Hz: %{customdata[1]:.6f}<br>"
            "伏特(V): %{customdata[2]:.2f}<br>"
            "dV/d(logf): %{customdata[3]:.4f}<br>"
            "dV/df(kHz): %{customdata[4]:.6f}<extra></extra>"
        )
    ))

    # -----------------------------------------------------
    # 畫峰值對應的垂直虛線
    # x0 = x1 = 峰值點的 log(f)Hz
    # 從 y=0 畫到 y=峰值電壓
    # line_dash="dash" 表示虛線
    # opacity 讓虛線稍微透明
    # -----------------------------------------------------
    fig.add_shape(
        type="line",
        x0=x_sorted[peak_idx],
        x1=x_sorted[peak_idx],
        y0=0,
        y1=y_sorted[peak_idx],
        line=dict(
            color=colors[name],
            width=2,
            dash="dash"
        ),
        opacity=0.8
    )

    # -----------------------------------------------------
    # 在虛線底部附近加註峰值頻率文字
    # 這樣可以直接看出峰值出現在哪個頻率
    # -----------------------------------------------------
    fig.add_annotation(
        x=x_sorted[peak_idx],
        y=0.15,
        text=f"{name}<br>f={fkhz_sorted[peak_idx]:g} kHz",
        showarrow=False,
        font=dict(size=12, color=colors[name]),
        bgcolor="rgba(255,255,255,0.7)"
    )

# ---------------------------------------------------------
# 設定整體版面
# ---------------------------------------------------------
fig.update_layout(
    title=dict(
        text="對頻率取對數與伏特值的關係圖",
        x=0.5
    ),
    xaxis=dict(
        title="log(f)Hz",
        range=[2, 5],
        tick0=2,
        dtick=0.5,
        gridcolor="#cfcfcf"
    ),
    yaxis=dict(
        title="伏特(V)",
        range=[0, 6],
        tick0=0,
        dtick=1,
        gridcolor="#cfcfcf"
    ),
    plot_bgcolor="#f2f2f2",
    paper_bgcolor="white",
    legend=dict(
        x=1.02,
        y=0.5,
        yanchor="middle"
    ),
    width=1250,
    height=720,
    font=dict(
        family="Microsoft JhengHei, Noto Sans CJK TC, Arial",
        size=18
    ),
    hovermode="closest"
)

# ---------------------------------------------------------
# 輸出 HTML
# ---------------------------------------------------------
out_html = "RLC_logf_spline_interactive_with_peak_and_slope_marks.html"
fig.write_html(out_html, include_plotlyjs="cdn")

print(f"已輸出: {out_html}")