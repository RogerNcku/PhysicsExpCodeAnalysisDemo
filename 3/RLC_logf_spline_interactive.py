# -*- coding: utf-8 -*-
# 匯入需要用到的套件
import re                       # 正規表示式，用來處理字串
import numpy as np              # 數值運算套件
import pandas as pd             # 讀取與處理 Excel 表格
from scipy.interpolate import CubicSpline   # Spline 平滑曲線
import plotly.graph_objects as go           # 繪製互動式圖表

# Excel 檔案路徑
xlsx_path = r"RLC交流電路.xlsx"

# 指定要讀取的工作表名稱
sheet_name = "RLC 交流電路"

# 先不指定表頭讀取 Excel
# 因為有時前面可能有說明文字，所以先全部讀進來再找真正表頭
df_raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)


# ---------------------------------------------------
# 字串正規化函式
# 用來去除字串中的空白，方便後面判斷欄位名稱
# ---------------------------------------------------
def normalize(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


# ---------------------------------------------------
# 自動尋找 Excel 表頭列
# 找同時包含
# f(kHz)、log(f)Hz、VR1kΩ(V)
# 的那一列當作表頭
# ---------------------------------------------------
header_row = None

for r in range(min(len(df_raw), 20)):  # 只掃描前20列
    row = [normalize(x) for x in df_raw.iloc[r].tolist()]

    if any("f(kHz)" in c for c in row) and \
       any("log(f)Hz" in c for c in row) and \
       any("VR1kΩ(V)" in c for c in row):

        header_row = r
        break


# 如果找不到表頭就報錯
if header_row is None:
    raise ValueError("找不到表頭列，請檢查 Excel 欄位名稱。")


# ---------------------------------------------------
# 重新讀取 Excel
# 使用剛找到的表頭列作為欄位名稱
# ---------------------------------------------------
df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)

# 只保留需要的欄位
df = df[["f(kHz)", "log(f)Hz", "VR1kΩ(V)", "VR2kΩ(V)", "VR4kΩ(V)"]]

# 移除空值
df = df.dropna()


# ---------------------------------------------------
# 取得橫軸資料
# x = log(f)Hz
# fkhz = 原始頻率 (kHz)
# ---------------------------------------------------
x = df["log(f)Hz"].to_numpy(float)
fkhz = df["f(kHz)"].to_numpy(float)


# ---------------------------------------------------
# 三條曲線資料
# ---------------------------------------------------
series = {
    "4kΩ": df["VR4kΩ(V)"].to_numpy(float),
    "2kΩ": df["VR2kΩ(V)"].to_numpy(float),
    "1kΩ": df["VR1kΩ(V)"].to_numpy(float),
}

# 每條曲線的顏色
colors = {
    "4kΩ": "#f4b000",
    "2kΩ": "#6aa84f",
    "1kΩ": "#e67e22"
}


# 建立 Plotly 圖表
fig = go.Figure()


# ---------------------------------------------------
# 依序畫出三條曲線
# ---------------------------------------------------
for name in ["4kΩ", "2kΩ", "1kΩ"]:

    # 對應電壓資料
    y = series[name]

    # 取得唯一的 x 值 (避免 spline 重複點錯誤)
    xu, idx = np.unique(x, return_index=True)
    yu = y[idx]

    # 建立平滑曲線 x 座標
    xs = np.linspace(xu.min(), xu.max(), 800)

    # 使用 CubicSpline 產生平滑曲線
    ys = CubicSpline(xu, yu, bc_type="natural")(xs)

    # ---------------------------------------------------
    # 加入平滑曲線 (Spline)
    # ---------------------------------------------------
    fig.add_trace(go.Scatter(
        x=xs,
        y=ys,
        mode="lines",
        name=name,
        line=dict(width=3, color=colors[name]),
        hoverinfo="skip"   # 滑鼠移到線上不顯示資訊
    ))

    # ---------------------------------------------------
    # 原始資料點 (可 hover 顯示數值)
    # ---------------------------------------------------

    # customdata 用來存 hover 顯示資料
    customdata = np.column_stack([fkhz, x, y])

    fig.add_trace(go.Scatter(
        x=x,
        y=y,
        mode="markers",
        name=name,
        showlegend=False,
        marker=dict(size=10, color=colors[name]),

        customdata=customdata,

        # 滑鼠移到點上時顯示資訊
        hovertemplate=(
            f"{name}<br>"
            "f(kHz): %{customdata[0]:g}<br>"
            "log(f)Hz: %{customdata[1]:.6f}<br>"
            "伏特(V): %{customdata[2]:.2f}<extra></extra>"
        )
    ))


# ---------------------------------------------------
# 圖表版面設定
# ---------------------------------------------------
fig.update_layout(

    # 標題
    title=dict(
        text="對頻率取對數與伏特值的關係圖",
        x=0.5
    ),

    # X 軸設定
    xaxis=dict(
        title="log(f)Hz",
        range=[2, 5],
        tick0=2,
        dtick=0.5,
        gridcolor="#cfcfcf"
    ),

    # Y 軸設定
    yaxis=dict(
        title="伏特(V)",
        range=[0, 6],
        tick0=0,
        dtick=1,
        gridcolor="#cfcfcf"
    ),

    # 背景顏色
    plot_bgcolor="#f2f2f2",
    paper_bgcolor="white",

    # 圖例位置
    legend=dict(
        x=1.02,
        y=0.5,
        yanchor="middle"
    ),

    # 圖大小
    width=1100,
    height=620,

    # 字型
    font=dict(
        family="Microsoft JhengHei, Noto Sans CJK TC, Arial",
        size=18
    ),

    # 滑鼠模式
    hovermode="closest"
)


# ---------------------------------------------------
# 輸出互動 HTML 圖
# ---------------------------------------------------
out_html = "RLC_logf_spline_interactive.html"

fig.write_html(out_html, include_plotlyjs="cdn")

print(f"已輸出: {out_html}")