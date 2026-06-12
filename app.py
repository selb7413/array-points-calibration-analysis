import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(
    page_title="XY 點位分析工具",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# 穩定版 CSS：淺色背景、深色字體、不遮擋元件
# =========================
st.markdown("""
<style>
/* 全介面白底黑字 */
.stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], [data-testid="stHeader"] {
    background-color: #FFFFFF !important;
    color: #111111 !important;
}
.block-container {
    padding-top: 1.2rem;
    padding-left: 1.4rem;
    padding-right: 1.4rem;
    max-width: 100%;
}
h1, h2, h3, h4, h5, h6, p, label, span, div {
    color: #111111 !important;
}
[data-testid="stSidebar"] {
    border-right: 1px solid #DDDDDD;
}
[data-testid="stSidebar"] * {
    color: #111111 !important;
}
.clean-card, .metric-card {
    background-color: #FFFFFF !important;
    border: 1px solid #DDDDDD;
    border-radius: 14px;
    padding: 14px;
    box-shadow: none;
    margin-bottom: 16px;
}
.metric-label {
    color: #333333 !important;
    font-size: 13px;
}
.metric-value {
    color: #111111 !important;
    font-size: 24px;
    font-weight: 800;
}
div[data-testid="stFileUploader"] section {
    background-color: #FFFFFF !important;
    border: 1px dashed #999999;
    border-radius: 10px;
}
div[data-testid="stFileUploader"] * {
    color: #111111 !important;
}
[data-baseweb="input"], [data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border-color: #999999 !important;
}
[data-baseweb="input"] input {
    color: #111111 !important;
    -webkit-text-fill-color: #111111 !important;
    opacity: 1 !important;
    caret-color: #111111 !important;
}
[data-testid="stNumberInput"] button {
    background-color: #F5F5F5 !important;
    color: #111111 !important;
    border-color: #999999 !important;
}
[data-testid="stNumberInput"] button svg {
    fill: #111111 !important;
}
[data-baseweb="select"] * {
    color: #111111 !important;
}
[data-testid="stRadio"] label span,
[data-testid="stCheckbox"] label span {
    color: #111111 !important;
}
.color-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 6px 0 10px 0;
    font-size: 13px;
    color: #111111 !important;
}
.color-chip {
    width: 18px;
    height: 18px;
    border-radius: 5px;
    border: 1px solid #555555;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)



# 固定基本顏色：正常點固定綠色、異常判定固定紅色
NORMAL_COLOR = "#22C55E"
ABNORMAL_COLOR = "#EF4444"
DEFAULT_RULE1_COLOR = "#F59E0B"
DEFAULT_RULE2_COLOR = "#7C3AED"

# =========================
# 工具函數
# =========================
def direction_arrow(x_err, y_err, mode):
    arrow = ""

    if pd.isna(x_err) or pd.isna(y_err):
        return ""

    if mode in ["XY都顯示", "只顯示Y方向"]:
        if y_err > 0:
            arrow += "↑"
        elif y_err < 0:
            arrow += "↓"

    if mode in ["XY都顯示", "只顯示X方向"]:
        if x_err > 0:
            arrow += "→"
        elif x_err < 0:
            arrow += "←"

    if mode == "不顯示方向":
        return ""

    return arrow if arrow else "●"


def get_status_and_color(axis_error, threshold,
                         use_rule1, rule1_name, rule1_threshold, rule1_color,
                         use_rule2, rule2_name, rule2_threshold, rule2_color):
    # 判定原則：X 或 Y 任一軸的絕對誤差超過門檻就套用顏色
    if use_rule2 and axis_error >= rule2_threshold:
        return rule2_name, rule2_color
    if use_rule1 and axis_error >= rule1_threshold:
        return rule1_name, rule1_color
    if axis_error >= threshold:
        return "異常", ABNORMAL_COLOR
    return "正常", NORMAL_COLOR


def build_ideal_df(index_range, step):
    data = []
    for iy in index_range[::-1]:
        for ix in index_range:
            data.append({
                "索引X": ix,
                "索引Y": iy,
                "理論X": ix * step,
                "理論Y": iy * step
            })
    return pd.DataFrame(data).round(3)


def parse_excel(uploaded_file, file_name, step,
                threshold,
                use_rule1, rule1_name, rule1_threshold, rule1_color,
                use_rule2, rule2_name, rule2_threshold, rule2_color,
                arrow_mode):

    raw = pd.read_excel(uploaded_file, header=None)

    header_row = None
    for i in range(len(raw)):
        row_text = raw.iloc[i].astype(str).tolist()
        if "基元屬性" in row_text and "實測值" in row_text and "標準值" in row_text:
            header_row = i
            break

    if header_row is None:
        st.error(f"{file_name} 找不到『基元屬性 / 實測值 / 標準值』表頭")
        return None

    df = pd.read_excel(uploaded_file, header=header_row)

    required_cols = ["基元屬性", "實測值", "標準值"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"{file_name} 缺少欄位：{missing_cols}")
        return None

    df = df[required_cols].dropna()
    df = df[df["基元屬性"].astype(str).str.contains("座標", na=False)]

    df["點位編號"] = df["基元屬性"].astype(str).str.extract(r"(\d+)")
    df["座標類型"] = df["基元屬性"].astype(str).apply(
        lambda x: "X" if "座標X" in x else "Y"
    )

    x_df = df[df["座標類型"] == "X"][["點位編號", "實測值", "標準值"]].rename(
        columns={"實測值": "實際X", "標準值": "理論X"}
    )

    y_df = df[df["座標類型"] == "Y"][["點位編號", "實測值", "標準值"]].rename(
        columns={"實測值": "實際Y", "標準值": "理論Y"}
    )

    result = pd.merge(x_df, y_df, on="點位編號", how="inner")

    for col in ["理論X", "理論Y", "實際X", "實際Y"]:
        result[col] = pd.to_numeric(result[col], errors="coerce")

    result = result.dropna(subset=["理論X", "理論Y", "實際X", "實際Y"])

    if result.empty:
        st.error(f"{file_name} 沒有成功解析到 X/Y 座標資料")
        return None

    result["索引X"] = (result["理論X"] / step).round().astype(int)
    result["索引Y"] = (result["理論Y"] / step).round().astype(int)

    result["X誤差"] = result["實際X"] - result["理論X"]
    result["Y誤差"] = result["實際Y"] - result["理論Y"]
    result["總誤差"] = np.sqrt(result["X誤差"] ** 2 + result["Y誤差"] ** 2)
    result["軸向判定誤差"] = np.maximum(result["X誤差"].abs(), result["Y誤差"].abs())

    result["偏移方向"] = result.apply(
        lambda r: direction_arrow(r["X誤差"], r["Y誤差"], arrow_mode),
        axis=1
    )

    status_color_df = result["軸向判定誤差"].apply(
        lambda v: pd.Series(
            get_status_and_color(
                v,
                threshold,
                use_rule1, rule1_name, rule1_threshold, rule1_color,
                use_rule2, rule2_name, rule2_threshold, rule2_color
            )
        )
    )

    result["狀態"] = status_color_df[0]
    result["顏色"] = status_color_df[1]
    result["是否超標"] = result["軸向判定誤差"] >= threshold
    result["檔案名稱"] = file_name
    result["點位"] = "P" + result["點位編號"].astype(str)
    result["點位索引"] = "(" + result["索引X"].astype(str) + ", " + result["索引Y"].astype(str) + ")"

    return result.round(3)




def abnormal_axis_direction(x_err, y_err, threshold,
                            use_rule1, rule1_threshold,
                            use_rule2, rule2_threshold):
    """異常點警示用：只顯示超出規格的軸向偏移方向。"""
    x_limit = threshold
    y_limit = threshold

    # 基準：只要超過異常判定門檻就算該軸異常
    show_x = abs(x_err) >= x_limit
    show_y = abs(y_err) >= y_limit

    arrow = ""

    if show_y:
        if y_err > 0:
            arrow += "↑"
        elif y_err < 0:
            arrow += "↓"

    if show_x:
        if x_err > 0:
            arrow += "→"
        elif x_err < 0:
            arrow += "←"

    return arrow

def axis_cell_style(value, threshold,
                    use_rule1, rule1_threshold, rule1_color,
                    use_rule2, rule2_threshold, rule2_color):
    """只針對 X誤差 / Y誤差 欄位依照單軸絕對誤差著色。"""
    try:
        v = abs(float(value))
    except Exception:
        return ""

    if use_rule2 and v >= rule2_threshold:
        return f"background-color: {rule2_color}; color: white; font-weight: 700;"
    if use_rule1 and v >= rule1_threshold:
        return f"background-color: {rule1_color}; color: white; font-weight: 700;"
    if v >= threshold:
        return f"background-color: {ABNORMAL_COLOR}; color: white; font-weight: 700;"
    return "background-color: white; color: #111111;"


def style_alert_table(df):
    styled = pd.DataFrame("", index=df.index, columns=df.columns)
    for col in ["X誤差", "Y誤差"]:
        if col in df.columns:
            styled[col] = df[col].apply(
                lambda v: axis_cell_style(
                    v,
                    threshold,
                    use_rule1, rule1_threshold, rule1_color,
                    use_rule2, rule2_threshold, rule2_color
                )
            )
    return styled

# =========================
# Sidebar
# =========================
st.sidebar.title("XY 點位分析工具")

uploaded_files = st.sidebar.file_uploader(
    "上傳 Excel 檔案",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

st.sidebar.divider()

st.sidebar.subheader("矩陣設定")
matrix_length = st.sidebar.number_input(
    "矩陣總長度 mm",
    min_value=1.0,
    value=200.0,
    step=1.0
)

point_count = st.sidebar.number_input(
    "單邊點位數",
    min_value=3,
    value=15,
    step=2
)

st.sidebar.subheader("誤差門檻與顏色")

st.sidebar.markdown(
    f'<div class="color-row"><span class="color-chip" style="background:{NORMAL_COLOR};"></span>正常點：固定綠色</div>',
    unsafe_allow_html=True
)

threshold = st.sidebar.number_input(
    "異常判定門檻 mm",
    min_value=0.0,
    value=0.050,
    step=0.001,
    format="%.3f"
)

st.sidebar.markdown(
    f'<div class="color-row"><span class="color-chip" style="background:{ABNORMAL_COLOR};"></span>總誤差 ≥ {threshold:.3f} mm：異常</div>',
    unsafe_allow_html=True
)

st.sidebar.divider()

use_rule1 = st.sidebar.checkbox("啟用自訂門檻 1", value=True)
rule1_name = st.sidebar.text_input(
    "自訂門檻 1 名稱",
    value="警告",
    disabled=not use_rule1
)
rule1_threshold = st.sidebar.number_input(
    "自訂門檻 1 mm",
    min_value=0.0,
    value=0.080,
    step=0.001,
    format="%.3f",
    disabled=not use_rule1
)
rule1_color = st.sidebar.color_picker(
    "自訂門檻 1 顏色",
    DEFAULT_RULE1_COLOR,
    disabled=not use_rule1
)
if use_rule1:
    st.sidebar.markdown(
        f'<div class="color-row"><span class="color-chip" style="background:{rule1_color};"></span>|X誤差| 或 |Y誤差| ≥ {rule1_threshold:.3f} mm：{rule1_name}</div>',
        unsafe_allow_html=True
    )

use_rule2 = st.sidebar.checkbox("啟用自訂門檻 2", value=True)
rule2_name = st.sidebar.text_input(
    "自訂門檻 2 名稱",
    value="嚴重異常",
    disabled=not use_rule2
)
rule2_threshold = st.sidebar.number_input(
    "自訂門檻 2 mm",
    min_value=0.0,
    value=0.100,
    step=0.001,
    format="%.3f",
    disabled=not use_rule2
)
rule2_color = st.sidebar.color_picker(
    "自訂門檻 2 顏色",
    DEFAULT_RULE2_COLOR,
    disabled=not use_rule2
)
if use_rule2:
    st.sidebar.markdown(
        f'<div class="color-row"><span class="color-chip" style="background:{rule2_color};"></span>|X誤差| 或 |Y誤差| ≥ {rule2_threshold:.3f} mm：{rule2_name}</div>',
        unsafe_allow_html=True
    )

st.sidebar.subheader("分布圖設定")
plot_mode = st.sidebar.radio(
    "點位誤差分布圖呈現方式",
    ["靶心圖：X/Y誤差(mm)", "矩陣圖：理想座標 + 實測偏移"],
    index=0
)

arrow_mode = st.sidebar.radio(
    "偏移方向顯示",
    ["XY都顯示", "只顯示X方向", "只顯示Y方向", "不顯示方向"],
    index=0
)

show_label = st.sidebar.checkbox("顯示點位名稱", value=False)

vector_scale = st.sidebar.number_input(
    "矩陣圖偏移箭頭放大倍率",
    min_value=1.0,
    value=30.0,
    step=1.0
)

# =========================
# 基礎矩陣
# =========================
index_range = np.arange(-(point_count // 2), point_count // 2 + 1)
step = matrix_length / (point_count - 1)
ideal_df = build_ideal_df(index_range, step)

# =========================
# 解析 Excel
# =========================
all_results = []

if uploaded_files:
    for file in uploaded_files:
        parsed = parse_excel(
            file,
            file.name,
            step,
            threshold,
            use_rule1, rule1_name, rule1_threshold, rule1_color,
            use_rule2, rule2_name, rule2_threshold, rule2_color,
            arrow_mode
        )
        if parsed is not None:
            all_results.append(parsed)

result_df = pd.concat(all_results, ignore_index=True).round(3) if all_results else pd.DataFrame()

# =========================
# Main Header
# =========================
st.title("XY 點位分析工具")
st.caption(f"更新時間：{datetime.now().strftime('%Y/%m/%d %H:%M:%S')}")

if result_df.empty:
    st.info("請先從左側上傳 Excel 檔案開始分析。")
    with st.expander("理論座標矩陣", expanded=False):
        st.dataframe(ideal_df, use_container_width=True)
    st.stop()

# =========================
# 選檔
# =========================
file_list = sorted(result_df["檔案名稱"].unique())
selected_file = st.selectbox("選擇目前要分析的 Excel", file_list)
plot_df = result_df[result_df["檔案名稱"] == selected_file].copy()

ng_df = plot_df[plot_df["是否超標"] == True].sort_values("總誤差", ascending=False)

# =========================
# KPI
# =========================
avg_error = plot_df["總誤差"].mean()
max_error = plot_df["總誤差"].max()
ng_count = len(ng_df)
std_error = plot_df["總誤差"].std()
cpk_est = threshold / (3 * std_error) if std_error and std_error != 0 else np.nan
cpk_text = "N/A" if pd.isna(cpk_est) else f"{cpk_est:.2f}"

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric("平均誤差(R)", f"{avg_error:.3f} mm")
with k2:
    st.metric("最大誤差(R)", f"{max_error:.3f} mm")
with k3:
    st.metric("異常點數", f"{ng_count} 個")
with k4:
    st.metric("CPK 預估", cpk_text)

# =========================
# 主區：左圖右摘要
# =========================
chart_col, summary_col = st.columns([3.2, 1.2], gap="large")

with chart_col:
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)

    if plot_mode == "靶心圖：X/Y誤差(mm)":
        st.subheader("點位誤差分布圖：靶心圖")
    else:
        st.subheader("點位誤差分布圖：理想矩陣 + 實測偏移")

    text_display = plot_df["點位"] if show_label else plot_df["偏移方向"]

    fig = go.Figure()

    if plot_mode == "靶心圖：X/Y誤差(mm)":
        # -------------------------
        # 模式 1：靶心圖
        # -------------------------
        fig.add_trace(go.Scatter(
            x=plot_df["X誤差"],
            y=plot_df["Y誤差"],
            mode="markers+text",
            text=text_display,
            textposition="top center",
            marker=dict(
                size=14,
                color=plot_df["顏色"],
                line=dict(width=1, color="#333333")
            ),
            customdata=np.stack([
                plot_df["點位"],
                plot_df["索引X"],
                plot_df["索引Y"],
                plot_df["理論X"],
                plot_df["理論Y"],
                plot_df["實際X"],
                plot_df["實際Y"],
                plot_df["X誤差"],
                plot_df["Y誤差"],
                plot_df["總誤差"],
                plot_df["狀態"],
                plot_df["檔案名稱"],
            ], axis=-1),
            hovertemplate=
                "檔案：%{customdata[11]}<br>" +
                "點位：%{customdata[0]}<br>" +
                "索引X：%{customdata[1]}<br>" +
                "索引Y：%{customdata[2]}<br>" +
                "理論X：%{customdata[3]:.3f}<br>" +
                "理論Y：%{customdata[4]:.3f}<br>" +
                "實際X：%{customdata[5]:.3f}<br>" +
                "實際Y：%{customdata[6]:.3f}<br>" +
                "X誤差：%{customdata[7]:.3f} mm<br>" +
                "Y誤差：%{customdata[8]:.3f} mm<br>" +
                "總誤差：%{customdata[9]:.3f} mm<br>" +
                "狀態：%{customdata[10]}<extra></extra>"
        ))

        fig.add_hline(y=0, line_dash="dash", line_color="#777777", line_width=1)
        fig.add_vline(x=0, line_dash="dash", line_color="#777777", line_width=1)

        max_axis = max(
            abs(plot_df["X誤差"]).max(),
            abs(plot_df["Y誤差"]).max(),
            rule2_threshold if use_rule2 else 0,
            rule1_threshold if use_rule1 else 0,
            threshold,
            0.01
        ) * 1.35

        x_axis_title = "X 誤差 (mm)"
        y_axis_title = "Y 誤差 (mm)"
        x_range = [-max_axis, max_axis]
        y_range = [-max_axis, max_axis]

    else:
        # -------------------------
        # 模式 2：理想座標矩陣 + 實測偏移
        # -------------------------
        plot_df["顯示實測X"] = plot_df["理論X"] + plot_df["X誤差"] * vector_scale
        plot_df["顯示實測Y"] = plot_df["理論Y"] + plot_df["Y誤差"] * vector_scale

        # 理想點：淡灰色底點
        fig.add_trace(go.Scatter(
            x=plot_df["理論X"],
            y=plot_df["理論Y"],
            mode="markers",
            marker=dict(size=8, color="#C8C8C8", symbol="circle-open", line=dict(width=1, color="#999999")),
            name="理想位置",
            hovertemplate="理想X：%{x:.3f}<br>理想Y：%{y:.3f}<extra></extra>"
        ))

        # 偏移線段
        for _, row in plot_df.iterrows():
            fig.add_trace(go.Scatter(
                x=[row["理論X"], row["顯示實測X"]],
                y=[row["理論Y"], row["顯示實測Y"]],
                mode="lines",
                line=dict(color=row["顏色"], width=1.6),
                showlegend=False,
                hoverinfo="skip"
            ))

        # 實測偏移後位置：用顏色判斷異常
        fig.add_trace(go.Scatter(
            x=plot_df["顯示實測X"],
            y=plot_df["顯示實測Y"],
            mode="markers+text",
            text=text_display,
            textposition="top center",
            marker=dict(
                size=13,
                color=plot_df["顏色"],
                line=dict(width=1, color="#333333")
            ),
            customdata=np.stack([
                plot_df["點位"],
                plot_df["索引X"],
                plot_df["索引Y"],
                plot_df["理論X"],
                plot_df["理論Y"],
                plot_df["實際X"],
                plot_df["實際Y"],
                plot_df["X誤差"],
                plot_df["Y誤差"],
                plot_df["總誤差"],
                plot_df["狀態"],
                plot_df["檔案名稱"],
            ], axis=-1),
            hovertemplate=
                "檔案：%{customdata[11]}<br>" +
                "點位：%{customdata[0]}<br>" +
                "索引X：%{customdata[1]}<br>" +
                "索引Y：%{customdata[2]}<br>" +
                "理論X：%{customdata[3]:.3f}<br>" +
                "理論Y：%{customdata[4]:.3f}<br>" +
                "實際X：%{customdata[5]:.3f}<br>" +
                "實際Y：%{customdata[6]:.3f}<br>" +
                "X誤差：%{customdata[7]:.3f} mm<br>" +
                "Y誤差：%{customdata[8]:.3f} mm<br>" +
                "總誤差：%{customdata[9]:.3f} mm<br>" +
                "狀態：%{customdata[10]}<extra></extra>"
        ))

        st.caption(f"矩陣圖中的偏移線已放大 {vector_scale:.0f} 倍，方便目視判斷偏移方向。滑鼠移到點位可看實際數值。")

        half = matrix_length / 2
        margin = step * 0.8
        x_axis_title = "理論 / 實測 X 座標 (mm)"
        y_axis_title = "理論 / 實測 Y 座標 (mm)"
        x_range = [-half - margin, half + margin]
        y_range = [-half - margin, half + margin]

    fig.update_layout(
        height=640,
        xaxis=dict(
            title=dict(text=x_axis_title, font=dict(color="#333333")),
            range=x_range,
            zeroline=True,
            zerolinecolor="#999999",
            gridcolor="#E6E6E6",
            linecolor="#CCCCCC",
            tickfont=dict(color="#333333")
        ),
        yaxis=dict(
            title=dict(text=y_axis_title, font=dict(color="#333333")),
            range=y_range,
            zeroline=True,
            zerolinecolor="#999999",
            gridcolor="#E6E6E6",
            linecolor="#CCCCCC",
            tickfont=dict(color="#333333"),
            scaleanchor="x",
            scaleratio=1
        ),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#333333"),
        margin=dict(l=55, r=25, t=25, b=55),
        showlegend=False
    )

    st.plotly_chart(fig, use_container_width=True, config={"responsive": True})
    st.markdown('</div>', unsafe_allow_html=True)

with summary_col:
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.subheader("異常點警示")

    if ng_df.empty:
        st.success("目前無異常點")
    else:
        alert_df = ng_df.head(8)[["點位索引", "X誤差", "Y誤差", "狀態"]].rename(
            columns={"點位索引": "索引座標(X,Y)"}
        )

        alert_df["偏移方向"] = ng_df.head(8).apply(
            lambda r: abnormal_axis_direction(
                r["X誤差"],
                r["Y誤差"],
                threshold,
                use_rule1, rule1_threshold,
                use_rule2, rule2_threshold
            ),
            axis=1
        ).values

        st.dataframe(
            alert_df.style.apply(style_alert_table, axis=None),
            use_container_width=True,
            hide_index=True
        )

    st.subheader("分析結論")

    mean_x = plot_df["X誤差"].mean()
    mean_y = plot_df["Y誤差"].mean()

    if abs(mean_x) > abs(mean_y):
        st.write("• X 方向偏移較明顯，建議優先檢查 X 軸校正。")
    else:
        st.write("• Y 方向偏移較明顯，建議優先檢查 Y 軸校正。")

    if ng_count > 0:
        st.write(f"• 發現 {ng_count} 個異常點，建議檢查 AOI / CCD / 雷射補償參數。")
    else:
        st.write("• 所有點位皆在設定誤差範圍內。")

    if use_rule2 and max_error >= rule2_threshold:
        st.write("• 最大誤差已達自訂門檻2，建議重新確認治具固定與校正參數。")

    st.markdown('</div>', unsafe_allow_html=True)

# =========================
# 表格區
# =========================
with st.expander("理論座標矩陣", expanded=False):
    st.dataframe(ideal_df, use_container_width=True)

with st.expander("解析後資料", expanded=False):
    st.dataframe(plot_df, use_container_width=True)

with st.expander("異常點位總表", expanded=True):
    if ng_df.empty:
        st.success("沒有異常點位")
    else:
        st.error(f"共有 {len(ng_df)} 個點位任一軸超過 {threshold:.3f} mm")
        st.dataframe(
            ng_df[
                [
                    "檔案名稱", "點位", "點位編號", "索引X", "索引Y",
                    "理論X", "理論Y", "實際X", "實際Y",
                    "X誤差", "Y誤差", "總誤差", "軸向判定誤差",
                    "偏移方向", "狀態"
                ]
            ],
            use_container_width=True
        )

# =========================
# 趨勢圖
# =========================
st.markdown('<div class="clean-card">', unsafe_allow_html=True)
st.subheader("點位誤差時間軸")

timeline_df = plot_df.copy()
timeline_df["點位編號排序"] = pd.to_numeric(timeline_df["點位編號"], errors="coerce")
timeline_df = timeline_df.sort_values("點位編號排序")

fig2 = go.Figure()
fig2.add_trace(go.Bar(
    x=timeline_df["點位"],
    y=timeline_df["軸向判定誤差"],
    marker_color=timeline_df["顏色"],
    hovertemplate="點位：%{x}<br>軸向判定誤差：%{y:.3f} mm<extra></extra>"
))

fig2.add_hline(y=threshold, line_dash="dash", line_color=ABNORMAL_COLOR)

fig2.update_layout(
    height=280,
    xaxis=dict(
        title=dict(text="點位", font=dict(color="#333333")),
        tickfont=dict(color="#333333")
    ),
    yaxis=dict(
        title=dict(text="軸向判定誤差(mm)", font=dict(color="#333333")),
        tickfont=dict(color="#333333")
    ),
    plot_bgcolor="#FFFFFF",
    paper_bgcolor="#FFFFFF",
    font=dict(color="#333333"),
    margin=dict(l=55, r=25, t=20, b=55)
)

st.plotly_chart(fig2, use_container_width=True, config={"responsive": True})
st.markdown('</div>', unsafe_allow_html=True)

# =========================
# 匯出
# =========================
output_file = "XY誤差分析結果.xlsx"
result_df.to_excel(output_file, index=False)

with open(output_file, "rb") as f:
    st.download_button(
        label="⬇ 分析報告匯出",
        data=f,
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
