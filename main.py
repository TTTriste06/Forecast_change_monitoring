import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import re
from datetime import datetime

st.set_page_config(page_title="预测分析工具", layout="wide")

# 页面选择
page = st.sidebar.selectbox("📂 请选择页面", ["📊 主计划生成", "📈 图表查看"])

# 模拟预测数据（你可换成真实处理流程）
def generate_mock_df():
    data = {
        "品名": ["ABC"],
        "2025-08-订单": [100], "2025-08-出货": [90], "2025-08的预测（2025-07生成）": [95],
        "2025-09-订单": [120], "2025-09-出货": [110], "2025-09的预测（2025-07生成）": [130],
        "2025-10-订单": [85], "2025-10-出货": [80], "2025-10的预测（2025-07生成）": [92],
    }
    return pd.DataFrame(data)

# 图表绘制函数
def plot_combined_chart(df, product_name):
    row = df[df["品名"] == product_name]
    if row.empty:
        st.warning("未找到品名")
        return
    row = row.iloc[0]

    order_cols = [c for c in df.columns if re.match(r"\d{4}-\d{2}-订单", c)]
    ship_cols = [c for c in df.columns if re.match(r"\d{4}-\d{2}-出货", c)]
    forecast_cols = [c for c in df.columns if re.match(r"\d{4}-\d{2}的预测（\d{4}-\d{2}生成）", c)]

    months = sorted(set(col[:7] for col in order_cols + ship_cols + forecast_cols))
    x = list(range(len(months)))
    order_data = [row.get(f"{m}-订单", 0) for m in months]
    ship_data = [row.get(f"{m}-出货", 0) for m in months]

    forecast_groups = {}
    for col in forecast_cols:
        m, g = re.findall(r"\d{4}-\d{2}", col)
        forecast_groups.setdefault(g, {})[m] = row[col]

    fig, ax1 = plt.subplots(figsize=(12, 5))
    bar_width = 0.35

    ax1.bar([i - bar_width / 2 for i in x], order_data, bar_width, label="订单", color="skyblue")
    ax1.bar([i + bar_width / 2 for i in x], ship_data, bar_width, label="出货", color="orange")

    for gen_month, forecast_dict in forecast_groups.items():
        y = [forecast_dict.get(m, 0) for m in months]
        ax1.plot(x, y, label=f"预测（{gen_month}生成）", marker='o')

    ax1.set_xticks(x)
    ax1.set_xticklabels(months, rotation=45)
    ax1.set_ylabel("数量")
    ax1.set_title(f"{product_name} - 月度订单、出货、预测")
    ax1.legend()
    ax1.grid(True)
    st.pyplot(fig)

# 页面一：主计划生成
if page == "📊 主计划生成":
    st.title("📊 主计划生成页面")
    if st.button("生成模拟数据"):
        df_result = generate_mock_df()
        st.session_state["df_result"] = df_result
        st.success("✅ 模拟数据已生成")
        st.dataframe(df_result)

# 页面二：图表查看
elif page == "📈 图表查看":
    st.title("📈 图表查看页面")
    if "df_result" not in st.session_state:
        st.warning("请先在“主计划生成”页面生成数据")
    else:
        df_result = st.session_state["df_result"]
        product_list = df_result["品名"].dropna().unique().tolist()
        selected = st.selectbox("选择品名", product_list)
        plot_combined_chart(df_result, selected)
