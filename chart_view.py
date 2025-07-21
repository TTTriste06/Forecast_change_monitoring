import streamlit as st
import re
import matplotlib.pyplot as plt

def plot_combined_chart(df, product_name):
    row = df[df["品名"] == product_name]
    if row.empty:
        st.warning(f"未找到品名：{product_name}")
        return
    row = row.iloc[0]

    order_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}-订单", col)]
    ship_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}-出货", col)]
    forecast_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}的预测（\d{4}-\d{2}生成）", col)]

    months = sorted({col[:7] for col in order_cols + ship_cols + forecast_cols})
    x = list(range(len(months)))

    order_data = [row.get(f"{m}-订单", 0) for m in months]
    ship_data = [row.get(f"{m}-出货", 0) for m in months]

    forecast_groups = {}
    for col in forecast_cols:
        match = re.match(r"(\d{4}-\d{2})的预测（(\d{4}-\d{2})生成）", col)
        if match:
            forecast_month, gen_month = match.groups()
            forecast_groups.setdefault(gen_month, {})[forecast_month] = row.get(col, 0)

    fig, ax1 = plt.subplots(figsize=(12, 5))
    bar_width = 0.35

    ax1.bar([i - bar_width/2 for i in x], order_data, bar_width, label="订单", color="skyblue")
    ax1.bar([i + bar_width/2 for i in x], ship_data, bar_width, label="出货", color="orange")

    for gen_month, forecast_dict in forecast_groups.items():
        y = [forecast_dict.get(m, 0) for m in months]
        ax1.plot(x, y, label=f"预测（{gen_month}生成）", marker='o')

    ax1.set_xticks(x)
    ax1.set_xticklabels(months, rotation=45)
    ax1.set_ylabel("数量")
    ax1.set_title(f"{product_name} - 每月订单、出货与预测")
    ax1.legend()
    ax1.grid(True)

    st.pyplot(fig)


def main():
    st.set_page_config(page_title="图表分析", layout="wide")
    st.title("📈 品名预测趋势图")

    if "df_result" not in st.session_state:
        st.warning("请先前往主页面生成主计划数据。")
        return

    df_result = st.session_state["df_result"]
    st.caption(f"🕒 数据更新时间：{st.session_state.get('last_updated', '未知')}")

    product_list = df_result["品名"].dropna().unique().tolist()
    selected = st.selectbox("请选择品名", product_list)
    plot_combined_chart(df_result, selected)

if __name__ == "__main__":
    main()
