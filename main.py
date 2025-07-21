import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pivot_processor import PivotProcessor  # 假设你的类已定义在 pivot_processor.py
import re

def plot_combined_chart_for_product(df: pd.DataFrame, product_name: str):
    row = df[df["品名"] == product_name]
    if row.empty:
        st.warning(f"未找到品名：{product_name}")
        return

    row = row.iloc[0]

    order_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}-订单", col)]
    ship_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}-出货", col)]
    forecast_cols = [col for col in df.columns if re.match(r"\d{4}-\d{2}的预测（\d{4}-\d{2}生成）", col)]

    months = sorted(list({col[:7] for col in order_cols + ship_cols + forecast_cols}))
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
        ax1.plot(x, y, marker='o', label=f"预测（{gen_month}生成）")

    ax1.set_xticks(x)
    ax1.set_xticklabels(months, rotation=45)
    ax1.set_ylabel("数量")
    ax1.set_title(f"{product_name} - 每月订单、出货与预测")
    ax1.legend()
    ax1.grid(True)

    st.pyplot(fig)


def main():
    st.title("📊 预测数据分析工具")

    # 上传所需文件
    forecast_files = st.file_uploader("上传多个预测文件", type=["xlsx"], accept_multiple_files=True)
    order_file = st.file_uploader("上传订单文件", type=["xlsx"])
    sales_file = st.file_uploader("上传出货文件", type=["xlsx"])
    mapping_file = st.file_uploader("上传新旧料号映射文件", type=["xlsx"])

    if st.button("开始处理") and forecast_files and order_file and sales_file and mapping_file:
        try:
            main_df, output = processor.process(forecast_files, order_df, sales_df, mapping_df)
            order_df = pd.read_excel(order_file)
            sales_df = pd.read_excel(sales_file)
            mapping_df = pd.read_excel(mapping_file)

            processor = PivotProcessor()
            main_df, output = processor.process(forecast_dfs, order_df, sales_df, mapping_df)

            # 展示结果
            st.success("✅ 数据处理完成！")
            st.dataframe(main_df)

            # 下载按钮
            st.download_button(
                label="📥 下载预测分析结果 Excel",
                data=output,
                file_name=f"预测分析_{pd.Timestamp.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 品名图表展示
            if not main_df.empty:
                st.header("📈 品名预测趋势图")
                product_list = main_df["品名"].dropna().unique().tolist()
                selected_product = st.selectbox("请选择品名", product_list)
                plot_combined_chart_for_product(main_df, selected_product)

        except Exception as e:
            st.error(f"❌ 处理失败：{e}")

if __name__ == "__main__":
    main()
