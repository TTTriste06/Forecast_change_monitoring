import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import matplotlib.pyplot as plt
import matplotlib

from ui import get_uploaded_files
from pivot_processor import PivotProcessor
from github_utils import load_file_with_github_fallback

# ✅ 设置中文字体防止乱码
matplotlib.rcParams["font.family"] = "SimHei"
matplotlib.rcParams["axes.unicode_minus"] = False

def main():
    st.set_page_config(page_title="预测分析主计划工具", layout="wide")
    st.title("📊 预测分析主计划生成器")

    forecast_files, order_file, sales_file, mapping_file, start = get_uploaded_files()

    if start:
        try:
            order_df = load_file_with_github_fallback("order", order_file, sheet_name="Sheet")
            sales_df = load_file_with_github_fallback("sales", sales_file, sheet_name="原表")
            mapping_df = load_file_with_github_fallback("mapping", mapping_file, sheet_name=0)

            processor = PivotProcessor()
            df_result, excel_output = processor.process(forecast_files, order_df, sales_df, mapping_df)

            # ✅ 清洗所有预测/订单/出货列为数值，避免后续崩溃
            num_cols = [col for col in df_result.columns if any(kw in col for kw in ["预测", "订单", "出货"])]
            df_result[num_cols] = df_result[num_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

            st.success("✅ 主计划生成成功！")
            st.dataframe(df_result, use_container_width=True)

            st.download_button(
                label="📥 下载主计划 Excel 文件",
                data=excel_output.getvalue(),
                file_name=f"预测分析主计划_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ✅ 图表可视化模块
            st.subheader("📈 可视化分析图表")
            product_list = df_result["品名"].dropna().unique().tolist()
            selected_product = st.selectbox("选择品名进行图表展示", product_list)

            row = df_result[df_result["品名"] == selected_product]
            if not row.empty:
                row = row.iloc[0]

                forecast_cols = [col for col in df_result.columns if "的预测" in col]
                order_cols = [col for col in df_result.columns if "-订单" in col]
                ship_cols = [col for col in df_result.columns if "-出货" in col]

                months = sorted(set(
                    [col.split("的预测")[0] for col in forecast_cols] +
                    [col.split("-订单")[0] for col in order_cols] +
                    [col.split("-出货")[0] for col in ship_cols]
                ))

                x = list(range(len(months)))
                order_data = [row.get(f"{m}-订单", 0) for m in months]
                ship_data = [row.get(f"{m}-出货", 0) for m in months]

                forecast_lines = {}
                for col in forecast_cols:
                    ym = col.split("的预测")[0]
                    gen_date = col.split("（")[-1].replace("生成）", "")
                    forecast_lines.setdefault(gen_date, []).append((ym, row.get(col, 0)))

                for gen_date in forecast_lines:
                    forecast_lines[gen_date].sort()

                fig, ax = plt.subplots(figsize=(12, 6))
                bar_width = 0.35
                ax.bar([i - bar_width/2 for i in x], order_data, width=bar_width, label="订单", color="skyblue")
                ax.bar([i + bar_width/2 for i in x], ship_data, width=bar_width, label="出货", color="lightgreen")

                for gen_date, y_pairs in forecast_lines.items():
                    # ✅ 补齐缺失月份
                    month_value_map = {m: v for m, v in y_pairs}
                    y_sorted = [month_value_map.get(m, 0) for m in months]
                    ax.plot(x, y_sorted, marker="o", label=f"预测（{gen_date}）")

                ax.set_xticks(x)
                ax.set_xticklabels(months, rotation=45)
                ax.set_title(f"{selected_product} 每月订单/出货与预测")
                ax.set_ylabel("数量")
                ax.legend()
                ax.grid(True)

                st.pyplot(fig)

        except Exception as e:
            st.error("❌ 处理过程中出错，请检查数据格式或上传的文件是否正确")
            st.exception(e)

if __name__ == "__main__":
    main()
