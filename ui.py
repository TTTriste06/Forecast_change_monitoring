import streamlit as st

def setup_sidebar():
    st.sidebar.header("📤 工具简介")
    st.sidebar.markdown("请上传以下文件以生成主计划（不更新文件不用上传）")

def get_uploaded_files():
    st.subheader("📈 上传预测数据")
    forecast_file = st.file_uploader("上传预测数据", type="xlsx", key="forecast")

    st.subheader("📦 上传未交订单")
    order_file = st.file_uploader("上传未交订单(Sheet)", type="xlsx", key="order")

    st.subheader("🚚 上传出货明细")
    sales_file = st.file_uploader("上传出货明细(原表)", type="xlsx", key="sales")

    st.subheader("🔁 上传新旧料号")
    mapping_file = st.file_uploader("上传新旧料号", type="xlsx", key="mapping")

    start = st.button("🚀 生成主计划")
    return template_file, forecast_file, order_file, sales_file, mapping_file, start
