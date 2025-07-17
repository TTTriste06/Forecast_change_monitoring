import streamlit as st
import pandas as pd
import re
from datetime import datetime
from mapping_utils import (
    apply_all_name_replacements,
    replace_all_names_with_mapping,
    clean_mapping_headers,
    split_mapping_data,
)
from io import BytesIO

st.set_page_config("📊 多预测整合工具", layout="wide")
st.title("📊 多预测整合与品名提取")

# ===== 上传区 =====
forecast_files = st.file_uploader("📁 上传多个预测文件", type=["xlsx"], accept_multiple_files=True)
order_file = st.file_uploader("📄 上传订单文件（含“晶圆品名”）", type=["xlsx"])
sales_file = st.file_uploader("📄 上传出货文件（含“品名、晶圆、规格”）", type=["xlsx"])
mapping_file = st.file_uploader("🧭 上传新旧料号映射表", type=["xlsx"])

# ===== 配置字段映射（用于 apply_all_name_replacements） =====
FIELD_MAPPINGS = {
    "预测": {"品名": "品名"},
    "订单": {"品名": "晶圆品名"},
    "出货": {"品名": "品名"},
}

# ===== 按钮触发主流程 =====
if st.button("🚀 开始处理") and forecast_files and order_file and sales_file and mapping_file:
    # 1️⃣ 解析新旧料号映射
    mapping_raw = pd.read_excel(mapping_file)
    st.write(mapping_raw)
    mapping_df = clean_mapping_headers(mapping_raw)
    st.write(mapping_df)
    mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

    current_year = datetime.now().year
    all_parts = []

    # 2️⃣ 处理每个预测文件
    for uploaded_file in forecast_files:
        filename = uploaded_file.name
        # 读取最长 sheet
        xls = pd.ExcelFile(uploaded_file)
        sheet_lengths = {sheet: pd.read_excel(xls, sheet).shape[0] for sheet in xls.sheet_names}
        longest_sheet = max(sheet_lengths, key=sheet_lengths.get)
        df_forecast = pd.read_excel(xls, sheet_name=longest_sheet, header=None)
    
        # 检测 header 行：前三行中有“产品型号”者为 header
        header_row = None
        for i in range(3):
            if any("产品型号" in str(cell) for cell in df_forecast.iloc[i]):
                header_row = i
                break
        if header_row is None:
            for i in range(3):
                if any(re.search(r"\d{1,2}月预测", str(cell)) for cell in df_forecast.iloc[i]):
                    header_row = i
                    break
    
        if header_row is not None:
            df_forecast.columns = df_forecast.iloc[header_row]
            df_forecast = df_forecast.iloc[header_row + 1:].reset_index(drop=True)
    
            # ✅ 重命名第2列为“品名”（防止字段名异常）
            if df_forecast.shape[1] >= 2:
                df_forecast.columns = list(df_forecast.columns)
                df_forecast.columns.values[1] = "品名"
        else:
            st.warning(f"⚠️ 无法识别预测文件 `{filename}` 的 header，已跳过")
            continue
    
        # 👀 显示预测数据
        st.write(f"📁 读取到的预测文件 `{filename}`：", df_forecast.head())
    
        df_forecast = df_forecast.rename(columns=lambda x: str(x).strip())
        if "品名" not in df_forecast.columns:
            st.warning(f"⚠️ 预测文件 `{filename}` 缺少“品名”列，已跳过")
            continue
    
        df_forecast = df_forecast[["品名"]].copy()
        df_forecast["品名"] = df_forecast["品名"].astype(str).str.strip()
    
        # 替换新旧料号
        df_forecast, _ = apply_all_name_replacements(
            df_forecast,
            mapping_new,
            mapping_sub,
            sheet_name="预测",
            field_mappings=FIELD_MAPPINGS,
        )
        all_parts.append(df_forecast)


    # 3️⃣ 处理订单文件（Sheet）
    df_order = pd.read_excel(order_file, sheet_name="Sheet")
    st.write("📄 读取到的订单数据：", df_order.head())
    
    if "晶圆品名" not in df_order.columns:
        st.error("❌ 订单文件中缺少“晶圆品名”字段，请检查 Sheet 表格。")
        st.stop()
    
    df_order["晶圆品名"] = df_order["晶圆品名"].astype(str).str.strip()
    df_order, _ = apply_all_name_replacements(
        df_order, mapping_new, mapping_sub, "订单", FIELD_MAPPINGS
    )
    all_parts.append(df_order[["晶圆品名"]].rename(columns={"晶圆品名": "品名"}))
    
    
    # 4️⃣ 处理出货文件（原表）
    df_sales = pd.read_excel(sales_file, sheet_name="原表")
    st.write("📄 读取到的出货数据：", df_sales.head())
    
    if "品名" not in df_sales.columns:
        st.error("❌ 出货文件中缺少“品名”字段，请检查 原表 表格。")
        st.stop()
    
    df_sales["品名"] = df_sales["品名"].astype(str).str.strip()
    df_sales, _ = apply_all_name_replacements(
        df_sales, mapping_new, mapping_sub, "出货", FIELD_MAPPINGS
    )
    all_parts.append(df_sales[["品名"]])

    # 5️⃣ 合并去重品名并进行再次统一替换
    combined_names = pd.concat(all_parts, ignore_index=True)
    all_names = combined_names["品名"].dropna().drop_duplicates().reset_index(drop=True)
    replaced_names = replace_all_names_with_mapping(all_names, mapping_new, mapping_sub)

    # 6️⃣ 构造总表：晶圆 + 规格 + 品名，优先从 mapping 表中取
    mapping_dict = mapping_new.set_index("新品名")[["新晶圆", "新规格"]].copy()
    mapping_dict.columns = ["晶圆", "规格"]
    
    df_final = pd.DataFrame({"品名": replaced_names})
    df_final = df_final.merge(mapping_dict, how="left", left_on="品名", right_index=True)
    
    # 🧽 清理展示用 DataFrame，防止 Arrow 错误
    df_display = df_final.copy()
    df_display.columns = df_display.columns.map(str)
    for col in df_display.columns:
        df_display[col] = df_display[col].astype(str)
    
    # 显示当前初步结果
    st.write("🔎 替换后的主品名表（含规格与晶圆）前几行：", df_display.head())
    
    # 从订单或出货中补充缺失规格/晶圆
    missing_spec = df_final["规格"].isna()
    if missing_spec.any():
        # 合并出货和订单字段（品名、规格、晶圆）
        alt_spec = (
            pd.concat([df_order.rename(columns={"晶圆品名": "品名"}), df_sales], ignore_index=True)
            .dropna(subset=["品名"])
            .drop_duplicates(subset=["品名"])  # 🛡️ 确保唯一
            [["品名", "规格", "晶圆"]]
        )
    
        # 🔐 断言合并前唯一性
        assert alt_spec["品名"].is_unique, "❌ alt_spec 中品名不是唯一的"
    
        # 合并补规格
        df_final = df_final.merge(alt_spec, on="品名", how="left", suffixes=("", "_alt"))
        df_final["规格"] = df_final["规格"].fillna(df_final["规格_alt"])
        df_final["晶圆"] = df_final["晶圆"].fillna(df_final["晶圆_alt"])
        df_final = df_final.drop(columns=["规格_alt", "晶圆_alt"])
    
    # ✅ 最终结果展示
    df_final = df_final[["晶圆", "规格", "品名"]]
    
    # 展示前再次清理，确保安全显示
    df_final_display = df_final.copy()
    df_final_display.columns = df_final_display.columns.map(str)
    for col in df_final_display.columns:
        df_final_display[col] = df_final_display[col].astype(str)
    
    st.success("✅ 总品名表生成成功！")
    st.dataframe(df_final_display, use_container_width=True)
    
    # 📥 下载
    csv = df_final.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("📥 下载结果 CSV", csv, file_name="总品名列表.csv")
