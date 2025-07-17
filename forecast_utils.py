import pandas as pd
import streamlit as st
import re

def merge_forecast_columns(forecast_dfs: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    合并多个预测文件，按“品名”横向合并所有预测列。
    返回值为：包含“品名”列 + 多个“yyyy-mm的预测（生成年月）”列
    """
    all_parts = []
    for file_name, df in forecast_dfs.items():
        try:
            forecast_parts = parse_forecast_columns(df, file_name)
            for subdf in forecast_parts.values():
                all_parts.append(subdf)
        except Exception as e:
            st.warning(f"⚠ 处理 {file_name} 失败：{e}")

    # 统一合并
    if not all_parts:
        return pd.DataFrame(columns=["品名"])
    merged = all_parts[0]
    for part in all_parts[1:]:
        merged = pd.merge(merged, part, on="品名", how="outer")

    return merged.fillna(0)


def parse_forecast_columns(df: pd.DataFrame, file_name: str) -> dict[str, pd.Series]:
    """
    从预测表中提取“x月预测”列，并转换为“yyyy-mm的预测（生成年月）”格式
    """
    # 从文件名中提取生成日期
    match = re.search(r"(\d{8})", file_name)
    if not match:
        raise ValueError(f"文件名中未找到 8 位日期：{file_name}")
    gen_date = pd.to_datetime(match.group(1), format="%Y%m%d")
    gen_year = gen_date.year
    gen_month = gen_date.month
    gen_ym_str = gen_date.strftime("%Y-%m")

    forecast_cols = {}
    month_pattern = re.compile(r"^(\d{1,2})月预测$")

    for col in df.columns:
        if isinstance(col, str):
            match = month_pattern.match(col.strip())
            if match:
                month_num = int(match.group(1))
                # 决定年份（跨年判断）
                if month_num >= gen_month:
                    year = gen_year
                else:
                    year = gen_year + 1
                ym_key = f"{year}-{str(month_num).zfill(2)}"
                new_col_name = f"{ym_key}的预测（{gen_ym_str}）"
                forecast_cols[new_col_name] = df[[col, "品名"]].rename(columns={col: new_col_name})

    return forecast_cols  # 返回列名 → dataframe with 品名 + 单列


def load_forecast_files(files: dict) -> dict[str, pd.DataFrame]:
    """
    对上传的多个预测 Excel 文件执行以下操作：
    - 找到每个文件中最长的 sheet
    - 自动识别 header 行（含“产品型号”的那一行）
    - 将第二列统一命名为“品名”
    - 用 st.write 打印每个文件读取结果
    返回值：dict[file_name -> cleaned DataFrame]
    """
    result = {}

    for uploaded_file in files:
        file_name = uploaded_file.name
        try:
            xls = pd.ExcelFile(uploaded_file)
            longest_sheet = max(xls.sheet_names, key=lambda name: pd.read_excel(xls, sheet_name=name).shape[0])
            df_raw = pd.read_excel(xls, sheet_name=longest_sheet, header=None)

            # 自动识别 header 行：包含“产品型号”的行
            header_row_idx = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("产品型号").any(), axis=1)].index
            if header_row_idx.empty:
                st.warning(f"⚠ 文件 {file_name} 中未找到包含“产品型号”的表头行，跳过")
                continue

            header_row = header_row_idx[0]
            df = pd.read_excel(xls, sheet_name=longest_sheet, header=header_row)

            # 统一第二列为“品名”
            if df.shape[1] >= 2:
                df.columns.values[1] = "品名"

            st.write(f"📄 读取成功：{file_name}（使用 sheet：{longest_sheet}，header 行：第 {header_row+1} 行）")
            st.dataframe(df)

            result[file_name] = df

        except Exception as e:
            st.error(f"❌ 无法读取文件 {file_name}: {e}")

    return result
