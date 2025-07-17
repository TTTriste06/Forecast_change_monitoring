import pandas as pd
import streamlit as st
import re
from datetime import datetime
from io import BytesIO

def extract_forecast_generation_date(self, filename: str) -> str:
    """从文件名中提取生成日期，返回格式为 yyyy-mm"""
    match = re.search(r"_(\d{8})", filename)
    if match:
        date_str = match.group(1)
        try:
            dt = datetime.strptime(date_str, "%Y%m%d")
            return dt.strftime("%Y-%m")
        except ValueError:
            pass
    return "unknown"

def extract_forecast_data(self, file: BytesIO) -> tuple[pd.DataFrame, str]:
    """读取预测文件最长的sheet，并自动识别header行（包含‘产品型号’），返回DataFrame和生成年月"""
    xls = pd.ExcelFile(file)
    max_len = 0
    selected_sheet = None

    for sheet in xls.sheet_names:
        df = xls.parse(sheet, header=None)
        if df.shape[0] > max_len:
            max_len = df.shape[0]
            selected_sheet = sheet

    df_raw = xls.parse(selected_sheet, header=None)

    header_row_idx = None
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("产品型号").any():
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError("❌ 未在文件中识别到包含‘产品型号’的header行")

    df = pd.read_excel(file, sheet_name=selected_sheet, header=header_row_idx)
    df.columns.values[1] = "品名"
    return df, selected_sheet  # sheet名可选作备份信息

def parse_forecast_months(self, forecast_df: pd.DataFrame, base_year: int) -> dict:
    """
    输入包含“x月预测”列的df，返回一个dict：
    {"yyyy-mm": 原始列名}，自动判断跨年
    """
    pattern = re.compile(r"^(\d{1,2})月预测$")
    col_map = {}

    start_year = base_year
    last_month = 0
    for col in forecast_df.columns:
        match = pattern.match(col)
        if match:
            month = int(match.group(1))
            if last_month and month < last_month:
                start_year += 1
            last_month = month
            ym = f"{start_year}-{month:02d}"
            col_map[ym] = col
    return col_map

def append_multi_forecast_columns(
    self,
    main_df: pd.DataFrame,
    forecast_df: pd.DataFrame,
    col_map: dict,
    label: str
) -> pd.DataFrame:
    """
    在 main_df 中添加来自 forecast_df 的预测列，列名为 “label（yyyy-mm）”，
    forecast_df 中“生产料号”为品名，col_map 为 {yyyy-mm: 原始列名}
    """
    forecast_df["生产料号"] = forecast_df["生产料号"].astype(str).str.strip()
    forecast_df = forecast_df.rename(columns={"生产料号": "品名"})
    main_df["品名"] = main_df["品名"].astype(str).str.strip()

    for ym, orig_col in col_map.items():
        new_col = f"{label}（{ym}）"
        if new_col not in main_df.columns:
            main_df[new_col] = 0
        for i, row in main_df.iterrows():
            name = row["品名"]
            val = forecast_df.loc[forecast_df["品名"] == name, orig_col]
            if not val.empty:
                main_df.at[i, new_col] = val.values[0]
    return main_df


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

def load_forecast_files(uploaded_files):
    """
    加载多个上传的预测文件，返回两个列表：
    - forecast_dfs: 所有 DataFrame
    - forecast_files: 原始文件对象（用于获取 .name）
    """
    forecast_dfs = []
    forecast_files = []

    for file in uploaded_files:
        df = pd.read_excel(file, sheet_name=0)
        forecast_dfs.append(df)
        forecast_files.append(file)

    return forecast_dfs, forecast_files
