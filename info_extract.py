import re
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def extract_year_month_from_filename(filename: str) -> str:
    match = re.search(r"(20\d{2})(\d{2})\d{2}", filename)
    if match:
        year, month = match.group(1), match.group(2)
        return f"{year}-{month}"
    return None

def extract_all_year_months(df_forecast, df_order, df_sales, forecast_year = None):
    # 1. 从 forecast header 提取 x月预测 列中的月份
    if forecast_year is None:
        forecast_year = datetime.today().year
    month_pattern = re.compile(r"(\d{1,2})月预测")
    forecast_months = []
    for col in df_forecast.columns:
        match = month_pattern.match(str(col))
        if match:
            month = match.group(1).zfill(2)
            forecast_months.append(f"{forecast_year}-{month}")  # ✅ 根据需要调整年份

    # 2. 从 order 文件第 B 列（假设是“订单日期”）
    order_date_col = df_order.columns[11]
    df_order[order_date_col] = pd.to_datetime(df_order[order_date_col],  format="%Y-%m", errors="coerce")
    order_months = (
        df_order[order_date_col]
        .dropna()
        .dt.to_period("M")
        .astype(str)
        .loc[lambda x: x != "NaT"]
        .unique()
        .tolist()
    )

    # 3. 从 sales 文件第 F 列（假设是“交易日期”）
    sales_date_col = df_sales.columns[5]
    df_sales[sales_date_col] = pd.to_datetime(df_sales[sales_date_col], format="%Y-%m", errors="coerce")
    sales_months = (
        df_sales[sales_date_col]
        .dropna()
        .dt.to_period("M")
        .astype(str)
        .loc[lambda x: x != "NaT"]
        .unique()
        .tolist()
    )

    # 合并并去重
    all_months = sorted(set(forecast_months + order_months + sales_months))

    # 生成从最小到最大之间的所有月份
    if all_months:
        min_month = pd.Period(min(all_months), freq="M")
        max_month = pd.Period(max(all_months), freq="M")
        full_months = [str(p) for p in pd.period_range(min_month, max_month, freq="M")]
    else:
        full_months = []
    
    return full_months

def fill_forecast_data(main_df: pd.DataFrame, df_forecast: pd.DataFrame) -> pd.DataFrame:
    """
    从 df_forecast 中提取“生产料号”作为品名，解析“x月预测”列，填入 main_df 中“yyyy-mm-预测”字段。
    默认使用当前年份。
    """
    # 使用当前年份
    current_year = datetime.today().year

    # 统一格式处理
    df_forecast["生产料号"] = df_forecast["生产料号"].astype(str).str.strip()
    df_forecast["品名"] = df_forecast["生产料号"]

    # 正则提取“x月预测”字段
    month_pattern = re.compile(r"^\s*(\d{1,2})月\s*预测\s*$")
    forecast_cols = {
        f"{current_year}-{match.group(1).zfill(2)}": col
        for col in df_forecast.columns
        if (match := month_pattern.match(str(col)))
    }

    for ym, month_col in forecast_cols.items():
        target_col = f"{ym}-预测"
        if target_col not in main_df.columns:
            main_df[target_col] = 0  # 若不存在则新建

        # 按“品名”聚合预测数据并写入
        forecast_series = (
            df_forecast.groupby("品名")[month_col]
            .sum(min_count=1)
        )

        main_df[target_col] = main_df["品名"].map(forecast_series).fillna(0)

    return main_df



def fill_order_data(main_df, df_order, forecast_months):
    """
    将订单数据按“订单日期”和“品名”聚合并填入 main_df 中每月的“订单”列。
    
    参数：
    - main_df: 主计划 DataFrame，需包含“品名”列
    - df_order: 上传的未交订单 DataFrame，包含“订单日期”和“品名”
    - forecast_months: 所有涉及的 yyyy-mm 字符串列表
    """
    df_order = df_order.copy()

    # 确保日期字段为 datetime 类型
    df_order["客户要求交期"] = pd.to_datetime(df_order["客户要求交期"], errors="coerce")
    df_order["年月"] = df_order["客户要求交期"].dt.to_period("M").astype(str)

    # 数值字段清洗
    df_order["订单数量"] = pd.to_numeric(df_order["订单数量"], errors="coerce").fillna(0)

    # 聚合出每品名每月的订单量
    grouped = (
        df_order.groupby(["品名", "年月"])["订单数量"]
        .sum()
        .unstack()
        .fillna(0)
    )

    for ym in forecast_months:
        colname = f"{ym}-订单"
        if colname in main_df.columns and ym in grouped.columns:
            main_df[colname] = main_df["品名"].map(grouped[ym]).fillna(0)

    return main_df

def fill_sales_data(main_df, df_sales, forecast_months):
    """
    将出货数据按“交易日期”和“品名”聚合并填入 main_df 中每月的“出货”列。
    
    参数：
    - main_df: 主计划 DataFrame，需包含“品名”列
    - df_sales: 出货明细 DataFrame，包含“交易日期”和“品名”
    - forecast_months: 所有涉及的 yyyy-mm 字符串列表
    """
    df_sales = df_sales.copy()

    # 确保交易日期为 datetime
    df_sales["交易日期"] = pd.to_datetime(df_sales["交易日期"], errors="coerce")
    df_sales["年月"] = df_sales["交易日期"].dt.to_period("M").astype(str)

    # 数值字段清洗
    df_sales["数量"] = pd.to_numeric(df_sales["数量"], errors="coerce").fillna(0)

    # 聚合：每品名每月出货数量
    grouped = (
        df_sales.groupby(["品名", "年月"])["数量"]
        .sum()
        .unstack()
        .fillna(0)
    )

    for ym in forecast_months:
        colname = f"{ym}-出货"
        if colname in main_df.columns and ym in grouped.columns:
            main_df[colname] = main_df["品名"].map(grouped[ym]).fillna(0)

    return main_df

def highlight_by_detecting_column_headers(ws):
    """
    自动识别表头第二行中连续的“预测/订单”列对，并对值为：预测>0且订单=0 的单元格标红。
    """
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    max_col = ws.max_column
    max_row = ws.max_row

    header = [cell.value for cell in ws[2]]

    # 遍历所有列找出“预测-订单”成对列索引
    column_pairs = []
    for i in range(len(header) - 1):
        name1 = str(header[i]).strip()
        name2 = str(header[i + 1]).strip()
        if name1.endswith("预测") and name2.endswith("订单"):
            column_pairs.append((i + 1, i + 2))  # openpyxl列从1开始

    # 遍历每行，检查成对列
    for row in range(3, ws.max_row + 1):
        for forecast_col, order_col in column_pairs:
            cell_forecast = ws.cell(row=row, column=forecast_col)
            cell_order = ws.cell(row=row, column=order_col)

            try:
                val_forecast = float(cell_forecast.value or 0)
                val_order = float(cell_order.value or 0)
            except:
                continue

            if val_forecast > 0 and val_order == 0:
                cell_forecast.fill = red_fill
                cell_order.fill = red_fill
