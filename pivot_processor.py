import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import re
from datetime import datetime
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


class PivotProcessor:
    def process(self, forecast_files, order_file, sales_file, mapping_file):
        from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping, split_mapping_data
        from info_extract import extract_all_year_months, fill_order_data, fill_sales_data, highlight_by_detecting_column_headers
        from name_utils import build_main_df
        from forecast_utils import load_forecast_files

        # ✅ 加载原始预测文件
        forecast_dfs = load_forecast_files(forecast_files)
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)
        main_df = build_main_df(forecast_dfs, order_file, sales_file, mapping_new, mapping_sub)

        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }

        # ✅ 替换预测中品名
        def apply_mapping_to_all_forecasts(forecast_dfs: dict[str, pd.DataFrame], mapping_new, mapping_sub) -> dict[str, pd.DataFrame]:
            from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping
            mapped_dfs = {}
            for name, df in forecast_dfs.items():
                if df.shape[1] < 2:
                    continue
                second_col = df.columns[1]
                field_mapping = {"品名": second_col}
                try:
                    df_mapped, _ = apply_mapping_and_merge(df.copy(), mapping_new, field_mapping)
                    df_mapped, _ = apply_extended_substitute_mapping(df_mapped, mapping_sub, field_mapping)
                    mapped_dfs[name] = df_mapped
                except KeyError as e:
                    raise ValueError(f"❌ `{name}` 缺失列: {e}. 实际列: {df.columns.tolist()}") from e
            return mapped_dfs

        forecast_dfs = apply_mapping_to_all_forecasts(forecast_dfs, mapping_new, mapping_sub)
        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])
        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])

        # ✅ 提取所有月份（订单/出货用）
        all_months = extract_all_year_months(forecast_dfs, order_file, sales_file)
        for ym in all_months:
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        # ✅ 填充所有预测数据（独立列）
        def extract_file_date(file_name: str) -> str:
            match = re.search(r"(\d{8})", file_name)
            return match.group(1) if match else "00000000"

        def detect_header_row(df: pd.DataFrame) -> int:
            for i, row in df.iterrows():
                if any(isinstance(cell, str) and "产品型号" in str(cell) for cell in row):
                    return i
            return 0

        def standardize_column_name(forecast_col: str, file_date: str) -> str:
            month_match = re.match(r"^(\d{1,2})月预测$", forecast_col.strip())
            alt_match = re.match(r"^(\d{1,2})月预测\d*$", forecast_col.strip())
            if month_match or alt_match:
                month = (month_match or alt_match).group(1).zfill(2)
            else:
                return f"{file_date}-{forecast_col.strip()}"
            file_year = file_date[:4]
            file_month = file_date[4:6]
            return f"{file_year}-{month}的预测（{file_year}-{file_month}生成）"

        def fill_forecast_data(main_df: pd.DataFrame, forecast_dfs: dict[str, pd.DataFrame]) -> pd.DataFrame:
            for file_name, df in forecast_dfs.items():
                file_date = extract_file_date(file_name)
                name_col = "生产料号" if "生产料号" in df.columns else (df.columns[1] if df.shape[1] >= 2 else None)
                if name_col is None:
                    continue
                df["品名"] = df[name_col].astype(str).str.strip()
                for col in df.columns:
                    if isinstance(col, str) and "预测" in col:
                        new_col = standardize_column_name(col, file_date)
                        if new_col not in main_df.columns:
                            main_df[new_col] = 0
                        forecast_series = df[["品名", col]].dropna().groupby("品名")[col].sum(min_count=1)
                        main_df[new_col] = main_df["品名"].map(forecast_series).fillna(0)
            return main_df

        main_df = fill_forecast_data(main_df, forecast_dfs)
        main_df = fill_order_data(main_df, order_file, all_months)
        main_df = fill_sales_data(main_df, sales_file, all_months)

        # ✅ 写入 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=2)
            ws = writer.sheets["预测分析"]
            wb = writer.book

            highlight_by_detecting_column_headers(ws)

            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            from collections import defaultdict
            fill_colors = ["FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFE599"]
            month_grouped_columns = defaultdict(lambda: {"预测": [], "订单": None, "出货": None})

            # 分组所有列
            for col in main_df.columns:
                if re.match(r"\\d{4}-\\d{2}的预测（\\d{4}-\\d{2}生成）", str(col)):
                    ym = col[:7]
                    month_grouped_columns[ym]["预测"].append(col)
                elif col.endswith("-订单") and re.match(r"\\d{4}-\\d{2}-订单", col):
                    ym = col[:7]
                    month_grouped_columns[ym]["订单"] = col
                elif col.endswith("-出货") and re.match(r"\\d{4}-\\d{2}-出货", col):
                    ym = col[:7]
                    month_grouped_columns[ym]["出货"] = col

            col = 4
            for i, ym in enumerate(sorted(month_grouped_columns)):
                sub_cols = []
                if month_grouped_columns[ym]["预测"]:
                    sub_cols.extend(month_grouped_columns[ym]["预测"])
                if month_grouped_columns[ym]["订单"]:
                    sub_cols.append(month_grouped_columns[ym]["订单"])
                if month_grouped_columns[ym]["出货"]:
                    sub_cols.append(month_grouped_columns[ym]["出货"])

                span = len(sub_cols)
                if span == 0:
                    continue

                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)

                for j, sub_col in enumerate(sub_cols):
                    ws.cell(row=2, column=col + j).value = sub_col
                    ws.cell(row=2, column=col + j).alignment = Alignment(horizontal="center", vertical="center")
                    fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], end_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                    ws.cell(row=1, column=col + j).fill = fill
                    ws.cell(row=2, column=col + j).fill = fill

                col += span

            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 10

        output.seek(0)
        return main_df, output
