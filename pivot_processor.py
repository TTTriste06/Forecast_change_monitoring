import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import re

class PivotProcessor:
    def process(self, forecast_files, order_file, sales_file, mapping_file):
        from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping, split_mapping_data
        from info_extract import fill_order_data, fill_sales_data, highlight_by_detecting_column_headers
        from name_utils import build_main_df
        from forecast_utils import load_forecast_files

        def extract_forecast_generation_date(filename: str) -> str:
            match = re.search(r'_(\d{8})', filename)
            if not match:
                raise ValueError(f"❌ 文件名中未找到日期信息: {filename}")
            dt = pd.to_datetime(match.group(1), format="%Y%m%d")
            return dt.strftime("%Y-%m")

        def extract_forecast_data(df: pd.DataFrame, generation_month: str):
            df = df.rename(columns={"生产料号": "品名"}).copy()
            df["品名"] = df["品名"].astype(str).str.strip()

            base_year, base_month = map(int, generation_month.split("-"))
            month_map = {}

            for col in df.columns:
                match = re.match(r"^(\d{1,2})月预测$", str(col))
                if match:
                    m = int(match.group(1))
                    year = base_year if m >= base_month else base_year + 1
                    ym = f"{year}-{str(m).zfill(2)}"
                    month_map[ym] = col
            return df, month_map

        def append_multi_forecast_columns(main_df, forecast_dfs, forecast_files):
            all_months = set()
            forecast_origin_map = {}

            for df, file in zip(forecast_dfs, forecast_files):
                gen_month = extract_forecast_generation_date(file.name)
                df, ym_map = extract_forecast_data(df, gen_month)

                for ym, col in ym_map.items():
                    label = f"{gen_month}的预测（{ym}）"
                    if label not in main_df.columns:
                        main_df[label] = 0
                    forecast_origin_map.setdefault(ym, []).append(label)
                    all_months.add(ym)

                    df_part = df[["品名", col]].copy()
                    df_part.columns = ["品名", "预测值"]
                    main_df = main_df.merge(df_part, on="品名", how="left")
                    main_df[label] = main_df["预测值"].fillna(0).astype(float)
                    main_df.drop(columns=["预测值"], inplace=True)

            return main_df, sorted(all_months), forecast_origin_map

        # === 主流程 ===
        forecast_dfs = load_forecast_files(forecast_files)
        forecast_file = pd.concat(forecast_dfs, ignore_index=True)

        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)

        for i in range(len(forecast_dfs)):
            forecast_dfs[i], _ = apply_mapping_and_merge(forecast_dfs[i], mapping_new, {"品名": "生产料号"})
            forecast_dfs[i], _ = apply_extended_substitute_mapping(forecast_dfs[i], mapping_sub, {"品名": "生产料号"})

        forecast_file, _ = apply_mapping_and_merge(forecast_file, mapping_new, {"品名": "生产料号"})
        forecast_file, _ = apply_extended_substitute_mapping(forecast_file, mapping_sub, {"品名": "生产料号"})
        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, {"品名": "品名"})
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, {"品名": "品名"})
        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, {"品名": "品名"})
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, {"品名": "品名"})

        main_df = build_main_df(forecast_file, order_file, sales_file, mapping_new, mapping_sub)

        # 添加多个预测列
        main_df, forecast_months, forecast_origin_map = append_multi_forecast_columns(main_df, forecast_dfs, forecast_files)

        # 添加订单和出货
        main_df = fill_order_data(main_df, order_file, forecast_months)
        main_df = fill_sales_data(main_df, sales_file, forecast_months)

        # 写入 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=2)
            wb = writer.book
            ws = writer.sheets["预测分析"]

            # 样式：合并标题
            from openpyxl.styles import Alignment, PatternFill, Font
            from openpyxl.utils import get_column_letter

            # 基础字段
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            col = 4
            fill_colors = ["FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFE599"]
            for i, ym in enumerate(forecast_months):
                labels = forecast_origin_map.get(ym, [])
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + len(labels) + 1)

                # 一级标题
                cell = ws.cell(row=1, column=col)
                cell.value = ym
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

                # 二级标题
                for j, sub in enumerate(labels):
                    ws.cell(row=2, column=col + j).value = sub
                ws.cell(row=2, column=col + len(labels)).value = "订单"
                ws.cell(row=2, column=col + len(labels) + 1).value = "出货"

                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)],
                                   end_color=fill_colors[i % len(fill_colors)],
                                   fill_type="solid")
                for j in range(col, col + len(labels) + 2):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill

                col += len(labels) + 2

            # 自动列宽
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 8

            # 高亮
            highlight_by_detecting_column_headers(ws)

            # 附加原始表
            forecast_file.to_excel(writer, index=False, sheet_name="原始-预测")
            order_file.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_file.to_excel(writer, index=False, sheet_name="原始-出货")

        output.seek(0)
        return main_df, output
