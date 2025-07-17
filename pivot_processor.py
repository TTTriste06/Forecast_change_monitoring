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
        from info_extract import extract_all_year_months, fill_forecast_data, fill_order_data, fill_sales_data, highlight_by_detecting_column_headers
        from name_utils import extract_unique_rows_from_all_sources, build_main_df
        from forecast_utils import load_forecast_files


        forecast_dfs = load_forecast_files(forecast_files)
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)


        main_df = build_main_df(forecast_dfs, order_file, sales_file, mapping_new, mapping_sub)

        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }

        # 替换所有 forecast_dfs 中的品名为新料号或替代料号
        def apply_mapping_to_all_forecasts(forecast_dfs: dict[str, pd.DataFrame], mapping_new, mapping_sub, field_mapping: dict) -> dict[str, pd.DataFrame]:
            from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping
        
            mapped_dfs = {}
            for name, df in forecast_dfs.items():
                df_mapped, _ = apply_mapping_and_merge(df.copy(), mapping_new, field_mapping)
                df_mapped, _ = apply_extended_substitute_mapping(df_mapped, mapping_sub, field_mapping)
                mapped_dfs[name] = df_mapped
        
            return mapped_dfs


        forecast_dfs = apply_mapping_to_all_forecasts(forecast_dfs, mapping_new, mapping_sub, FIELD_MAPPINGS["forecast"])
        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])
        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])

        
        all_months = extract_all_year_months(forecast_file, order_file, sales_file)

        for ym in all_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        main_df = fill_forecast_data(main_df, forecast_file)
        main_df = fill_order_data(main_df, order_file, all_months)
        main_df = fill_sales_data(main_df, sales_file, all_months)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
            wb = writer.book

            highlight_by_detecting_column_headers(ws)

            from openpyxl.styles import Alignment, PatternFill
            from openpyxl.utils import get_column_letter

            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            fill_colors = [
                "FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFE599"
            ]

            col = 4
            for i, ym in enumerate(all_months):
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)

                ws.cell(row=2, column=col).value = "预测"
                ws.cell(row=2, column=col + 1).value = "订单"
                ws.cell(row=2, column=col + 2).value = "出货"

                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], end_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                for j in range(col, col + 3):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill

                col += 3

            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 10

            forecast_file.to_excel(writer, index=False, sheet_name="原始-预测")
            order_file.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_file.to_excel(writer, index=False, sheet_name="原始-出货")

        output.seek(0)
        return main_df, output
