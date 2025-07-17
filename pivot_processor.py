import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO

class PivotProcessor:
    def process(self, forecast_files, order_file, sales_file, mapping_file):
        from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping, split_mapping_data
        from info_extract import extract_all_year_months, fill_forecast_data, fill_order_data, fill_sales_data, highlight_by_detecting_column_headers
        from name_utils import build_main_df
        from forecast_utils import (
            extract_forecast_generation_date,
            parse_forecast_months,
            append_multi_forecast_columns,
            merge_forecast_columns,
            load_forecast_files
        )

        # === 1. 读取预测文件并合并
        forecast_dfs = load_forecast_files(forecast_files)
        forecast_file = pd.concat(forecast_dfs, ignore_index=True)

        # === 2. 拆分 mapping 表
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)

        # === 3. 统一替换 品名
        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }

        forecast_file, _ = apply_mapping_and_merge(forecast_file, mapping_new, FIELD_MAPPINGS["forecast"])
        forecast_file, _ = apply_extended_substitute_mapping(forecast_file, mapping_sub, FIELD_MAPPINGS["forecast"])

        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])

        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])

        # === 4. 构建主表品名清单
        main_df = build_main_df(forecast_file, order_file, sales_file, mapping_new, mapping_sub)

        # === 5. 提取所有预测月份、拼接完整列名
        forecast_file = append_multi_forecast_columns(forecast_file, forecast_files)
        forecast_months = parse_forecast_months(forecast_file.columns)

        # === 6. 填入预测、订单、出货数据
        main_df = fill_forecast_data(main_df, forecast_file)
        all_months = extract_all_year_months(forecast_file, order_file, sales_file)

        main_df = fill_order_data(main_df, order_file, all_months)
        main_df = fill_sales_data(main_df, sales_file, all_months)

        # === 7. 输出为 Excel 带格式
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
            wb = writer.book

            highlight_by_detecting_column_headers(ws)

            # === 合并列标题 “晶圆品名”~“品名”
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            # === 构建分组列（每月一组：预测、订单、出货）
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

                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)],
                                   end_color=fill_colors[i % len(fill_colors)],
                                   fill_type="solid")
                for j in range(col, col + 3):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill
                col += 3

            # 自动列宽
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = max((len(str(cell.value)) if cell.value else 0) for cell in column_cells)
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 4

            # === 写入原始数据
            for i, df in enumerate(forecast_dfs):
                name = forecast_files[i].name if hasattr(forecast_files[i], "name") else f"预测源{i+1}"
                safe_name = f"预测源{i+1}"
                df.to_excel(writer, index=False, sheet_name=safe_name[:31])

            order_file.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_file.to_excel(writer, index=False, sheet_name="原始-出货")

        output.seek(0)
        return main_df, output
