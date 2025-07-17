import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import re

class PivotProcessor:
    def process(self, forecast_file, order_file, sales_file, mapping_file):
        from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping, split_mapping_data
        from info_extract import extract_all_year_months, fill_forecast_data, fill_order_data, fill_sales_data, highlight_by_detecting_column_headers
        from name_utils import extract_unique_rows_from_all_sources
        
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)

        # ✅ 自动构造 main_df，不再依赖 template_df
        def build_main_df():
            def extract_and_standardize(df, col_name):
                df = df[[col_name]].copy()
                df.columns = ["品名"]
                df["品名"] = df["品名"].astype(str).str.strip()
                return df
        
            df_forecast_names = extract_and_standardize(forecast_file.iloc[:, [1]], forecast_file.columns[1])
            df_order_names = extract_and_standardize(order_file, "品名")
            df_sales_names = extract_and_standardize(sales_file, "品名")
        
            all_names = pd.concat([df_forecast_names, df_order_names, df_sales_names], ignore_index=True).drop_duplicates()
        
            all_names, _ = apply_mapping_and_merge(all_names, mapping_new, {"品名": "品名"})
            all_names, _ = apply_extended_substitute_mapping(all_names, mapping_sub, {"品名": "品名"})
        
            mapping_clean = mapping_new.copy()
            mapping_clean["新品名"] = mapping_clean["新品名"].astype(str).str.strip()
            mapping_clean["新晶圆"] = mapping_clean["新晶圆"].astype(str).str.strip()
            mapping_clean["新规格"] = mapping_clean["新规格"].astype(str).str.strip()
        
            merged = all_names.merge(
                mapping_clean[["新品名", "新晶圆", "新规格"]],
                left_on="品名",
                right_on="新品名",
                how="left"
            ).drop(columns=["新品名"], errors="ignore")
        
            merged = merged.rename(columns={"新晶圆": "晶圆品名", "新规格": "规格"})
            merged["晶圆品名"] = merged["晶圆品名"].fillna("")
            merged["规格"] = merged["规格"].fillna("")
        
            return merged[["晶圆品名", "规格", "品名"]]
        
        main_df = build_main_df()


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

            self.add_detail_link_and_sheets(
                wb=wb,
                ws_main=ws,
                df_order=order_file,
                df_sales=sales_file,
                df_forecast=forecast_file,
                all_months=all_months
            )

        output.seek(0)
        return main_df, output

    def safe_sheet_name(self, name: str) -> str:
        name = re.sub(r"[\\/*?:\[\]]", "", name)
        return name

    def add_detail_link_and_sheets(self, wb, ws_main, df_order, df_sales, df_forecast, all_months):
        created_sheets = set()

        for row in range(3, ws_main.max_row + 1):
            item_name = ws_main.cell(row=row, column=3).value
            for i, ym in enumerate(all_months):
                for offset, df, date_col, value_col, prefix in [
                    (0, df_forecast, None, f"{ym}-预测", "预测"),
                    (1, df_order, "客户要求交期", "订单数量", "订单"),
                    (2, df_sales, "交易日期", "数量", "出货"),
                ]:
                    col = 4 + i * 3 + offset
                    cell = ws_main.cell(row=row, column=col)
                    val = cell.value
                    if not val or float(val) == 0:
                        continue

                    raw_name = f"{prefix}-{ym}-{item_name}"
                    raw_name = self.safe_sheet_name(raw_name)
                    sheet_name = raw_name[:31]
                    suffix = 1
                    base_name = sheet_name
                    while sheet_name in created_sheets:
                        sheet_name = f"{base_name[:27]}-{suffix}"
                        suffix += 1

                    cell.value = f'=HYPERLINK("#\'{sheet_name}\'!A1", "{cell.value}")'

                    cell.font = Font(underline="single", color="0000FF")

                    created_sheets.add(sheet_name)

                    if prefix == "预测":
                        # 只匹配生产料号（品名），不再按月份列筛选
                        df_filtered = df[df["生产料号"] == item_name]
                    else:
                        df_temp = df.copy()
                        if date_col in df_temp.columns:
                            df_temp[date_col] = pd.to_datetime(df_temp[date_col], errors="coerce")
                            df_temp["年月"] = df_temp[date_col].dt.to_period("M").astype(str)
                            df_filtered = df_temp[(df_temp["品名"] == item_name) & (df_temp["年月"] == ym)]
                        else:
                            df_filtered = pd.DataFrame()

                    ws_detail = wb.create_sheet(sheet_name)

                    # 添加“返回主页”按钮
                    return_cell = ws_detail.cell(row=1, column=1)
                    return_cell.value = '=HYPERLINK("#预测分析!A1", "⬅ 返回主页")'
                    return_cell.font = Font(underline="single", color="0000FF")
                    
                    start_row = 2  # 从第 2 行开始写明细数据
                    if not df_filtered.empty:
                        for r_idx, row_data in enumerate(dataframe_to_rows(df_filtered, index=False, header=True), start=start_row):
                            for c_idx, val in enumerate(row_data, start=1):
                                ws_detail.cell(r_idx, c_idx, value=val)
                    else:
                        ws_detail.cell(start_row, 1, value="无匹配数据")
                    
