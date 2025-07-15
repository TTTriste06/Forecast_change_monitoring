import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

class PivotProcessor:
    def process(self, forecast_files, order_df, sales_df, mapping_df):
        from mapping_utils import (
            apply_mapping_and_merge,
            apply_extended_substitute_mapping,
            split_mapping_data
        )
        from info_extract import (
            extract_all_year_months,
            fill_forecast_data,
            fill_order_data,
            fill_sales_data,
            highlight_by_detecting_column_headers
        )

        # 拆分新旧料号映射表
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

        # ========== 1. 读取所有预测文件中的最后一个 sheet ==========
        forecast_dfs = []
        for file in forecast_files:
            xls = pd.ExcelFile(file)
            last_sheet = xls.sheet_names[-1]
            df = xls.parse(last_sheet)
            forecast_dfs.append(df)

        # 合并所有预测 DataFrame（可选保留唯一）
        forecast_df = pd.concat(forecast_dfs, ignore_index=True)

        # ========== 2. 提取唯一的 品名/晶圆/规格 ==========
        def extract_unique_rows(df, rename_dict):
            df = df.rename(columns=rename_dict)
            required_cols = ["晶圆品名", "规格", "品名"]
            return df[required_cols].dropna().drop_duplicates()

        rename_forecast = {"生产料号": "品名"}  # 假设 forecast 用“生产料号”字段表示品名
        rename_order = {"品名": "品名"}
        rename_sales = {"品名": "品名"}

        forecast_part = extract_unique_rows(forecast_df.rename(columns=rename_forecast), {
            "品名": "品名", "晶圆": "晶圆品名", "规格": "规格"
        })
        order_part = extract_unique_rows(order_df, {
            "品名": "品名", "晶圆": "晶圆品名", "规格": "规格"
        })
        sales_part = extract_unique_rows(sales_df, {
            "品名": "品名", "晶圆": "晶圆品名", "规格": "规格"
        })

        main_df = pd.concat([forecast_part, order_part, sales_part]).drop_duplicates().reset_index(drop=True)

        # ========== 3. 映射料号 ==========
        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }

        forecast_df, _ = apply_mapping_and_merge(forecast_df, mapping_new, FIELD_MAPPINGS["forecast"])
        forecast_df, _ = apply_extended_substitute_mapping(forecast_df, mapping_sub, FIELD_MAPPINGS["forecast"])
        order_df, _ = apply_mapping_and_merge(order_df, mapping_new, FIELD_MAPPINGS["order"])
        order_df, _ = apply_extended_substitute_mapping(order_df, mapping_sub, FIELD_MAPPINGS["order"])
        sales_df, _ = apply_mapping_and_merge(sales_df, mapping_new, FIELD_MAPPINGS["sales"])
        sales_df, _ = apply_extended_substitute_mapping(sales_df, mapping_sub, FIELD_MAPPINGS["sales"])

        # ========== 4. 获取所有预测月份 ==========
        all_months = extract_all_year_months(forecast_df, order_df, sales_df)

        # 初始化列
        for ym in all_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        # ========== 5. 填充数据 ==========
        main_df = fill_forecast_data(main_df, forecast_df)
        main_df = fill_order_data(main_df, order_df, all_months)
        main_df = fill_sales_data(main_df, sales_df, all_months)

        # ========== 6. 输出 Excel ==========
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]

            highlight_by_detecting_column_headers(ws)

            # 标题格式
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            # 月份合并单元格和颜色
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
                max_length = 0
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 10

            # 保存原始数据
            forecast_df.to_excel(writer, index=False, sheet_name="原始-预测")
            order_df.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_df.to_excel(writer, index=False, sheet_name="原始-出货")

        output.seek(0)
        return main_df, output
