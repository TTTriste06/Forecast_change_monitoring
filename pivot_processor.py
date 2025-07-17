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
        from forecast_utils import (
            extract_forecast_generation_date, 
            extract_forecast_data, parse_forecast_months, 
            append_multi_forecast_columns, 
            merge_forecast_columns,
            parse_forecast_columns,
            load_forecast_files
        )
    
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_file)
    
        # ✅ 统一品名+晶圆+规格列表
        main_df = build_main_df(pd.DataFrame(), order_file, sales_file, mapping_new, mapping_sub)
    
        # ✅ 多文件读取 & 解析预测
        forecast_cols_to_merge = []
        for uploaded_file in forecast_files:
            file_name = uploaded_file.name
    
            # 🔍 1. 从文件名提取预测生成年月
            match = re.search(r"_(\d{8})", file_name)
            if not match:
                st.warning(f"⚠ 无法从文件名提取日期：{file_name}")
                continue
            gen_date = datetime.strptime(match.group(1), "%Y%m%d")
            gen_year, gen_month = gen_date.year, gen_date.month
            gen_ym_str = f"{gen_year}-{str(gen_month).zfill(2)}"
            label = f"{gen_ym_str}的预测"
    
            # 📄 2. 提取最长 sheet 和 header 行
            xls = pd.ExcelFile(uploaded_file)
            sheet_lens = {s: pd.read_excel(xls, sheet_name=s, header=None).shape[0] for s in xls.sheet_names}
            longest_sheet = max(sheet_lens, key=sheet_lens.get)
            df_raw = pd.read_excel(xls, sheet_name=longest_sheet, header=None)
    
            header_row = None
            for idx, row in df_raw.iterrows():
                if row.astype(str).str.contains("产品型号").any():
                    header_row = idx
                    break
            if header_row is None:
                st.warning(f"⚠ 文件 {file_name} 中未找到包含“产品型号”的表头行，跳过")
                continue
    
            df_forecast = pd.read_excel(uploaded_file, sheet_name=longest_sheet, header=header_row)
            df_forecast.columns.values[1] = "品名"
            df_forecast["品名"] = df_forecast["品名"].astype(str).str.strip()
    
            # 📅 3. 提取月份列（如 “5月预测”）并判断跨年，生成 yyyy-mm
            forecast_months = {}
            last_month = 0
            forecast_year = gen_year
            for col in df_forecast.columns:
                m = re.match(r"^(\d{1,2})月预测$", str(col).strip())
                if m:
                    m_num = int(m.group(1))
                    if last_month and m_num < last_month:
                        forecast_year += 1  # 跨年
                    ym = f"{forecast_year}-{str(m_num).zfill(2)}"
                    forecast_months[ym] = col
                    last_month = m_num
    
            # 🧩 4. 添加预测列到 main_df
            for ym, colname in forecast_months.items():
                new_col = f"{ym}的预测（{gen_ym_str}）"
                if new_col not in main_df.columns:
                    main_df[new_col] = 0
                for i, row in main_df.iterrows():
                    val = df_forecast.loc[df_forecast["品名"] == row["品名"], colname]
                    if not val.empty:
                        main_df.at[i, new_col] = val.values[0]
    
            st.success(f"✅ 成功处理文件：{file_name}")
            st.dataframe(df_forecast.head())
    
        # ✅ 替换订单、出货文件字段
        FIELD_MAPPINGS = {
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }
        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])
        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])
    
        # ✅ 提取所有月份
        all_months = extract_all_year_months(pd.DataFrame(), order_file, sales_file)
        for ym in all_months:
            if f"{ym}-订单" not in main_df.columns:
                main_df[f"{ym}-订单"] = 0
            if f"{ym}-出货" not in main_df.columns:
                main_df[f"{ym}-出货"] = 0
    
        main_df = fill_order_data(main_df, order_file, all_months)
        main_df = fill_sales_data(main_df, sales_file, all_months)
    
        # ✅ 写入 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
            wb = writer.book
    
            highlight_by_detecting_column_headers(ws)
    
            from openpyxl.styles import Alignment, PatternFill, Font
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
            forecast_cols = [c for c in main_df.columns if "预测（" in c]
            order_cols = [f"{m}-订单" for m in all_months]
            sales_cols = [f"{m}-出货" for m in all_months]
    
            grouped = sorted(set([c[:7] for c in forecast_cols]))
            for i, ym in enumerate(grouped):
                subcols = [c for c in forecast_cols if c.startswith(ym)] + \
                          ([f"{ym}-订单"] if f"{ym}-订单" in main_df.columns else []) + \
                          ([f"{ym}-出货"] if f"{ym}-出货" in main_df.columns else [])
    
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + len(subcols) - 1)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)
    
                for j, sub in enumerate(subcols):
                    ws.cell(row=2, column=col + j).value = sub.replace(f"{ym}-", "").replace("的预测", "").replace("（", "\n（")
                    fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                    ws.cell(row=1, column=col + j).fill = fill
                    ws.cell(row=2, column=col + j).fill = fill
    
                col += len(subcols)
    
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = max((len(str(cell.value)) if cell.value else 0) for cell in column_cells)
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 8
    
        output.seek(0)
        return main_df, output
