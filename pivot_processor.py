import os
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
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
            fill_order_data,
            fill_sales_data,
            highlight_by_detecting_column_headers,
            detect_forecast_header,
            merge_by_product_name_and_fill_specs
        )
        
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

        # ✅ 替换订单和出货品名字段（通常字段是“品名”）
        order_df, _ = apply_mapping_and_merge(order_df, mapping_new, {"品名": "品名"})
        order_df, _ = apply_extended_substitute_mapping(order_df, mapping_sub, {"品名": "品名"})
        
        sales_df, _ = apply_mapping_and_merge(sales_df, mapping_new, {"品名": "品名"})
        sales_df, _ = apply_extended_substitute_mapping(sales_df, mapping_sub, {"品名": "品名"})

        order_rename = {}
        sales_rename = {"晶圆": "晶圆品名"}
    
        def extract_unique_rows(df, rename_map):
            df = df.rename(columns=rename_map).copy()
            return df[["晶圆品名", "规格", "品名"]].dropna().drop_duplicates()
    
        order_part = extract_unique_rows(order_df, order_rename)
        sales_part = extract_unique_rows(sales_df, sales_rename)
        main_df = pd.concat([order_part, sales_part]).drop_duplicates().reset_index(drop=True)
    
        forecast_column_names = []
    
        # ✅ 用第一个预测文件确定 forecast_months
        if not forecast_files:
            raise ValueError("❌ 未上传任何预测文件")
        _, first_forecast_df = detect_forecast_header(forecast_files[0])
        all_months = extract_all_year_months(first_forecast_df, order_df, sales_df)
    
        for file in forecast_files:
            filename = os.path.basename(file.name)
            match = re.search(r'(\d{8})', filename)
            if not match:
                raise ValueError(f"❌ 文件名中缺少日期（8位数字）：{filename}")
            gen_date = datetime.strptime(match.group(1), "%Y%m%d")
            gen_ym = gen_date.strftime("%Y-%m")
            gen_month = gen_date.month
            gen_year = gen_date.year
    
            header_row, df = detect_forecast_header(file)
    
            # ✅ 设置品名和规格，并强制转字符串
            df["品名"] = df.iloc[:, 1].astype(str).str.strip()
            df = df.rename(columns={"产品型号": "规格"})
            df["规格"] = df["规格"].astype(str).str.strip()
    
            st.write(f"📂 已读取预测文件：**{filename}**（生成日期：{gen_ym}） header 行：第 {header_row + 1} 行")
            st.dataframe(df.head(10))
    
            df, _ = apply_mapping_and_merge(df, mapping_new, {"品名": "品名"})
            df, _ = apply_extended_substitute_mapping(df, mapping_sub, {"品名": "品名"})
    
            part_df = df[["规格", "品名"]].dropna().drop_duplicates()
            part_df["晶圆品名"] = ""
            main_df = pd.concat([main_df, part_df[["晶圆品名", "规格", "品名"]]]).drop_duplicates().reset_index(drop=True)
    
            month_only_pattern = re.compile(r"^(\d{1,2})月预测$")
            month_map = {}
            for col in df.columns:
                if not isinstance(col, str):
                    continue
                match = month_only_pattern.match(col)
                if match:
                    month_num = int(match.group(1))
                    year = gen_year + 1 if month_num < gen_month else gen_year
                    ym = f"{year}-{month_num:02d}"
                    month_map[ym] = col
    
            for ym, original_col in month_map.items():
                new_col_name = f"{gen_ym}的预测（{ym}）"
                if new_col_name not in forecast_column_names:
                    forecast_column_names.append(new_col_name)
                if new_col_name not in main_df.columns:
                    main_df[new_col_name] = 0.0  # ✅ 初始化为 float 避免 dtype 警告
    
                for _, row in df.iterrows():
                    product = str(row.get("品名", "")).strip()
                    spec = str(row.get("规格", "")).strip()
                    val = row.get(original_col, 0)
                    val = float(val) if pd.notna(val) else 0.0  # ✅ 显式转 float
                    mask = (
                        (main_df["品名"] == product)
                        & (main_df["规格"] == spec)
                    )
                    main_df.loc[mask, new_col_name] = val
    
        for ym in all_months:
            main_df[f"{ym}-订单"] = 0.0
            main_df[f"{ym}-出货"] = 0.0
    
        main_df = fill_order_data(main_df, order_df.rename(columns=order_rename), all_months)
        main_df = fill_sales_data(main_df, sales_df.rename(columns=sales_rename), all_months)

        # main_df = merge_by_product_name_and_fill_specs(main_df, mapping_df, order_df, sales_df)
    
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
    
            highlight_by_detecting_column_headers(ws)
    
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)
    
            col = 4
            dynamic_cols = [c for c in main_df.columns if c not in ["晶圆品名", "规格", "品名"]]
            group_keys = {}
            for col_name in dynamic_cols:
                if "预测" in col_name:
                    m = re.search(r"（(\d{4}-\d{2})）", col_name)
                    group = m.group(1) if m else col_name
                elif "-订单" in col_name or "-出货" in col_name:
                    group = col_name[:7]
                else:
                    group = "其它"
                group_keys.setdefault(group, []).append(col_name)
    
            fill_colors = [
                "FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFE599"
            ]
    
            col_idx = 4
            for i, (group, col_names) in enumerate(group_keys.items()):
                ws.merge_cells(start_row=1, start_column=col_idx,
                               end_row=1, end_column=col_idx + len(col_names) - 1)
                top_cell = ws.cell(row=1, column=col_idx)
                top_cell.value = group
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)
    
                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)],
                                   end_color=fill_colors[i % len(fill_colors)],
                                   fill_type="solid")
    
                for j, cname in enumerate(col_names):
                    ws.cell(row=2, column=col_idx + j).value = cname.split("的预测")[-1] if "预测" in cname else cname[-2:]
                    for r in [1, 2]:
                        ws.cell(row=r, column=col_idx + j).fill = fill
                col_idx += len(col_names)
    
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 8
    
            order_df.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_df.to_excel(writer, index=False, sheet_name="原始-出货")
    
        output.seek(0)
        return main_df, output
