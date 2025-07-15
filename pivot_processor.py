import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from github_utils import load_file_with_github_fallback
from info_extract import extract_year_month_from_filename
from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping, split_mapping_data

class PivotProcessor:
    def process(self, forecast_files, order_df, sales_df, mapping_df):
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

        # 提取所有晶圆/规格/品名用于初始化
        dfs_with_fields = []
        for file in forecast_files:
            df = load_file_with_github_fallback("forecast", file)
            df = df.rename(columns=lambda x: x.strip())
            if {"晶圆", "规格", "生产料号"}.issubset(df.columns):
                df["晶圆"] = df["晶圆"].astype(str).str.strip()
                df["规格"] = df["规格"].astype(str).str.strip()
                df["品名"] = df["生产料号"].astype(str).str.strip()
                dfs_with_fields.append(df[["晶圆", "规格", "品名"]])

        for df in [order_df, sales_df]:
            if {"晶圆", "规格", "品名"}.issubset(df.columns):
                df = df.rename(columns=lambda x: x.strip())
                dfs_with_fields.append(df[["晶圆", "规格", "品名"]].copy())

        main_df = pd.concat(dfs_with_fields).drop_duplicates().reset_index(drop=True)
        main_df.columns = ["晶圆品名", "规格", "品名"]

        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }

        order_df, _ = apply_mapping_and_merge(order_df, mapping_new, FIELD_MAPPINGS["order"])
        order_df, _ = apply_extended_substitute_mapping(order_df, mapping_sub, FIELD_MAPPINGS["order"])
        sales_df, _ = apply_mapping_and_merge(sales_df, mapping_new, FIELD_MAPPINGS["sales"])
        sales_df, _ = apply_extended_substitute_mapping(sales_df, mapping_sub, FIELD_MAPPINGS["sales"])

        # 多预测文件处理
        all_months = set()
        for file in forecast_files:
            df = load_file_with_github_fallback("forecast", file, sheet_name="预测")
            df = df.rename(columns=lambda x: x.strip())
            df["生产料号"] = df["生产料号"].astype(str).str.strip()
            df["品名"] = df["生产料号"]

            year_month = extract_year_month_from_filename(file.name)
            if year_month is None:
                continue

            col_name = f"{year_month}的预测"
            forecast_series = df.groupby("品名")["预测"].sum(min_count=1)
            main_df[col_name] = main_df["品名"].map(forecast_series).fillna(0)
            all_months.add(year_month)

        all_months = sorted(all_months)

        from info_extract import fill_order_data, fill_sales_data
        main_df = fill_order_data(main_df, order_df, all_months)
        main_df = fill_sales_data(main_df, sales_df, all_months)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)

            ws = writer.sheets["预测分析"]
            wb = writer.book

            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.font = Font(bold=True)

        output.seek(0)
        return main_df, output
