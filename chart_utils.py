import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import re

def write_all_forecast_sheets(wb, df_main: pd.DataFrame):
    """
    一键生成所有预测分析相关 Sheet：预测展示、预测展开、预测展开（横向）。
    """
    def build_forecast_long_table(df: pd.DataFrame) -> pd.DataFrame:
        records = []
        forecast_cols = [col for col in df.columns if "预测" in col and "生成" in col]
        for _, row in df.iterrows():
            for col in forecast_cols:
                match = re.match(r"(\d{4}-\d{2})的预测（(\d{4}-\d{2})生成）", col)
                if not match:
                    continue
                forecast_month, gen_month = match.groups()
                records.append({
                    "品名": row["品名"],
                    "预测月份": forecast_month,
                    "生成月份": gen_month,
                    "预测值": row[col],
                    "订单量": row.get(f"{forecast_month}-订单", 0),
                    "出货量": row.get(f"{forecast_month}-出货", 0),
                })
        return pd.DataFrame(records)

    def write_forecast_expanded_sheet(wb, df_out: pd.DataFrame, sheet_name="预测展开"):
        ws = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(df_out, index=False, header=True):
            ws.append(r)
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)
        for i, col_cells in enumerate(ws.columns, 1):
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
            ws.column_dimensions[get_column_letter(i)].width = max_len + 4

    def write_forecast_expanded_wide_sheet(wb, df_out: pd.DataFrame, sheet_name="预测展开（横向）"):
        df = df_out.copy()
        df["预测月份"] = df["预测月份"].astype(str)
        df["生成月份"] = df["生成月份"].astype(str)
        group_fields = ["预测值", "订单量", "出货量"]
        wide = df.pivot_table(
            index=["品名", "预测月份"],
            columns="生成月份",
            values=group_fields,
            aggfunc="first"
        )
        wide.columns = [f"{col[1]}_{col[0]}" for col in wide.columns]
        wide.reset_index(inplace=True)
        ws = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(wide, index=False, header=True):
            ws.append(r)
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)
        for i, col_cells in enumerate(ws.columns, 1):
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
            ws.column_dimensions[get_column_letter(i)].width = max_len + 4

    # ✅ 构建长表
    df_out = build_forecast_long_table(df_main)

    # ✅ 写入三个 Sheet
    from forecast_utils import merge_monthly_group_headers, merge_and_color_monthly_group_headers
    ws = wb.create_sheet(title="预测展示")
    for r in dataframe_to_rows(df_main, index=False, header=True):
        ws.append(r)
    merge_monthly_group_headers(ws, df_main)
    merge_and_color_monthly_group_headers(ws, df_main)

    write_forecast_expanded_sheet(wb, df_out)
    write_forecast_expanded_wide_sheet(wb, df_out)
