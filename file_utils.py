import pandas as pd
import re


def extract_forecast_month_from_filename(filename: str, current_year: int) -> str:
    match = re.search(r"(\d{1,2})月", filename)
    if match:
        month = int(match.group(1))
        year = current_year
        if month < datetime.now().month:
            year += 1  # 跨年
        return f"{year}-{month:02d}"
    return None

def read_forecast_file(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    sheet_lens = {sheet: pd.read_excel(xls, sheet).shape[0] for sheet in xls.sheet_names}
    longest_sheet = max(sheet_lens, key=sheet_lens.get)
    df = pd.read_excel(xls, longest_sheet, header=None)

    for i in range(3):
        header_row = df.iloc[i].astype(str)
        if any("产品型号" in str(cell) for cell in header_row):
            df.columns = header_row
            df = df.iloc[i+1:]
            break
    else:
        for i in range(3):
            if any("预测" in str(cell) for cell in df.iloc[i]):
                df.columns = df.iloc[i]
                df = df.iloc[i+1:]
                break
    df = df.reset_index(drop=True)
    df = df.rename(columns={df.columns[1]: "品名"})
    return df
