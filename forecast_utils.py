import pandas as pd
import streamlit as st

def load_forecast_files(files: dict) -> dict[str, pd.DataFrame]:
    """
    对上传的多个预测 Excel 文件执行以下操作：
    - 找到每个文件中最长的 sheet
    - 自动识别 header 行（含“产品型号”的那一行）
    - 将第二列统一命名为“品名”
    - 用 st.write 打印每个文件读取结果
    返回值：dict[file_name -> cleaned DataFrame]
    """
    result = {}

    for file_name, file in files.items():
        try:
            xls = pd.ExcelFile(file)
            longest_sheet = max(xls.sheet_names, key=lambda name: pd.read_excel(xls, sheet_name=name).shape[0])
            df_raw = pd.read_excel(xls, sheet_name=longest_sheet, header=None)

            # 自动识别 header 行：包含“产品型号”的行
            header_row_idx = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("产品型号").any(), axis=1)].index
            if header_row_idx.empty:
                st.warning(f"⚠ 文件 {file_name} 中未找到包含“产品型号”的表头行，跳过")
                continue

            header_row = header_row_idx[0]
            df = pd.read_excel(xls, sheet_name=longest_sheet, header=header_row)

            # 统一第二列为“品名”
            if df.shape[1] >= 2:
                df.columns.values[1] = "品名"

            st.write(f"📄 读取成功：{file_name}（使用 sheet：{longest_sheet}，header 行：第 {header_row+1} 行）")
            st.dataframe(df)

            result[file_name] = df

        except Exception as e:
            st.error(f"❌ 无法读取文件 {file_name}: {e}")

    return result
