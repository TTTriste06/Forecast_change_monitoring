import os
import re
import pandas as pd
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
            highlight_by_detecting_column_headers
        )

        # ========== 工具函数：标准化字段 ==========
        def extract_unique_rows(df, rename_dict):
            df = df.rename(columns=rename_dict)
            required_cols = ["晶圆品名", "规格", "品名"]
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                st.warning(f"⚠️ 缺少字段 {missing}，已跳过该部分唯一值提取")
                return pd.DataFrame(columns=required_cols)
            return df[required_cols].dropna().drop_duplicates()

        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

        # ========== 1. 初始化主表 ==========
        forecast_column_names = []
        main_df = pd.DataFrame()

        for file in forecast_files:
            # 1.1 文件名提取生成年月
            filename = os.path.basename(file.name)
            match = re.search(r'(\d{8})', filename)
            if not match:
                raise ValueError(f"❌ 无法从文件名中提取日期（应包含 8 位数字）：{filename}")
            gen_date = datetime.strptime(match.group(1), "%Y%m%d")
            gen_ym = gen_date.strftime("%Y-%m")
            gen_year = gen_date.year
            gen_month = gen_date.month

            # 1.2 读取最后一个 sheet
            xls = pd.ExcelFile(file)
            df = xls.parse(xls.sheet_names[-1])
            df = df.rename(columns={"生产料号": "品名"})  # 标准品名字段

            # 1.3 应用映射
            df, _ = apply_mapping_and_merge(df, mapping_new, {"品名": "品名"})
            df, _ = apply_extended_substitute_mapping(df, mapping_sub, {"品名": "品名"})

            # 1.4 提取标准字段用于主表初始化
            part_df = extract_unique_rows(df, {"品名": "品名", "规格": "产品型号", "晶圆品名": "晶圆"})
            main_df = pd.concat([main_df, part_df]).drop_duplicates().reset_index(drop=True)

            # 1.5 识别月份预测列
            month_only_pattern = re.compile(r"^(\d{1,2})月预测")
            month_map = {}
            for col in df.columns:
                if not isinstance(col, str):
                    continue
                match = month_only_pattern.match(col)
                if match:
                    m_num = int(match.group(1))
                    year = gen_year + 1 if m_num < gen_month else gen_year
                    ym = f"{year}-{m_num:02d}"
                    month_map[ym] = col

            for ym, original_col in month_map.items():
                new_col = f"{gen_ym}的预测（{ym}）"
                forecast_column_names.append(new_col)
                if new_col not in main_df.columns:
                    main_df[new_col] = 0

                for idx, row in df.iterrows():
                    pname = str(row.get("品名", "")).strip()
                    wafer = str(row.get("晶圆品名", "")).strip()
                    spec = str(row.get("产品型号", "")).strip()
                    val = row.get(original_col, 0)
                    mask = (
                        (main_df["品名"] == pname)
                        & (main_df["晶圆品名"] == wafer)
                        & (main_df["规格"] == spec)
                    )
                    main_df.loc[mask, new_col] = val

        # ========== 2. 加入订单、出货信息 ==========
        order_part = extract_unique_rows(order_df, {
            "品名": "品名", "规格": "规格", "晶圆品名": "晶圆"
        })
        sales_part = extract_unique_rows(sales_df, {
            "品名": "品名", "规格": "规格", "晶圆品名": "晶圆"
        })
        main_df = pd.concat([main_df, order_part, sales_part]).drop_duplicates().reset_index(drop=True)

        all_months = extract_all_year_months(None, order_df, sales_df)
        for ym in all_months:
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        main_df = fill_order_data(main_df, order_df, all_months)
        main_df = fill_sales_data(main_df, sales_df, all_months)

        # ========== 3. 输出 Excel ==========
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
            highlight_by_detecting_column_headers(ws)

            # 表头格式
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            # 动态列合并与配色
