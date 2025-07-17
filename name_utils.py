def extract_unique_rows_from_all_sources(forecast_files, order_df, sales_df, mapping_df):
    from mapping_utils import (
        apply_mapping_and_merge,
        apply_extended_substitute_mapping,
        split_mapping_data
    )
    
    mapping_semi, mapping_main, mapping_sub = split_mapping_data(mapping_df)

    def extract_and_map(df, col_name):
        df = df.rename(columns={col_name: "品名"})
        df["品名"] = df["品名"].astype(str).str.strip()
        df = apply_mapping_and_merge(df, mapping_main)
        df = apply_extended_substitute_mapping(df, mapping_sub)
        return df[["品名"]]

    # 提取预测中第2列作为品名
    forecast_parts = []
    for _, file in forecast_files.items():
        df = pd.read_excel(file, sheet_name=0)
        if df.shape[1] >= 2:
            df_forecast = df.iloc[:, [1]].copy()
            df_forecast.columns = ["品名"]
            df_forecast = apply_mapping_and_merge(df_forecast, mapping_main)
            df_forecast = apply_extended_substitute_mapping(df_forecast, mapping_sub)
            forecast_parts.append(df_forecast[["品名"]])
    
    forecast_df = pd.concat(forecast_parts, ignore_index=True) if forecast_parts else pd.DataFrame(columns=["品名"])
    order_df = extract_and_map(order_df, "品名")
    sales_df = extract_and_map(sales_df, "品名")

    # 合并并去重
    all_names = pd.concat([forecast_df, order_df, sales_df], ignore_index=True)
    all_names.drop_duplicates(inplace=True)

    # 加入晶圆和规格
    mapping_main["品名"] = mapping_main["新料号"].astype(str).str.strip()
    mapping_main["晶圆"] = mapping_main["晶圆"].astype(str).str.strip()
    mapping_main["规格"] = mapping_main["规格"].astype(str).str.strip()

    result = all_names.merge(mapping_main[["品名", "晶圆", "规格"]], on="品名", how="left")
    result = result[["品名", "晶圆", "规格"]].fillna("")

    return result

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

    # 新旧料号替换
    all_names, _ = apply_mapping_and_merge(all_names, mapping_new, {"品名": "品名"})
    all_names, _ = apply_extended_substitute_mapping(all_names, mapping_sub, {"品名": "品名"})

    # 从映射表获取晶圆、规格
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

    # ✅ 如果还有空规格或晶圆品名，尝试从原始文件中补全
    def try_fill_from(df, col_map, source_name):
        df_temp = df.copy()
        df_temp = df_temp.rename(columns=col_map)
        df_temp = df_temp[["品名", "规格", "晶圆品名"]].dropna(subset=["品名"])
        df_temp["品名"] = df_temp["品名"].astype(str).str.strip()
        df_temp["规格"] = df_temp["规格"].astype(str).str.strip()
        df_temp["晶圆品名"] = df_temp["晶圆品名"].astype(str).str.strip()
        return df_temp.drop_duplicates()

    order_info = try_fill_from(order_file, {}, "order")
    sales_info = try_fill_from(sales_file, {"晶圆": "晶圆品名"}, "sales")
    forecast_info = try_fill_from(forecast_file, {"生产料号": "品名", "产品型号": "规格"}, "forecast")
    forecast_info["晶圆品名"] = ""  # 预测中没有晶圆

    combined = pd.concat([order_info, sales_info, forecast_info], ignore_index=True)

    # 按照品名左连接补齐规格和晶圆品名
    merged = merged.merge(
        combined,
        on="品名",
        how="left",
        suffixes=("", "_补")
    )

    # 如果原规格/晶圆为空，用补字段补上
    merged["规格"] = merged["规格"].mask(merged["规格"] == "", merged["规格_补"])
    merged["晶圆品名"] = merged["晶圆品名"].mask(merged["晶圆品名"] == "", merged["晶圆品名_补"])

    return merged[["晶圆品名", "规格", "品名"]]

