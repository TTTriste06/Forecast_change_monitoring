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
