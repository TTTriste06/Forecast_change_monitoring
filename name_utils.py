import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import re

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

def build_main_df(forecast_df, order_df, sales_df, mapping_new, mapping_sub):
    from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping

    # 🧩 提取所有品名（预测第2列、订单、出货）
    def extract_and_standardize(df, col_name):
        df = df[[col_name]].copy()
        df.columns = ["品名"]
        df["品名"] = df["品名"].astype(str).str.strip()
        return df

    forecast_names = extract_and_standardize(forecast_df.iloc[:, [1]], forecast_df.columns[1])
    order_names = extract_and_standardize(order_df, "品名")
    sales_names = extract_and_standardize(sales_df, "品名")

    all_names = pd.concat([forecast_names, order_names, sales_names], ignore_index=True)
    all_names = all_names.drop_duplicates(subset=["品名"]).copy()
    all_names["规格"] = ""
    all_names["晶圆品名"] = ""

    # ✅ 替换新旧料号（主替换 + 替代替换）
    all_names, _ = apply_mapping_and_merge(all_names, mapping_new, {"品名": "品名"})
    all_names, _ = apply_extended_substitute_mapping(all_names, mapping_sub, {"品名": "品名"})

    # ✅ 从映射表提取规格、晶圆品名
    mapping_clean = mapping_new[["新品名", "新规格", "新晶圆"]].copy()
    mapping_clean = mapping_clean.rename(columns={"新品名": "品名", "新规格": "规格", "新晶圆": "晶圆品名"})

    main_df = all_names.merge(mapping_clean, on="品名", how="left", suffixes=("", "_映射"))
    main_df["规格"] = main_df["规格"].where(main_df["规格"] != "", main_df["规格_映射"])
    main_df["晶圆品名"] = main_df["晶圆品名"].where(main_df["晶圆品名"] != "", main_df["晶圆品名_映射"])
    main_df.drop(columns=["规格_映射", "晶圆品名_映射"], inplace=True)

    # ✅ 从订单、出货、预测中依次补齐空规格和晶圆品名
    def try_fill(df_main, df_source, col_map):
        df_temp = df_source.rename(columns=col_map).copy()
        for col in ["品名", "规格"]:
            if col in df_temp.columns:
                df_temp[col] = df_temp[col].astype(str).str.strip()
        if "晶圆品名" in df_temp.columns:
            df_temp["晶圆品名"] = df_temp["晶圆品名"].astype(str).str.strip()
        elif "晶圆" in df_temp.columns:
            df_temp = df_temp.rename(columns={"晶圆": "晶圆品名"})
            df_temp["晶圆品名"] = df_temp["晶圆品名"].astype(str).str.strip()
        else:
            df_temp["晶圆品名"] = ""

        df_temp = df_temp[["品名", "规格", "晶圆品名"]].dropna(subset=["品名"]).drop_duplicates(subset=["品名"])
        df_main = df_main.merge(df_temp, on="品名", how="left", suffixes=("", "_补"))
        df_main["规格"] = df_main["规格"].where(df_main["规格"] != "", df_main["规格_补"])
        df_main["晶圆品名"] = df_main["晶圆品名"].where(df_main["晶圆品名"] != "", df_main["晶圆品名_补"])
        return df_main.drop(columns=["规格_补", "晶圆品名_补"])

    main_df = try_fill(main_df, order_df, {})
    main_df = try_fill(main_df, sales_df, {"晶圆": "晶圆品名"})
    main_df = try_fill(main_df, forecast_df.assign(晶圆品名=""), {"生产料号": "品名", "产品型号": "规格"})


    return main_df[["晶圆品名", "规格", "品名"]]
