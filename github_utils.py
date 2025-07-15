from io import BytesIO
import base64
import requests
import streamlit as st
import pandas as pd
from urllib.parse import quote

# GitHub 配置
GITHUB_TOKEN_KEY = "GITHUB_TOKEN"  # secrets.toml 中的密钥名
REPO_NAME = "TTTriste06/Forecast_change_monitoring"
BRANCH = "main"

FILENAME_KEYS = {
    "forecast": "预测.xlsx",
    "order": "未交订单.xlsx",
    "sales": "出货明细.xlsx",
    "template": "预测分析.xlsx"
}

def upload_to_github(file_obj, filename):
    """
    将 file_obj 文件上传至 GitHub 指定仓库
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)  # 支持中文

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    file_obj.seek(0)
    content = base64.b64encode(file_obj.read()).decode("utf-8")
    file_obj.seek(0)

    # 检查是否已存在
    sha = None
    get_resp = requests.get(url, headers=headers)
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")

    payload = {
        "message": f"upload {filename}",
        "content": content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(url, headers=headers, json=payload)
    if put_resp.status_code not in [200, 201]:
        raise Exception(f"❌ 上传失败：{put_resp.status_code} - {put_resp.text}")


def download_from_github(filename):
    """
    从 GitHub 下载文件内容（二进制返回）
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}?ref={BRANCH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        json_resp = response.json()
        return base64.b64decode(json_resp["content"])
    else:
        raise FileNotFoundError(f"❌ GitHub 上找不到文件：{filename} (HTTP {response.status_code})")


def load_file_with_github_fallback(file_key, uploaded_file, sheet_name=0, header=0):
    fallback_urls = {
        "order": "https://raw.githubusercontent.com/TTTriste06/forecast-analysis/main/未交订单.xlsx",
        "sales": "https://raw.githubusercontent.com/TTTriste06/forecast-analysis/main/出货明细.xlsx",
        "mapping": "https://raw.githubusercontent.com/TTTriste06/operation_planning-/main/新旧料号.xlsx"
    }

    if uploaded_file is not None:
        # ✅ 自动上传新文件到 GitHub
        filename = FILENAME_KEYS.get(file_key)
        if filename:
            upload_to_github(uploaded_file, filename)

        # ✅ 返回本地上传的文件内容
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header, engine="openpyxl")

    # fallback 下载
    if file_key not in fallback_urls:
        raise ValueError(f"⚠️ 未识别的辅助文件类型：{file_key}")

    url = fallback_urls[file_key]
    response = requests.get(url)
    if not response.ok:
        raise ValueError(f"❌ 无法从 GitHub 获取文件：{url}")

    content = response.content
    try:
        return pd.read_excel(BytesIO(content), sheet_name=sheet_name, header=header, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"❌ 无法读取 Excel 文件（可能不是 .xlsx 格式）：{e}")
