"""
读取飞书电子表格指定范围 (Q7PlXT!A1:B2) 的示例脚本
等价于：
curl --location --request GET \
  'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/XUMasQlMYhOnMbt5htXc96h0nOg/values/Q7PlXT!A1:B2?valueRenderOption=ToString&dateTimeRenderOption=FormattedString' \
  --header 'Authorization: Bearer <token>'
"""

import os
import requests

# === 必填参数 ===
SPREADSHEET_TOKEN = "F2n7sGrjjhRdARt7UMrcFXoYnWh"   # 电子表格 token
SHEET_RANGE       = "e580f2!C8:C9"                  # sheetId!范围
TOKEN             = os.getenv("FEISHU_TOKEN") or "t-g1046il34H3ETQC7LLVAOLG46V2XHG662GJCL7G6"

# === 组装请求 ===
url = (f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/"
       f"{SPREADSHEET_TOKEN}/values/{SHEET_RANGE}")

params = {
    "valueRenderOption": "ToString",
    "dateTimeRenderOption": "FormattedString"
}

headers = {
    "Authorization": f"Bearer {TOKEN}",
    # 根据官方文档要求必须带上 Content-Type
    "Content-Type": "application/json; charset=utf-8",
}

# === 发起 GET 请求 ===
resp = requests.get(url, headers=headers, params=params, timeout=10)
resp.raise_for_status()   # 非 2xx 会抛异常

# === 解析结果 ===
data = resp.json()
print(data)

# 若只关心单元格内容，可进一步提取：
# values = data["data"]["valueRange"]["values"]
# print(values)
