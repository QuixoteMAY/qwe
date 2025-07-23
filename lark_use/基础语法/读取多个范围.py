import os
import requests

# ========= 必填参数 =========
SPREADSHEET_TOKEN = "F2n7sGrjjhRdARt7UMrcFXoYnWh"           # 电子表格 token
RANGES = ["e580f2!C8:C8","e580f2!C9:C9"]                   # 要查询的多个范围
TOKEN = os.getenv("FEISHU_TOKEN") or "t-g1046j9TNKMV2YVL4LFYT3OBI3GW6L4ENFYUV2L3"

# ========= 组装 URL 与查询参数 =========
url = (
    f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/"
    f"{SPREADSHEET_TOKEN}/values_batch_get"
)

params = {
    "ranges": ",".join(RANGES),                # 逗号分隔多个范围
    "dateTimeRenderOption": "FormattedString"  # 日期时间按格式化字符串返回
}

headers = {
    "Authorization": f"Bearer {TOKEN}"
}

# ========= 发起 GET 请求 =========
response = requests.get(url, headers=headers, params=params, timeout=10)
response.raise_for_status()          # 若接口未返回 2xx，会抛出异常

data = response.json()
print(data)                          # 完整响应

# 若只关心每个范围的具体数据，可进一步提取：
# for vr in data["data"]["valueRanges"]:
#     print(f"Range: {vr['range']}")
#     print("Values:", vr["values"])
