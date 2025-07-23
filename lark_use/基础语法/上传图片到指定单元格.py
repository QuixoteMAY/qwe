import json
import lark_oapi as lark

#https://bytedance.larkoffice.com/sheets/VzjhsxJaehgWputBmRNcMqu1nRc?sheet=4a8bbe




#!/usr/bin/env python3
"""
============================================================
  把本地图片 ➜ 转成 [137, 80, 78, 71, ...] 这种整数列表
  —— 只改最上面的 FILE_PATH，然后直接运行就行
============================================================
"""

# ====== 配置区：只需要改这一行，把图片路径写进来 ======
FILE_PATH = "C:/Users/Admin/Desktop/image_create/image - 2025-06-26T211913.857.png"          # ← 换成你的图片 / 任意二进制文件
# ==============================================================


def image_to_byte_list(path: str) -> list[int]:
    """
    读取给定文件，按字节拆成 0-255 的整数列表
    参数
        path : str
            图片或其他二进制文件的路径
    返回
        list[int]
            文件内容对应的整数数组
    """
    # 1) 以二进制方式（rb）打开文件
    with open(path, "rb") as file:
        raw_bytes = file.read()            # 2) 读出全部字节

    byte_list = list(raw_bytes)            # 3) bytes ➜ list[int]
    return byte_list                       # 4) 返回给调用者


# ----------------------- 立即执行 -----------------------------
# 直接调用函数并把结果打印出来
data = image_to_byte_list(FILE_PATH)


# 总字节数，也就是文件大小
print("文件总长度：", len(data), "字节")

# 如果需要把完整列表写到文本文件，取消下面两行的注释
# with open("byte_list.txt", "w") as out:
#     out.write(str(data))






# ── 1. 创建 client ──────────────────────────────────────
client = (
    lark.Client.builder()
    .app_id("cli_a8baea0bc979d00b")
    .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX")
    .log_level(lark.LogLevel.DEBUG)
    .build()
)

# ── 2. 构造请求对象 ────────────────────────────────────
request: lark.BaseRequest = (
    lark.BaseRequest.builder()
    .http_method(lark.HttpMethod.POST)
    .uri(
        "/open-apis/sheets/v2/spreadsheets/VzjhsxJaehgWputBmRNcMqu1nRc/values_image"
    )
    .token_types([lark.AccessTokenType.TENANT])
    .body({"range":"4a8bbe!E2:E2","image":data,"name":"test.png"})
    .build()
)

# ── 3. 发起请求 ───────────────────────────────────────
response: lark.BaseResponse = client.request(request)

# ── 4. 处理返回 ───────────────────────────────────────
if not response.success():
    lark.logger.error("Request failed: %s", response.code)
    exit(1)



# 如果你还想查看完整 JSON，可保留下面这一行
lark.logger.info(str(response.raw.content, lark.UTF_8))










