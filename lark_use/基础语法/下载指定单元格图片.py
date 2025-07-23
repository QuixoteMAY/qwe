


import json

import lark_oapi as lark
from lark_oapi.api.drive.v1 import *


# SDK 使用说明: https://open.feishu.cn/document/uAjLw4CM/ukTMukTMukTM/server-side-sdk/python--sdk/preparations-before-development
# 以下示例代码默认根据文档示例值填充，如果存在代码问题，请在 API 调试台填上相关必要参数后再复制代码使用
# 复制该 Demo 后, 需要将 "YOUR_APP_ID", "YOUR_APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
def main():
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a8baea0bc979d00b") \
        .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: DownloadMediaRequest = DownloadMediaRequest.builder() \
        .file_token("Kg3ybK978o1V2ExaKkGcO68onxb") \
        .extra("无") \
        .build()

    # 发起请求
    response: DownloadMediaResponse = client.drive.v1.media.download(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.drive.v1.media.download failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    # 处理业务结果
    f = open(f"C:/Users/Admin/Desktop/image_create/{response.file_name}", "wb")
    f.write(response.file.read())
    f.close()


if __name__ == "__main__":
    main()






















