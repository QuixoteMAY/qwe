import json

import lark_oapi as lark
import json
from requests_toolbelt import MultipartEncoder

#往表格写入数据
# SDK 使用说明: https://open.feishu.cn/document/uAjLw4CM/ukTMukTMukTM/server-side-sdk/python--sdk/preparations-before-development
# 以下示例代码默认根据文档示例值填充，如果存在代码问题，请在 API 调试台填上相关必要参数后再复制代码使用
def main():
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a8baea0bc979d00b") \
        .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    json_str = "{\"values\":[[[{\"text\":{\"text\":\"abc\"},\"type\":\"text\"}],[{\"text\":{\"text\":\"def\"},\"type\":\"text\"}]],[[{\"text\":{\"text\":\"123\"},\"type\":\"text\"}],[{\"text\":{\"text\":\"456\"},\"type\":\"text\"}]]]}"
    body = json.loads(json_str)
    request: lark.BaseRequest = lark.BaseRequest.builder() \
        .http_method(lark.HttpMethod.POST) \
        .uri("/open-apis/sheets/v3/spreadsheets/Md6Vsmwhhh4bDLtqNsQcW3Ounpe/sheets/rxgYPF/values/rxgYPF!A1:B2/insert?user_id_type=open_id") \
        .token_types({lark.AccessTokenType.TENANT}) \
        .body(body) \
        .build()

    # 发起请求
    response: lark.BaseResponse = client.request(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.request failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    # 处理业务结果
    lark.logger.info(str(response.raw.content, lark.UTF_8))


if __name__ == "__main__":
    main()