import lark_oapi as lark

# 创建client
client = lark.Client.builder() \
    .app_id("cli_a8baea0bc979d00b") \
    .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX") \
    .log_level(lark.LogLevel.DEBUG) \
    .build()

# 构造请求对象
request:lark.BaseRequest = lark.BaseRequest.builder() \
    .http_method(lark.HttpMethod.POST) \
    .uri("/open-apis/sheets/v2/spreadsheets/VzjhsxJaehgWputBmRNcMqu1nRc/merge_cells") \
    .token_types([lark.AccessTokenType.TENANT]) \
    .body({
    "range": "qPO23n!A1:A2", 
    "mergeType": "MERGE_ALL"
}) \
    .build()

# 发起请求
response: lark.BaseResponse = client.request(request)

# 处理失败返回
if not response.success():
    lark.logger.error("...")

# 处理业务结果
lark.logger.info(str(response.raw.content, lark.UTF_8))

























