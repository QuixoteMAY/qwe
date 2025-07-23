import lark_oapi as lark

# 创建client
client = lark.Client.builder() \
    .app_id("cli_a8baea0bc979d00b") \
    .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX") \
    .log_level(lark.LogLevel.DEBUG) \
    .build()

# 构造请求对象
request:lark.BaseRequest = lark.BaseRequest.builder() \
    .http_method(lark.HttpMethod.PUT) \
    .uri("/open-apis/sheets/v2/spreadsheets/VzjhsxJaehgWputBmRNcMqu1nRc/values") \
    .token_types([lark.AccessTokenType.TENANT]) \
    .body({"valueRange":{"range":"4a8bbe!C2:C2","values":[["ceshiceshi"]]}}) \
    .build()

# 发起请求
response: lark.BaseResponse = client.request(request)

# 处理失败返回
if not response.success():
    lark.logger.error("...")

# 处理业务结果
lark.logger.info(str(response.raw.content, lark.UTF_8))

























