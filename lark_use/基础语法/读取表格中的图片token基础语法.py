import json
import lark_oapi as lark

#https://bytedance.larkoffice.com/sheets/VzjhsxJaehgWputBmRNcMqu1nRc?sheet=4a8bbe

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
    .http_method(lark.HttpMethod.GET)
    .uri(
        "/open-apis/sheets/v2/spreadsheets/VzjhsxJaehgWputBmRNcMqu1nRc/values_batch_get"
    )
    .token_types([lark.AccessTokenType.TENANT])
    .queries([("ranges", "4a8bbe!A2:A500")])
    .build()
)

# ── 3. 发起请求 ───────────────────────────────────────
response: lark.BaseResponse = client.request(request)

# ── 4. 处理返回 ───────────────────────────────────────
if not response.success():
    lark.logger.error("Request failed: %s", response.code)
    exit(1)

raw_json = json.loads(response.raw.content.decode("utf-8"))

try:
    file_token = (
        raw_json["data"]["valueRanges"][0]["values"][0][0]["fileToken"]
    )
except (KeyError, IndexError, TypeError) as e:
    lark.logger.error("fileToken 未找到，错误信息: %s", e)
    exit(1)

# 如果你还想查看完整 JSON，可保留下面这一行
lark.logger.info(json.dumps(raw_json, ensure_ascii=False, indent=2))

# ── 5. 最后只输出 fileToken ───────────────────────────
print(file_token)













