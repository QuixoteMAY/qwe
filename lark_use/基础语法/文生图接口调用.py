import base64
from pathlib import Path
from openai import OpenAI

# ==== 原始配置 ====
OPENAI_API_KEY = "AqMmbEXHndV5puKNnSXRIgXUYIBS6KLI_GPT_AK"
BASE_URL       = "https://search.bytedance.net/gpt/openapi/online/v2/crawl/openai"
MODEL_NAME     = "gpt-image-1"

# ==== 本地两张图片 ====

PROMPT = " 生成一张图"

# ==== 调用接口 ====
client = OpenAI(api_key=OPENAI_API_KEY, base_url=BASE_URL)


resp = client.images.generate(
    model=MODEL_NAME,                     # 同时上传3张图
    prompt=PROMPT,
    extra_headers={"api-key": OPENAI_API_KEY},
)

# ==== 保存唯一返回图 ====
out_path = Path("output_single.png")
out_path.write_bytes(base64.b64decode(resp.data[0].b64_json))
print(f"✓ 已保存：{out_path.resolve()}")
print(resp.data[0].b64_json)



