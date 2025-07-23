import base64
from pathlib import Path
from openai import OpenAI

# ==== 原始配置 ====
OPENAI_API_KEY = "AqMmbEXHndV5puKNnSXRIgXUYIBS6KLI_GPT_AK"
BASE_URL       = "https://search.bytedance.net/gpt/openapi/online/v2/crawl/openai"
MODEL_NAME     = "gpt-image-1"

# ==== 本地两张图片 ====
IMG1 = Path(r"C:/Users/Admin/Desktop/image_create/1.png")   # ← 修改为实际路径
IMG2 = Path(r"C:/Users/Admin/Desktop/image_create/2.png")   # ← 修改为实际路径
IMG3 = Path(r"C:/Users/Admin/Desktop/image_create/3.png")   # ← 修改为实际路径
PROMPT = "融合为一张图"

# ==== 调用接口 ====
client = OpenAI(api_key=OPENAI_API_KEY, base_url=BASE_URL)

with IMG1.open("rb") as f1, IMG2.open("rb") as f2, IMG3.open("rb") as f3:
    resp = client.images.edit(
        model=MODEL_NAME,
        image=[f1, f2, f3],                      # 同时上传3张图
        prompt=PROMPT,
        extra_headers={"api-key": OPENAI_API_KEY},
    )

# ==== 保存唯一返回图 ====
out_path = Path("output_single.png")
out_path.write_bytes(base64.b64decode(resp.data[0].b64_json))
print(f"✓ 已保存：{out_path.resolve()}")
