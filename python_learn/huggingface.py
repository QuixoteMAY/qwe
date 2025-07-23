import re
import json
import io
import requests
from datasets import load_dataset
from tqdm import tqdm
from xlsxwriter.utility import xl_rowcol_to_cell

import lark_oapi as lark
from PIL.Image import Image as PILImage  # 用于识别 PIL 图片对象


# ======================= 配置区 =======================
dataset_path = "derek-thomas/ScienceQA"
total_items = 2
spreadsheet_token = "Rw9ys4Usmhlx1htgsXUcjtTpnde"
sheet_id = "49fa89"

# 初始化飞书 client
client = lark.Client.builder() \
    .app_id("cli_a8baea0bc979d00b") \
    .app_secret("RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX") \
    .log_level(lark.LogLevel.INFO) \
    .build()

# 匹配图片 URL 的正则表达式
IMAGE_URL_PATTERN = re.compile(r"https?://.*\.(png|jpg|jpeg|gif|webp|bmp)", re.IGNORECASE)


# ------------------------------------------------------------------
# 新增：把 CMYK 或其他模式转换成 RGB，确保能保存为 PNG
# ------------------------------------------------------------------
def safe_image_to_rgb(img: PILImage) -> PILImage:
    """
    如果图片是 CMYK，则转成 RGB；否则原样返回。
    """
    if img.mode == "CMYK":
        return img.convert("RGB")
    return img


# ------------------------------------------------------------------
# 其余函数与原先相同，不再赘述
# ------------------------------------------------------------------
def flatten_dict(d: dict, parent_key: str = "", sep: str = ".") -> dict:
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def download_image_as_byte_list(url: str) -> list[int]:
    headers = {"User-Agent": "MyApp/1.0"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    if not content_type.startswith("image/"):
        raise ValueError(f"URL {url} 未返回图片，Content-Type: {content_type}")

    return list(response.content)


def write_text_to_feishu(cell: str, text: str):
    print(f"[写入文字] → {cell}: {text}")
    request = lark.BaseRequest.builder() \
        .http_method(lark.HttpMethod.PUT) \
        .uri(f"/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values") \
        .token_types([lark.AccessTokenType.TENANT]) \
        .body({
            "valueRange": {
                "range": f"{sheet_id}!{cell}:{cell}",
                "values": [[text]]
            }
        }) \
        .build()

    response = client.request(request)
    if not response.success():
        lark.logger.error(f"文字写入失败: {response.code}")
    else:
        lark.logger.info(f"文字写入成功: {cell}")


def write_image_to_feishu(cell: str, image_byte_list: list[int], image_name="image.png"):
    print(f"[写入图片] → {cell}: 共 {len(image_byte_list)} 字节")
    request = lark.BaseRequest.builder() \
        .http_method(lark.HttpMethod.POST) \
        .uri(f"/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_image") \
        .token_types([lark.AccessTokenType.TENANT]) \
        .body({
            "range": f"{sheet_id}!{cell}:{cell}",
            "image": image_byte_list,
            "name": image_name
        }) \
        .build()

    response = client.request(request)
    if not response.success():
        lark.logger.error(f"图片写入失败: {response.code}")
    else:
        lark.logger.info(f"图片写入成功: {cell}")


def load_hf_dataset():
    return load_dataset(dataset_path, split="train", trust_remote_code=True, streaming=True)


def main():
    ds = load_hf_dataset().take(total_items)
    headers_written = False
    headers = []

    for row_idx, item in enumerate(tqdm(ds)):
        flat_item = flatten_dict(item)

        # 写入表头
        if not headers_written:
            headers = list(flat_item.keys())
            for col_idx, header in enumerate(headers):
                cell = xl_rowcol_to_cell(0, col_idx).upper()
                write_text_to_feishu(cell, header)
            headers_written = True

        # 写入数据行
        for col_idx, (key, value) in enumerate(flat_item.items()):
            cell = xl_rowcol_to_cell(row_idx + 1, col_idx).upper()

            # 1) 如果是 PIL 图片对象，直接写图片
            if isinstance(value, PILImage):
                value = safe_image_to_rgb(value)  # 关键：转 RGB
                buffer = io.BytesIO()
                value.save(buffer, format="PNG")
                write_image_to_feishu(
                    cell,
                    list(buffer.getvalue()),
                    f"img_{row_idx}_{col_idx}.png"
                )
                continue

            # 2) 如果是 dict / list / tuple → 序列化为 JSON 字符串
            if isinstance(value, (list, tuple, dict)):
                value = json.dumps(value, ensure_ascii=False)

            # 3) 如果是图片 URL → 下载后写图片
            if isinstance(value, str) and IMAGE_URL_PATTERN.match(value):
                try:
                    img_byte_list = download_image_as_byte_list(value)
                    write_image_to_feishu(cell, img_byte_list, f"img_{row_idx}_{col_idx}.jpg")
                    continue
                except Exception as e:
                    print(f"[图片跳过] 无法处理 URL: {value} → {e}")

            # 4) 其余按文本处理
            write_text_to_feishu(cell, str(value))


if __name__ == "__main__":
    main()