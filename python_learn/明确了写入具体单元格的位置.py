import os
import math
import re
import json
import hashlib
from io import BytesIO
from pathlib import Path

import requests
from datasets import load_dataset
from PIL import Image
from tqdm import tqdm
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell

# 图片相关设置
desired_width, desired_height = 100, 100  # 图片显示宽度和高度

# 数据集相关设置
dataset_path = "derek-thomas/ScienceQA"  # Hugging Face 数据集路径
dataset_format = "hf"  # 数据集格式：hf / json / jsonl / local_json / local_jsonl
local_file_path = "C:/Users/Admin/Downloads/sample_data.json"  # 本地 JSON 或 JSONL 文件路径
save_path = Path("C:/Users/Admin/Desktop/tt.xlsx")  # 保存路径

IMAGE_URL_PATTERN = re.compile(r"https?://.*\.(png|jpg|jpeg|gif|webp|bmp)", re.IGNORECASE)


def download_image(url: str) -> Image.Image:
    cache_dir = Path("cache")
    cache_dir.mkdir(parents=True, exist_ok=True)

    cache_filename = hashlib.sha256(url.encode()).hexdigest() + ".jpg"
    cache_path = cache_dir / cache_filename
    if cache_path.exists():
        image = Image.open(cache_path)
    else:
        headers = {
            "User-Agent": "MyApp/1.0 (https://example.com ; myemail@example.com)"
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if not content_type.startswith("image/"):
            raise ValueError(f"URL {url} did not return an image. Content-Type: {content_type}")

        try:
            image = Image.open(BytesIO(response.content))
            image.save(cache_path)
        except Exception as e:
            raise ValueError(f"Cannot identify image file from URL {url}: {e}")

    return image


def flatten_dict(d: dict, parent_key: str = '', sep: str = '.') -> dict:
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def save_items(
    items: list[dict],
    has_image: bool,
    save_path: Path,
    chunk_idx: int,
):
    workbook = Workbook(save_path.with_suffix(f".{chunk_idx}.xlsx"))
    worksheet = workbook.add_worksheet()

    header_format = workbook.add_format({"bold": True, "border": 1})
    headers = list(items[0].keys())

    # 写入表头
    for col, header in enumerate(headers):
        cell_name = xl_rowcol_to_cell(0, col)
        worksheet.write(0, col, header, header_format)
        print(f"[表头] 写入单元格 {cell_name}: {header}")

    # 写入内容
    for row, item in enumerate(items):
        for col, value in enumerate(item.values()):
            excel_row = row + 1  # 数据写入从第2行开始
            cell_name = xl_rowcol_to_cell(excel_row, col)

            if isinstance(value, Image.Image):
                width, height = value.size
                x_scale = desired_width / width
                y_scale = desired_height / height

                worksheet.set_column(col, col, desired_width / 7)
                worksheet.set_row_pixels(excel_row, desired_height)

                image_data = BytesIO()
                if value.mode in ("RGBA", "P"):
                    value = value.convert("RGB")
                value.save(image_data, format="JPEG")
                image_data.seek(0)

                worksheet.insert_image(
                    excel_row, col, "",
                    {"image_data": image_data, "x_scale": x_scale, "y_scale": y_scale}
                )
                print(f"[图片] 写入单元格 {cell_name}: 图片尺寸 {width}x{height}，缩放到 {desired_width}x{desired_height}")
            else:
                worksheet.write(excel_row, col, value)
                print(f"[文字] 写入单元格 {cell_name}: {value}")

    workbook.close()


def load_hf_dataset():
    split = "train"
    return load_dataset(dataset_path, split=split, trust_remote_code=True, streaming=True)


def load_local_json_file(path: str) -> list[dict]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
        if isinstance(data, dict):
            for v in data.values():
                if isinstance(v, list):
                    return v
            raise ValueError("JSON dict 中未找到可用于处理的主数组")
        elif isinstance(data, list):
            return data
        else:
            raise ValueError("不支持的 JSON 文件结构")


def load_local_jsonl_file(path: str) -> list[dict]:
    with open(path, "r", encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


def main():
    chunk_size = 20
    total_items = 20
    auto_download_image = True

    save_path.parent.mkdir(parents=True, exist_ok=True)

    ds = None
    if dataset_format == "hf":
        ds = load_hf_dataset()
        ds = ds.take(total_items)
    elif dataset_format == "local_json":
        records = load_local_json_file(local_file_path)
        ds = records[:total_items]
    elif dataset_format == "local_jsonl":
        records = load_local_jsonl_file(local_file_path)
        ds = records[:total_items]
    else:
        raise ValueError(f"Invalid dataset format: {dataset_format}")

    items = []
    has_image = False
    image_idx = 0
    chunk_idx = 0
    total_chunks = math.ceil(total_items / chunk_size)

    iterator = ds if isinstance(ds, list) else iter(ds)

    for item in tqdm(iterator):
        item = flatten_dict(item)
        new_item = {}
        for key, value in item.items():
            if isinstance(value, (list, tuple, dict)):
                value = json.dumps(value, ensure_ascii=False)

            if (
                isinstance(value, str)
                and auto_download_image
                and IMAGE_URL_PATTERN.match(value)
            ):
                try:
                    image = download_image(value)
                except Exception as e:
                    print(f"[错误] 下载图片失败 {value}: {e}")
                    image = None

                if image:
                    image_key = f"dl_image_{image_idx}"
                    new_item[image_key] = image
                    image_idx += 1
                    has_image = True

            new_item[key] = value

        items.append(new_item)

        if len(items) >= chunk_size:
            chunk_idx += 1
            print(f"\n[保存] Chunk {chunk_idx}/{total_chunks}")
            save_items(items, has_image, save_path, chunk_idx)
            items = []

    if len(items) > 0:
        chunk_idx += 1
        print(f"\n[保存] Chunk {chunk_idx}/{total_chunks}")
        save_items(items, has_image, save_path, chunk_idx)


if __name__ == "__main__":
    main()
