import os
import math
import re
import json
import hashlib
from io import BytesIO
from pathlib import Path

# os.environ["HF_TOKEN"] = "hf_..."

import requests
from datasets import load_dataset
from PIL import Image
from tqdm import tqdm
from xlsxwriter import Workbook

# 图片相关设置
desired_width, desired_height = 100, 100  # 图片显示宽度和高度
# 数据集相关设置
dataset_path = "derek-thomas/ScienceQA"  # Hugging Face 数据集路径
dataset_format = "hf"  # 数据集格式：hf / json / jsonl / local_json / local_jsonl
local_file_path = "C:/Users/Admin/Downloads/sample_data.json"  # 本地 JSON 或 JSONL 文件路径
save_path = Path("C:/Users/Admin/Desktop/tt.xlsx")  # 保存路径

IMAGE_URL_PATTERN = re.compile(r"https?://.*\.(png|jpg|jpeg|gif|webp|bmp)", re.IGNORECASE)


def download_image(url: str) -> Image.Image:
    """
    从 URL 下载图片并缓存到本地。

    Args:
        url (str): 图片的 URL。

    Returns:
        Image.Image: 下载的图片对象。
    """
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
    """
    将嵌套字典展平。

    Args:
        d (dict): 输入字典。
        parent_key (str): 父键。
        sep (str): 分隔符。

    Returns:
        dict: 展平后的字典。
    """
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
    """
    将数据保存到 Excel 文件。

    Args:
        items (list[dict]): 数据项列表。
        has_image (bool): 是否包含图片。
        save_path (Path): 保存路径。
        chunk_idx (int): 分块索引。
    """
    workbook = Workbook(save_path.with_suffix(f".{chunk_idx}.xlsx"))
    worksheet = workbook.add_worksheet()

    header_format = workbook.add_format({"bold": True, "border": 1})
    headers = items[0].keys()
    for i, header in enumerate(headers):
        worksheet.write(0, i, header, header_format)

    for row, item in enumerate(items):
        for col, value in enumerate(item.values()):
            if isinstance(value, Image.Image):
                width, height = value.size
                x_scale = desired_width / width
                y_scale = desired_height / height

                col_width = desired_width / 7
                row_height = desired_height

                worksheet.set_column(col, col, col_width)
                worksheet.set_row_pixels(row + 1, row_height)

                image_data = BytesIO()
                if value.mode in ("RGBA", "P"):
                    value = value.convert("RGB")
                value.save(image_data, format="JPEG")
                image_data.seek(0)

                worksheet.insert_image(
                    row + 1, col, "",
                    {"image_data": image_data, "x_scale": x_scale, "y_scale": y_scale}
                )
            else:
                worksheet.write(row + 1, col, value)

    workbook.close()


def load_hf_dataset():
    """
    加载 Hugging Face 数据集。

    Returns:
        datasets.Dataset: 加载的数据集。
    """
    split = "train"
    return load_dataset(dataset_path, split=split, trust_remote_code=True, streaming=True)


def load_local_json_file(path: str) -> list[dict]:
    """
    加载本地 JSON 文件。

    Args:
        path (str): 文件路径。

    Returns:
        list[dict]: 加载的数据列表。
    """
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
    """
    加载本地 JSONL 文件。

    Args:
        path (str): 文件路径。

    Returns:
        list[dict]: 加载的数据列表。
    """
    with open(path, "r", encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


def main():
    """
    主函数。
    """
    chunk_size = 10
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
                    print(f"Error downloading image {value}: {e}")
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
            print(f"Saving chunk {chunk_idx} of {total_chunks}")
            save_items(items, has_image, save_path, chunk_idx)
            items = []

    if len(items) > 0:
        chunk_idx += 1
        print(f"Saving chunk {chunk_idx} of {total_chunks}")
        save_items(items, has_image, save_path, chunk_idx)


if __name__ == "__main__":
    main()