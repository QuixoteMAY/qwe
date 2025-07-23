# -*- coding: utf-8 -*-
"""
将整段流程封装为函数（支持单字符串或列表形式的范围参数）
------------------------------------------------
调用示例：
    run_image_pipeline(
        "K4VHsfEQwhV0l4tVtFPceqEXnnx",      # SPREADSHEET_TOKEN
        "bd1919",                           # SHEET_TOKEN
        "E202:E202",                        # CELL_RANGES 可单个字符串
        "F202:F202",                        # TEXT_CELL_RANGES 可单个字符串
        r"C:/Users/Admin/Desktop/image_create/临时图片",  # OUTPUT_DIR
        "G",                                # OUTPUT_COL
    )
其余逻辑与原脚本完全一致。
"""

# ===== 通用导入 =====
import json
import re
import io
import base64
from pathlib import Path

import lark_oapi as lark
from lark_oapi.api.drive.v1 import DownloadMediaRequest
from openai import OpenAI


# ───────────────────── 固定配置（保持不变） ─────────────────────
APP_ID     = "cli_a8baea0bc979d00b"
APP_SECRET = "RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX"

OPENAI_API_KEY = "AqMmbEXHndV5puKNnSXRIgXUYIBS6KLI_GPT_AK"
BASE_URL       = "https://search.bytedance.net/gpt/openapi/online/v2/crawl/openai"
MODEL_NAME     = "gpt-image-1"
# ──────────────────────────────────────────────────────────────


def run_image_pipeline(
    SPREADSHEET_TOKEN: str,
    SHEET_TOKEN: str,
    CELL_RANGES,           # 可 str 或 list[str]
    TEXT_CELL_RANGES,      # 可 str 或 list[str]
    OUTPUT_DIR: str | Path,
    OUTPUT_COL: str,
):
    """根据飞书表格内容批量合成图片并写回表格"""

    # ★ 若传入单个字符串，自动转为列表 ★
    if isinstance(CELL_RANGES, str):
        CELL_RANGES = [CELL_RANGES]
    if isinstance(TEXT_CELL_RANGES, str):
        TEXT_CELL_RANGES = [TEXT_CELL_RANGES]

    OUTPUT_DIR = Path(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # =============================================================================
    # 读取飞书表格中的【图片 file_token】
    # =============================================================================
    print("🚀 开始读取图片 file_token …")
    client = (
        lark.Client.builder()
        .app_id(APP_ID)
        .app_secret(APP_SECRET)
        .log_level(lark.LogLevel.INFO)
        .build()
    )

    range_strings = [f"{SHEET_TOKEN}!{r}" for r in CELL_RANGES]
    print(f"📋 请求图片列 ranges: {range_strings}")
    ranges_combined = ",".join(range_strings)

    request = (
        lark.BaseRequest.builder()
        .http_method(lark.HttpMethod.GET)
        .uri(f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_get")
        .token_types([lark.AccessTokenType.TENANT])
        .queries([("ranges", ranges_combined)])
        .build()
    )

    response = client.request(request)
    if not response.success():
        lark.logger.error("请求图片列失败，错误码: %s", response.code)
        return

    value_ranges = json.loads(response.raw.content.decode("utf-8"))["data"]["valueRanges"]

    file_token_dicts = {}
    for rng, vr in zip(CELL_RANGES, value_ranges):
        col = re.match(r"[A-Z]+", rng)[0]
        row0 = int(re.findall(r"\d+", rng.split(":")[0])[0])
        file_token_dicts[col] = {}
        for idx, row in enumerate(vr.get("values", [])):
            addr = f"{col}{row0 + idx}"
            tok  = row[0]["fileToken"] if row and isinstance(row[0], dict) and "fileToken" in row[0] else None
            file_token_dicts[col][addr] = tok
            print(f"  ↳ {addr}: {'✅'+tok if tok else 'None'}")

    print("✅ 图片 file_token 读取完成")
    print(json.dumps(file_token_dicts, ensure_ascii=False, indent=2))

    # =============================================================================
    # 读取【提示词】列
    # =============================================================================
    print("\n🚀 开始读取提示词 …")
    range_strings = [f"{SHEET_TOKEN}!{r}" for r in TEXT_CELL_RANGES]
    ranges_combined = ",".join(range_strings)

    request = (
        lark.BaseRequest.builder()
        .http_method(lark.HttpMethod.GET)
        .uri(f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_get")
        .token_types([lark.AccessTokenType.TENANT])
        .queries([("ranges", ranges_combined)])
        .build()
    )

    response = client.request(request)
    if not response.success():
        lark.logger.error("请求提示词失败，错误码: %s", response.code)
        return

    value_ranges = json.loads(response.raw.content.decode("utf-8"))["data"]["valueRanges"]

    text_dicts = {}
    for rng, vr in zip(TEXT_CELL_RANGES, value_ranges):
        col  = re.match(r"[A-Z]+", rng)[0]
        row0 = int(re.findall(r"\d+", rng.split(":")[0])[0])
        text_dicts[col] = {}
        for idx, row in enumerate(vr.get("values", [])):
            addr = f"{col}{row0 + idx}"
            text = row[0] if row else None
            text_dicts[col][addr] = text
            print(f"  ↳ {addr}: {text}")

    print("✅ 提示词读取完成")
    print(json.dumps(text_dicts, ensure_ascii=False, indent=2))

    # =============================================================================
    # 下载图片 & 调用 OpenAI images.edit & 写回表格
    # =============================================================================
    print("\n🚀 开始逐行处理图片合成 …")
    ai_client = OpenAI(api_key=OPENAI_API_KEY, base_url=BASE_URL)

    image_cols  = sorted(file_token_dicts.keys())
    prompt_col  = next(iter(text_dicts.keys()))

    row_nums = sorted({
        int(re.findall(r"\d+", addr)[0])
        for col in image_cols
        for addr in file_token_dicts[col]
    })
    print(f"📌 需处理行号: {row_nums}\n")

    for row in row_nums:
        print(f"── 行 {row} 开始 ──")
        tokens = [file_token_dicts[col].get(f"{col}{row}") for col in image_cols]
        valid_tokens = [tok for tok in tokens if tok]

        if not valid_tokens:
            print("   ⚠️  无任何图片 token，跳过\n")
            continue

        prompt = text_dicts[prompt_col].get(f"{prompt_col}{row}", "") or " "
        print(f"   ✏️  提示词: {prompt}")
        print(f"   📷 有效图片 token 数: {len(valid_tokens)}")

        # 下载图片
        images_io, ok = [], True
        for i, tok in enumerate(valid_tokens, 1):
            print(f"   ↓ 下载第 {i} 张图片 token={tok} …", end="")
            client = (
                lark.Client.builder()
                .app_id(APP_ID)
                .app_secret(APP_SECRET)
                .log_level(lark.LogLevel.INFO)
                .build()
            )
            req  = DownloadMediaRequest.builder().file_token(tok).extra("无").build()
            resp = client.drive.v1.media.download(req)
            if not resp.success():
                print("失败")
                ok = False
                break
            buf = io.BytesIO(resp.file.read())
            buf.name = f"{tok}.png"
            images_io.append(buf)
            print("成功")
        if not ok:
            print("   ⚠️  图片下载失败，跳过\n")
            continue

        # 调用 OpenAI —— 自动重试 3 次
        oa_resp = None
        for attempt in range(1, 4):
            print(f"   🌀 调用 OpenAI images.edit … (第 {attempt}/3 次)")
            try:
                oa_resp = ai_client.images.edit(
                    model=MODEL_NAME,
                    image=images_io,
                    prompt=prompt,
                    extra_headers={"api-key": OPENAI_API_KEY},
                )
                if oa_resp and oa_resp.data and oa_resp.data[0].b64_json:
                    print("      ✅ 成功")
                    break
                raise RuntimeError("返回数据为空")
            except Exception as e:
                print(f"      ⚠️  失败: {e}")
                if attempt == 3:
                    print("      ❌ 连续 3 次失败，放弃该行\n")
                    oa_resp = None
                else:
                    print("      ↻ 准备重试 …")
        if oa_resp is None:
            continue

        # base64 图像数据
        img_b64 = oa_resp.data[0].b64_json

        # 保存本地
        out_path = OUTPUT_DIR / f"row{row}_output.png"
        out_path.write_bytes(base64.b64decode(img_b64))
        print(f"   ✅ 已保存合成图 → {out_path}")

        # 写回飞书表格
        write_range = f"{SHEET_TOKEN}!{OUTPUT_COL}{row}:{OUTPUT_COL}{row}"
        print(f"   ⬆️  写入表格单元格 {write_range} …")
        client = (
            lark.Client.builder()
            .app_id(APP_ID)
            .app_secret(APP_SECRET)
            .log_level(lark.LogLevel.INFO)
            .build()
        )
        upload_req = (
            lark.BaseRequest.builder()
            .http_method(lark.HttpMethod.POST)
            .uri(f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_image")
            .token_types([lark.AccessTokenType.TENANT])
            .body({"range": write_range, "image": img_b64, "name": f"row{row}_output.png"})
            .build()
        )
        upload_resp = client.request(upload_req)
        if not upload_resp.success():
            print(f"   ❌ 上传失败 code={upload_resp.code}, msg={upload_resp.msg}\n")
            continue
        print("   📥 上传成功\n")

    print("🎉 全部行处理完成")



run_image_pipeline(
        "K4VHsfEQwhV0l4tVtFPceqEXnnx",      # SPREADSHEET_TOKEN
        "bd1919",                           # SHEET_TOKEN
        "E202:E202",                        # CELL_RANGES 可单个字符串
        "F202:F202",                        # TEXT_CELL_RANGES 可单个字符串
        "C:/Users/Admin/Desktop/image_create/临时图片",  # OUTPUT_DIR
        "G",                                # OUTPUT_COL
    )
















