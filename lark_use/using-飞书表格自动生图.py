# -*- coding: utf-8 -*-
"""
升级版脚本
------------------------------------------------
新增功能：
1. 把 OpenAI 生成的合成图（base64）写回飞书表格指定列
2. 图生图失败自动重试 3 次
3. 行内无图片 → 自动改用文生图接口
4. ★ 当 CELL_RANGES 为空（[]）时，完全跳过“下载图片”请求，直接
   将 TEXT_CELL_RANGES 的提示词依次走文生图接口
5. ★ 本次补丁：支持 1-N 套 CELL/TEXT/OUTPUT 配置，程序依次执行；
   如某套 TEXT_CELL_RANGES_✖︎ 为空，则自动跳过
其余逻辑、注释完全保持不变。
"""

# ===== 通用导入 =====
import json, re, io, base64, sys
from pathlib import Path

import lark_oapi as lark
from lark_oapi.api.drive.v1 import DownloadMediaRequest
from openai import OpenAI


# ──────────────────────  固定参数区  ──────────────────────
SPREADSHEET_TOKEN = "AMDss2pkQhH0eBtRxBYcvRuhnEh"     #必填
SHEET_TOKEN       = "255e28"                          #必填

OUTPUT_DIR = Path(r"C:/Users/Admin/Desktop/image_create/临时图片")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

APP_ID     = "cli_a8baea0bc979d00b"
APP_SECRET = "RT0DDhjaWLRbXK0fzA5oNc4wL2huzXKX"

OPENAI_API_KEY = "AqMmbEXHndV5puKNnSXRIgXUYIBS6KLI_GPT_AK"
BASE_URL       = "https://search.bytedance.net/gpt/openapi/online/v2/crawl/openai"
MODEL_NAME     = "gpt-image-1"
# ──────────────────────────────────────────────────────────


# ──────────────────────  可配置参数区（最多 3 套） ──────────────────────
# -------- 套 1 --------
CELL_RANGES_1       = []                # ← 为空时只走文生图
TEXT_CELL_RANGES_1  = []
OUTPUT_COL_1        = ""

# -------- 套 2 --------
CELL_RANGES_2       = ["C923:C954"]
TEXT_CELL_RANGES_2  = ["D923:D954"]
OUTPUT_COL_2        = "E"

# -------- 套 3 --------
CELL_RANGES_3       = ["E700:E954"]        # ← 全文生图
TEXT_CELL_RANGES_3  = ["F700:F954"]
OUTPUT_COL_3        = "G"

CELL_RANGES_4       = ["G700:G954"]                # ← 为空时只走文生图
TEXT_CELL_RANGES_4  = ["H700:H954"]
OUTPUT_COL_4        = "I"

# -------- 套 2 --------
CELL_RANGES_5       = []
TEXT_CELL_RANGES_5  = []
OUTPUT_COL_5        = ""

# -------- 套 3 --------
CELL_RANGES_6       = []        # ← 全文生图
TEXT_CELL_RANGES_6  = []
OUTPUT_COL_6        = ""


# （如需更多套，请仿照以上格式继续向下编号：4、5、6 …）
# ──────────────────────────────────────────────────────────


# ★ 自动收集全部 *_N 变量，支持任意套数 ★
ALL_CONFIGS = []
for i in range(1, 100):                        # 理论上支持 99 套
    text_key = f"TEXT_CELL_RANGES_{i}"
    if text_key not in globals():
        break                                  # 编号中断即停止搜索
    cell_ranges  = globals().get(f"CELL_RANGES_{i}", [])
    text_ranges  = globals().get(text_key, [])
    output_col   = globals().get(f"OUTPUT_COL_{i}")
    # 仅当同时存在提示词和输出列时才加入执行列表
    if text_ranges and output_col:
        ALL_CONFIGS.append((cell_ranges, text_ranges, output_col))

CONFIG_SETS = ALL_CONFIGS                      # TEXT 已保证非空

if not CONFIG_SETS:
    print("⚠️  未找到可执行的任务套数（TEXT_CELL_RANGES_* 均为空），程序结束。")
    sys.exit(0)


# ════════════════ 主循环：依次执行每套配置 ════════════════
ai_client_global = OpenAI(api_key=OPENAI_API_KEY, base_url=BASE_URL)

for cfg_idx, (CELL_RANGES, TEXT_CELL_RANGES, OUTPUT_COL) in enumerate(CONFIG_SETS, start=1):
    print(f"\n==================== 🗂️ 开始任务组 {cfg_idx}，输出列 {OUTPUT_COL} ====================")

    #############################################################################################################################################################################################################
    # =============================================================================
    # 读取飞书表格中的【图片 file_token】
    # =============================================================================
    file_token_dicts = {}
    if CELL_RANGES:
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
            continue

        value_ranges = json.loads(response.raw.content.decode("utf-8"))["data"]["valueRanges"]

        for rng, vr in zip(CELL_RANGES, value_ranges):
            col = re.match(r"[A-Z]+", rng)[0]
            row0 = int(re.findall(r"\d+", rng.split(":")[0])[0])
            file_token_dicts[col] = {}
            for i, row in enumerate(vr.get("values", [])):
                addr = f"{col}{row0 + i}"
                tok  = row[0]["fileToken"] if row and isinstance(row[0], dict) and "fileToken" in row[0] else None
                file_token_dicts[col][addr] = tok
                print(f"  ↳ {addr}: {'✅'+tok if tok else 'None'}")

        print("✅ 图片 file_token 读取完成")
        print(json.dumps(file_token_dicts, ensure_ascii=False, indent=2))
    else:
        print("🚀 CELL_RANGES 为空，跳过图片下载步骤，全部走文生图接口")
    #############################################################################################################################################################################################################

    # =============================================================================
    # 读取【提示词】列
    # =============================================================================
    print("\n🚀 开始读取提示词 …")
    client = (
        lark.Client.builder()
        .app_id(APP_ID)
        .app_secret(APP_SECRET)
        .log_level(lark.LogLevel.INFO)
        .build()
    )
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
        continue

    value_ranges = json.loads(response.raw.content.decode("utf-8"))["data"]["valueRanges"]
    text_dicts = {}
    for rng, vr in zip(TEXT_CELL_RANGES, value_ranges):
        col  = re.match(r"[A-Z]+", rng)[0]
        row0 = int(re.findall(r"\d+", rng.split(":")[0])[0])
        text_dicts[col] = {}
        for i, row in enumerate(vr.get("values", [])):
            addr = f"{col}{row0 + i}"
            text = row[0] if row else None
            text_dicts[col][addr] = text
            print(f"  ↳ {addr}: {text}")

    print("✅ 提示词读取完成")
    print(json.dumps(text_dicts, ensure_ascii=False, indent=2))

    #############################################################################################################################################################################################################
    # =============================================================================
    # 下载图片 & 调用 OpenAI images.* & 写回表格
    # =============================================================================
    print("\n🚀 开始逐行处理图片合成 …")
    image_cols  = sorted(file_token_dicts.keys())
    prompt_col  = next(iter(text_dicts.keys()))

    row_nums_img = {
        int(re.findall(r"\d+", addr)[0])
        for col in image_cols
        for addr in file_token_dicts.get(col, {})
    }
    row_nums_txt = {
        int(re.findall(r"\d+", addr)[0])
        for addr in text_dicts[prompt_col]
    }
    row_nums = sorted(row_nums_img | row_nums_txt)
    print(f"📌 需处理行号: {row_nums}\n")

    for row in row_nums:
        print(f"── 行 {row} 开始 ──")
        tokens = [file_token_dicts.get(col, {}).get(f"{col}{row}") for col in image_cols] if image_cols else []
        valid_tokens = [tok for tok in tokens if tok]

        prompt = text_dicts[prompt_col].get(f"{prompt_col}{row}", "") or " "
        img_b64 = None  # 最终统一变量

        # =====================================================================
        # 有图片 → images.edit
        # =====================================================================
        if valid_tokens:
            print(f"   ✏️  提示词: {prompt}")
            print(f"   📷 有效图片 token 数: {len(valid_tokens)}")

            images_io, ok = [], True
            for tok in valid_tokens:
                client_dl = (
                    lark.Client.builder()
                    .app_id(APP_ID)
                    .app_secret(APP_SECRET)
                    .log_level(lark.LogLevel.INFO)
                    .build()
                )
                resp = client_dl.drive.v1.media.download(
                    DownloadMediaRequest.builder().file_token(tok).extra("无").build()
                )
                if not resp.success():
                    ok = False
                    break
                buf = io.BytesIO(resp.file.read())
                buf.name = f"{tok}.png"
                images_io.append(buf)
            if not ok:
                print("   ⚠️  图片下载失败，跳过\n")
                continue

            for attempt in range(1, 4):
                print(f"   🌀 调用 OpenAI images.edit … (第 {attempt}/3 次)")
                try:
                    oa_resp = ai_client_global.images.edit(
                        model=MODEL_NAME,
                        image=images_io,
                        prompt=prompt,
                        extra_headers={"api-key": OPENAI_API_KEY},
                    )
                    if oa_resp and oa_resp.data and oa_resp.data[0].b64_json:
                        img_b64 = oa_resp.data[0].b64_json
                        print("      ✅ 成功")
                        break
                    raise RuntimeError("返回数据为空")
                except Exception as e:
                    print(f"      ⚠️  失败: {e}")
                    if attempt == 3:
                        print("      ❌ 连续 3 次失败，放弃该行\n")
                    else:
                        print("      ↻ 准备重试 …")

        # =====================================================================
        # 无图片 → images.generate
        # =====================================================================
        else:
            print("   ⚠️  当前行无图片 token，将使用文生图接口")
            print(f"   ✏️  提示词: {prompt}")
            for attempt in range(1, 4):
                print(f"   🌀 调用 OpenAI images.generate … (第 {attempt}/3 次)")
                try:
                    oa_resp = ai_client_global.images.generate(
                        model=MODEL_NAME,
                        prompt=prompt,
                        extra_headers={"api-key": OPENAI_API_KEY},
                    )
                    if oa_resp and oa_resp.data and oa_resp.data[0].b64_json:
                        img_b64 = oa_resp.data[0].b64_json
                        print("      ✅ 成功")
                        break
                    raise RuntimeError("返回数据为空")
                except Exception as e:
                    print(f"      ⚠️  失败: {e}")
                    if attempt == 3:
                        print("      ❌ 连续 3 次失败，放弃该行\n")
                    else:
                        print("      ↻ 准备重试 …")

        # 若依旧未生成图片则跳过
        if img_b64 is None:
            continue

        # =====================================================================
        # 保存本地 & 写回飞书表格
        # =====================================================================
        out_path = OUTPUT_DIR / f"group{cfg_idx}_row{row}.png"
        out_path.write_bytes(base64.b64decode(img_b64))
        print(f"   ✅ 已保存合成图 → {out_path}")

        write_range = f"{SHEET_TOKEN}!{OUTPUT_COL}{row}:{OUTPUT_COL}{row}"
        print(f"   ⬆️  写入表格单元格 {write_range} …")
        client_up = (
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
            .body({"range": write_range, "image": img_b64, "name": f"group{cfg_idx}_row{row}.png"})
            .build()
        )
        upload_resp = client_up.request(upload_req)
        if upload_resp.success():
            print("   📥 上传成功\n")
        else:
            print(f"   ❌ 上传失败 code={upload_resp.code}, msg={upload_resp.msg}\n")

    print(f"✅ 任务组 {cfg_idx} 完成\n")

print("🎉 全部行处理完成")
#############################################################################################################################################################################################################








