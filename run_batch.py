#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
"""
一键批量处理：2015-2024 年年报
流程：爬取链接（缺则爬）→ 下载 PDF → 转 TXT → 存入 KESU 硬盘
"""

import importlib.util
import logging
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# ==================== 配置 ====================

START_YEAR = 2015
END_YEAR = 2024
GZH = "【公众号：凌小添】"

SCRIPT_DIR = Path(__file__).parent
KESU_BASE = Path("/Volumes/KESU/年报")

DELETE_PDF = True   # 转换后删除 PDF（节省硬盘空间）
MAX_RETRIES = 3
TIMEOUT = 15
PROCESSES = None    # None = 自动（使用 CPU 核数）

# =============================================


def _load_module(filename: str, module_name: str):
    path = SCRIPT_DIR / filename
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def xlsx_path(year: int) -> Path:
    return SCRIPT_DIR / f"年报链接_{year}{GZH}.xlsx"


def crawl_year(year: int) -> bool:
    """若该年 xlsx 不存在则爬取巨潮资讯获取链接，存为本地 xlsx。"""
    out = xlsx_path(year)
    if out.exists():
        logging.info(f"{year} 年链接文件已存在，跳过爬取: {out.name}")
        return True

    logging.info(f"开始爬取 {year} 年年报链接（约需数分钟）...")
    crawler_mod = _load_module("1.report_link_crawler.py", "crawler")
    config = crawler_mod.CrawlerConfig(
        target_year=year,
        exclude_keywords=["英文", "已取消", "摘要"],
        plate="sz;sh",
        max_retries=MAX_RETRIES,
        timeout=TIMEOUT,
        output_dir=str(SCRIPT_DIR),
    )
    crawler_mod.AnnualReportCrawler(config).run()

    if not out.exists():
        logging.error(f"{year} 年爬取结束但未生成文件: {out.name}")
        return False
    return True


def convert_year(year: int) -> None:
    """读取该年 xlsx，下载 PDF 并转为 TXT，存入 KESU 对应年份目录。"""
    xlsx = xlsx_path(year)
    if not xlsx.exists():
        logging.error(f"找不到 {year} 年链接文件，跳过: {xlsx.name}")
        return

    year_dir = KESU_BASE / str(year)
    pdf_tmp_dir = year_dir / "_pdf_tmp"  # PDF 临时目录，转换后删除

    converter_mod = _load_module("2.pdf_batch_converter.py", "converter")
    config = converter_mod.ConverterConfig(
        excel_file=str(xlsx),
        pdf_dir=str(pdf_tmp_dir),
        txt_dir=str(year_dir),
        target_year=year,
        delete_pdf=DELETE_PDF,
        max_retries=MAX_RETRIES,
        timeout=TIMEOUT,
        processes=PROCESSES,
        upload_to_gdrive=False,
    )
    converter_mod.AnnualReportProcessor(config).run()

    # 若 _pdf_tmp 已空则清理目录
    if pdf_tmp_dir.exists() and not any(pdf_tmp_dir.iterdir()):
        pdf_tmp_dir.rmdir()


def main() -> None:
    # 检查 KESU 是否挂载
    if not Path("/Volumes/KESU").exists():
        logging.error("KESU 硬盘未挂载，请先连接硬盘再运行。")
        sys.exit(1)

    KESU_BASE.mkdir(parents=True, exist_ok=True)
    logging.info(f"输出根目录: {KESU_BASE}")
    logging.info(f"处理年份范围: {START_YEAR} - {END_YEAR}（共 {END_YEAR - START_YEAR + 1} 年）")

    # ── 第一步：确保每年的链接 xlsx 都存在 ──
    logging.info("\n" + "=" * 60)
    logging.info("第一步：获取各年年报链接")
    logging.info("=" * 60)
    for year in range(START_YEAR, END_YEAR + 1):
        crawl_year(year)

    # ── 第二步：下载 PDF + 转 TXT ──
    logging.info("\n" + "=" * 60)
    logging.info("第二步：下载 PDF 并转换为 TXT")
    logging.info("=" * 60)
    for year in range(START_YEAR, END_YEAR + 1):
        logging.info(f"\n{'─' * 40}")
        logging.info(f"处理 {year} 年年报 → {KESU_BASE / str(year)}")
        logging.info(f"{'─' * 40}")
        convert_year(year)

    logging.info("\n" + "=" * 60)
    logging.info("全部完成！")
    logging.info(f"TXT 文件保存路径: {KESU_BASE}")
    logging.info("=" * 60)


if __name__ == "__main__":
    main()
