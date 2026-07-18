"""Locate and invoke the Tesseract runtime bundled with TonBanAPP."""
from __future__ import annotations

import csv
import io
import os
import subprocess
import sys
from pathlib import Path


def resolve_tesseract_runtime(base_dir: str | Path | None = None) -> tuple[Path, Path]:
    """Return the bundled Tesseract executable and tessdata directory."""
    if base_dir is None:
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            base_dir = Path(sys._MEIPASS)
        else:
            base_dir = Path(__file__).resolve().parent.parent.parent
    runtime = Path(base_dir) / "vendor" / "tesseract"
    executable = runtime / "tesseract.exe"
    tessdata = runtime / "tessdata"
    if not executable.exists():
        raise ValueError("内置 OCR 引擎缺失，请重新安装完整版本的 TonBanAPP。")
    if not (tessdata / "chi_sim.traineddata").exists():
        raise ValueError("内置 OCR 中文模型缺失，请重新安装完整版本的 TonBanAPP。")
    if not (tessdata / "eng.traineddata").exists():
        raise ValueError("内置 OCR 英文模型缺失，请重新安装完整版本的 TonBanAPP。")
    return executable, tessdata


def run_tesseract_tsv(
    image_path: str | Path,
    base_dir: str | Path | None = None,
    language: str = "chi_sim",
    psm: int = 11,
) -> list[dict]:
    """Run bundled Tesseract and return positioned word tokens from TSV output."""
    executable, tessdata = resolve_tesseract_runtime(base_dir)
    env = os.environ.copy()
    env["TESSDATA_PREFIX"] = str(tessdata)
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    command = [
        str(executable),
        str(image_path),
        "stdout",
        "-l",
        language,
        "--oem",
        "1",
        "--psm",
        str(psm),
        "tsv",
    ]
    try:
        completed = subprocess.run(
            command,
            cwd=str(executable.parent),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            check=False,
            creationflags=creationflags,
        )
    except OSError as exc:
        raise ValueError(f"内置 OCR 引擎无法启动: {exc}") from exc
    except subprocess.TimeoutExpired as exc:
        raise ValueError("扫描 PDF 的 OCR 识别超时，请降低页数或使用更清晰的文件。") from exc
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", errors="replace").strip()
        raise ValueError(f"扫描 PDF 的 OCR 识别失败: {detail or '未知错误'}")

    output = completed.stdout.decode("utf-8", errors="replace")
    tokens = []
    for row in csv.DictReader(io.StringIO(output), delimiter="\t"):
        text = (row.get("text") or "").strip()
        if not text or row.get("level") != "5":
            continue
        token = {"text": text}
        for key in (
            "page_num",
            "block_num",
            "par_num",
            "line_num",
            "word_num",
            "left",
            "top",
            "width",
            "height",
        ):
            try:
                token[key] = int(row.get(key, 0))
            except (TypeError, ValueError):
                token[key] = 0
        try:
            token["conf"] = float(row.get("conf", -1))
        except (TypeError, ValueError):
            token["conf"] = -1.0
        tokens.append(token)
    return tokens
