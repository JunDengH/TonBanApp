"""Unified parser for weekly schedule files (PDF preferred, Word compatible)."""
from pathlib import Path

from src.modules.word_parser import parse_weekly_schedule, scan_weekly_word


def parse_schedule_file(path: str, total_names: list[str], corrections: dict | None = None) -> dict[str, list[str]]:
    """Parse a weekly schedule from a supported PDF or Word file."""
    suffix = Path(path).suffix.lower()
    if suffix == ".docx":
        return parse_weekly_schedule(path, total_names, corrections)
    if suffix == ".pdf":
        from src.modules.pdf_schedule_parser import parse_pdf_schedule

        return parse_pdf_schedule(path, total_names, corrections)
    raise ValueError("排班文件仅支持 PDF 或 Word（.docx）。")


def scan_schedule_file(path: str, total_names: list[str], corrections: dict | None = None) -> dict:
    """Scan a supported weekly schedule file for matches and unknown names."""
    suffix = Path(path).suffix.lower()
    if suffix == ".docx":
        return scan_weekly_word(path, total_names, corrections)
    if suffix == ".pdf":
        from src.modules.pdf_schedule_parser import scan_pdf_schedule

        return scan_pdf_schedule(path, total_names, corrections)
    raise ValueError("排班文件仅支持 PDF 或 Word（.docx）。")
