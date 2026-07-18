"""Output path helpers shared by report generators and UI."""
from pathlib import Path


def build_unique_output_path(output_dir, filename: str, suffix: str) -> Path:
    """Return a non-existing output path, appending `` (n)`` when needed."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    clean_name = _sanitize_filename(filename)
    clean_suffix = suffix if str(suffix).startswith(".") else f".{suffix}"
    candidate = output_dir / f"{clean_name}{clean_suffix}"
    if not candidate.exists():
        return candidate

    index = 2
    while True:
        numbered = output_dir / f"{clean_name} ({index}){clean_suffix}"
        if not numbered.exists():
            return numbered
        index += 1


def _sanitize_filename(filename: str) -> str:
    clean = (filename or "").strip()
    for char in '<>:"/\\|?*':
        clean = clean.replace(char, "_")
    clean = clean.rstrip(" .")
    return clean or "未命名文件"
