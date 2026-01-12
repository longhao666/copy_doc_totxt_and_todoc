from __future__ import annotations

import argparse
import sys
from pathlib import Path

from doc_text_sync import run_sync


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将指定目录下的 .doc/.docx 导出为 TXT，并由 TXT 生成 DOCX（按当天日期分目录）。"
    )
    parser.add_argument(
        "source",
        nargs="?",
        default="copy",
        help="源文件夹路径（默认: copy）",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    return run_sync(Path(args.source))


if __name__ == "__main__":
    raise SystemExit(main())

