from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from doc_text_sync import write_docx_text

def _write_rtf_doc(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rtf = r"{\rtf1\ansi" + "\n" + text.replace("\n", r"\par" + "\n") + "\n}"
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(rtf)


def _write_docx(path: Path, text: str) -> None:
    write_docx_text(path, text)


def main() -> int:
    root = Path("copy")
    root.mkdir(parents=True, exist_ok=True)

    # docx
    _write_docx(root / "sample.docx", "这是一个 DOCX 样例。\n第二行。")
    _write_docx(root / "sub" / "nested.docx", "子目录里的 DOCX。")

    # doc（这里用 RTF 内容 + .doc 扩展名，便于在有 Word 的机器上打开/解析）
    _write_rtf_doc(root / "sample.doc", "这是一个 DOC 样例（RTF 伪装）。\n第二行。")

    with (root / "README.md").open("w", encoding="utf-8", newline="\n") as f:
        f.write(
            "该目录用于放置待处理的 .doc/.docx 文件。\n"
            "你也可以运行 scripts/generate_sample_tree.py 自动生成样例。\n"
        )

    print("已生成样例目录：copy/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
