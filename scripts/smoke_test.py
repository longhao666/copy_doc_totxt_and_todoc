from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _run(cmd: list[str], cwd: Path) -> None:
    subprocess.run(cmd, cwd=str(cwd), check=True)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]

    with tempfile.TemporaryDirectory() as td:
        work = Path(td)

        # 拷贝必要代码到临时目录，避免污染工作区输出目录
        shutil.copy2(repo_root / "main.py", work / "main.py")
        shutil.copy2(repo_root / "doc_text_sync.py", work / "doc_text_sync.py")
        (work / "scripts").mkdir(parents=True, exist_ok=True)
        shutil.copy2(repo_root / "scripts" / "generate_sample_tree.py", work / "scripts" / "generate_sample_tree.py")

        _run([sys.executable, "scripts/generate_sample_tree.py"], cwd=work)
        proc = subprocess.run([sys.executable, "main.py", "copy"], cwd=str(work))
        if proc.returncode not in (0, 1):
            raise RuntimeError(f"主程序退出码异常：{proc.returncode}")

        # 找到当天生成的 txt 目录，检查关键文件存在
        txt_dirs = sorted(work.glob("*-txt-copy"))
        doc_dirs = sorted(work.glob("*-doc-copy"))
        if not txt_dirs or not doc_dirs:
            raise RuntimeError("未生成日期前缀输出目录。")

        txt_root = txt_dirs[-1]
        doc_root = doc_dirs[-1]

        expected_txt = [
            txt_root / "sample.txt",
            txt_root / "sub" / "nested.txt",
        ]
        for p in expected_txt:
            if not p.exists():
                raise RuntimeError(f"缺少导出的 TXT：{p}")

        expected_docx = [
            doc_root / "sample.docx",
            doc_root / "sub" / "nested.docx",
        ]
        for p in expected_docx:
            if not p.exists():
                raise RuntimeError(f"缺少生成的 DOCX：{p}")

        print("smoke_test 通过。")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
