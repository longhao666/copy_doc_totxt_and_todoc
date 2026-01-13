from __future__ import annotations

import datetime as _dt
import os
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as _ET


@dataclass(frozen=True)
class SyncResult:
    docs_total: int
    docs_ok: int
    txt_total: int
    txt_ok: int
    failures: list[str]


def _today_yyyymmdd() -> str:
    return _dt.date.today().strftime("%Y%m%d")


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _read_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as zf:
        try:
            xml_bytes = zf.read("word/document.xml")
        except KeyError as exc:
            raise RuntimeError(f"无效的 .docx（缺少 word/document.xml）：{path}") from exc

    root = _ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    paras: list[str] = []
    for p in root.findall(".//w:p", ns):
        texts: list[str] = []
        for t in p.findall(".//w:t", ns):
            if t.text:
                texts.append(t.text)
        paras.append("".join(texts))

    return "\n".join(paras).rstrip("\n")


def _read_doc_text_via_word(path: Path) -> str:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "解析 .doc 需要 Windows + Microsoft Word（或兼容组件）以及可用的 pywin32。"
            "你可以尝试：pip install pywin32；若仍失败，请仅处理 .docx 或先将 .doc 转为 .docx。"
        ) from exc

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(path), ReadOnly=True)
        text = doc.Content.Text
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()

    return (text or "").replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def _read_doc_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _read_docx_text(path)
    if suffix == ".doc":
        return _read_doc_text_via_word(path)
    raise ValueError(f"不支持的文件类型: {path.name}")


def _write_utf8(path: Path, text: str) -> None:
    _ensure_dir(path.parent)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def _write_docx(path: Path, text: str) -> None:
    _ensure_dir(path.parent)
    _write_minimal_docx(path, text)


def _write_minimal_docx(path: Path, text: str) -> None:
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
    ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

    _ET.register_namespace("w", ns_w)
    _ET.register_namespace("r", ns_r)

    document = _ET.Element(f"{{{ns_w}}}document")
    body = _ET.SubElement(document, f"{{{ns_w}}}body")

    lines = text.splitlines()
    if not lines:
        lines = [""]

    for line in lines:
        p = _ET.SubElement(body, f"{{{ns_w}}}p")
        r = _ET.SubElement(p, f"{{{ns_w}}}r")
        t = _ET.SubElement(r, f"{{{ns_w}}}t")
        if line.startswith(" ") or line.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line

    sect_pr = _ET.SubElement(body, f"{{{ns_w}}}sectPr")
    _ET.SubElement(sect_pr, f"{{{ns_w}}}pgSz", {f"{{{ns_w}}}w": "11906", f"{{{ns_w}}}h": "16838"})
    _ET.SubElement(sect_pr, f"{{{ns_w}}}pgMar", {
        f"{{{ns_w}}}top": "1440",
        f"{{{ns_w}}}right": "1440",
        f"{{{ns_w}}}bottom": "1440",
        f"{{{ns_w}}}left": "1440",
        f"{{{ns_w}}}header": "708",
        f"{{{ns_w}}}footer": "708",
        f"{{{ns_w}}}gutter": "0",
    })

    document_xml = _ET.tostring(document, encoding="utf-8", xml_declaration=True)

    content_types = _ET.Element(f"{{{ns_ct}}}Types")
    _ET.SubElement(
        content_types,
        f"{{{ns_ct}}}Default",
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    _ET.SubElement(content_types, f"{{{ns_ct}}}Default", {"Extension": "xml", "ContentType": "application/xml"})
    _ET.SubElement(
        content_types,
        f"{{{ns_ct}}}Override",
        {
            "PartName": "/word/document.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        },
    )
    content_types_xml = _ET.tostring(content_types, encoding="utf-8", xml_declaration=True)

    rels = _ET.Element(f"{{{ns_rel}}}Relationships")
    _ET.SubElement(
        rels,
        f"{{{ns_rel}}}Relationship",
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "word/document.xml",
        },
    )
    rels_xml = _ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)


def write_docx_text(path: Path, text: str) -> None:
    _ensure_dir(path.parent)
    _write_minimal_docx(path, text)


def _relpath(child: Path, parent: Path) -> Path:
    return child.relative_to(parent)


def _iter_docs(source_dir: Path) -> list[Path]:
    docs: list[Path] = []
    for path in source_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() in {".doc", ".docx"}:
            docs.append(path)
    return docs


def _iter_txts(txt_root: Path) -> list[Path]:
    return [p for p in txt_root.rglob("*.txt") if p.is_file()]


def run_sync(source_dir: Path) -> int:
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"错误：源文件夹不存在或不是目录：{source_dir}")
        return 2

    source_dir = source_dir.resolve()
    source_name = source_dir.name
    date_prefix = _today_yyyymmdd()

    cwd = Path(os.getcwd()).resolve()
    txt_root = cwd / f"{date_prefix}-txt-{source_name}"
    doc_root = cwd / f"{date_prefix}-doc-{source_name}"
    _ensure_dir(txt_root)
    _ensure_dir(doc_root)

    failures: list[str] = []

    docs = _iter_docs(source_dir)
    docs_ok = 0
    for doc_path in docs:
        rel = _relpath(doc_path, source_dir)
        out_txt = (txt_root / rel).with_suffix(".txt")
        try:
            text = _read_doc_text(doc_path)
            _write_utf8(out_txt, text)
            docs_ok += 1
        except Exception as exc:
            failures.append(f"DOC->TXT 失败：{doc_path}（{exc}）")

    txts = _iter_txts(txt_root)
    txt_ok = 0
    for txt_path in txts:
        rel = _relpath(txt_path, txt_root)
        out_docx = (doc_root / rel).with_suffix(".docx")
        try:
            text = txt_path.read_text(encoding="utf-8")
            _write_docx(out_docx, text)
            txt_ok += 1
        except Exception as exc:
            failures.append(f"TXT->DOCX 失败：{txt_path}（{exc}）")

    print("完成。")
    print(f"- 源文档数（.doc/.docx）：{len(docs)}，成功：{docs_ok}，失败：{len(docs) - docs_ok}")
    print(f"- TXT 数：{len(txts)}，成功生成 DOCX：{txt_ok}，失败：{len(txts) - txt_ok}")
    if failures:
        print("- 失败列表：")
        for item in failures:
            print(f"  - {item}")
        return 1
    return 0
