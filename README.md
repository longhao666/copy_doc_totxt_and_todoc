# copy_doc_totxt_and_todoc

将指定目录下的 Word 文档导出为 TXT，再由 TXT 生成 DOCX，输出按当天日期分目录保存。

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> 说明：`.docx` 读写默认不依赖第三方库（通过标准库 `zipfile`/`xml` 实现）。`.doc` 解析在 Windows 上通常需要安装 Microsoft Word，并可选安装 `pywin32`：`pip install pywin32`。

## 使用

```powershell
# 默认源目录为 copy
python main.py

# 指定源目录
python main.py copy
```

运行后会在当前工作目录下生成（或复用）：
- `<YYYYMMDD>-txt-<源目录名>/`：导出的 TXT
- `<YYYYMMDD>-doc-<源目录名>/`：由 TXT 生成的 DOCX

## 生成样例与冒烟验证

```powershell
python scripts\generate_sample_tree.py
python scripts\smoke_test.py
```

## 常见问题

- `.doc` 解析失败：
  - 该功能通常依赖 Windows + Microsoft Word 的自动化能力；建议优先处理 `.docx`，或先将 `.doc` 转为 `.docx`。
- 编码与换行差异：
  - TXT 统一写为 UTF-8，换行统一为 `\n`；不同编辑器显示可能略有差异。
