# doc-text-sync Specification

## Purpose
TBD - created by archiving change add-doc-text-sync. Update Purpose after archive.
## Requirements
### Requirement: 按日期创建同步目录
系统 MUST 接受一个源文件夹名称（默认 `copy`），并在工作目录下创建两个同级输出目录：`<YYYYMMDD>-txt-<源>` 与 `<YYYYMMDD>-doc-<源>`；当目录已存在时 MUST 复用，不得删除已有数据。

#### Scenario: 为指定源目录创建当日输出目录
- **GIVEN** 操作员在 2026-01-08 运行脚本，目标源目录为 `copy`
- **WHEN** 脚本初始化
- **THEN** 它确保 `20260108-txt-copy/` 与 `20260108-doc-copy/` 存在，并在后续步骤中按 `copy/` 的相对路径结构写入输出文件。

### Requirement: 将 DOC/DOCX 内容导出为 TXT
系统 MUST 递归发现源文件夹中的 `.doc` 与 `.docx` 文件，提取其纯文本内容，并在当日 TXT 输出目录中按相同相对路径与同名基名写出 UTF-8 编码的 `.txt` 文件。

#### Scenario: 转换 DOCX 为 TXT 并保留目录结构
- **GIVEN** 存在 `copy/marketing/Q1.docx`
- **WHEN** 脚本在同一天运行
- **THEN** 它创建/更新 `20260108-txt-copy/marketing/Q1.txt`，其内容为该文档的文本内容。

#### Scenario: 单文件失败不影响整体处理
- **GIVEN** 某个 `.doc` 文件无法解析
- **WHEN** 脚本遍历文件列表
- **THEN** 它记录失败路径并继续处理其他文档，并在最终汇总中报告失败列表。

### Requirement: 由 TXT 生成 DOCX
系统 MUST 遍历当日 TXT 输出目录中的每个 `.txt` 文件，并在 DOC 输出目录中生成同名 `.docx`（相对路径一致），将 TXT 内容写入 DOCX 文档正文。

#### Scenario: 根据 TXT 编辑结果生成 DOCX
- **GIVEN** 操作员编辑了 `20260108-txt-copy/marketing/Q1.txt`
- **WHEN** 执行生成 DOCX 步骤
- **THEN** 脚本写出 `20260108-doc-copy/marketing/Q1.docx`，其内容反映 TXT 的最新文本。

