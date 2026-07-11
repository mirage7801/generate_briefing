# AI+硬件协同设计愿景自动化简报生成器

这是一个读取 AI+硬件协同设计愿景论文，并自动生成 `briefing.pptx` 的 Python 工具。程序支持 PDF 和 TXT 输入，会先做文档切分与 NIL 预处理，再调用 OpenAI API 提取结构化 briefing JSON，最后用 `python-pptx` 生成三页专业科技风简报。

## 文件结构

```text
generate_briefing.py
requirements.txt
README.md
briefing.pptx
```

`briefing.pptx` 会在程序运行成功后生成。

## 运行环境

建议使用 Python 3.10+。

本项目源码与文档统一使用 UTF-8 编码。若在旧版 Windows PowerShell 中查看中文出现乱码，可先执行：

```powershell
chcp 65001
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 设置 API Key

Windows PowerShell：

```powershell
$env:OPENAI_API_KEY="你的 API Key"
```

macOS / Linux：

```bash
export OPENAI_API_KEY="你的 API Key"
```

可选：通过环境变量修改模型名称。

```bash
export OPENAI_MODEL="gpt-4o-mini"
```

## 运行命令

```bash
python generate_briefing.py --input paper.pdf --output briefing.pptx
```

或：

```bash
python generate_briefing.py --input paper.txt --output briefing.pptx
```

也可以调整 chunk 参数：

```bash
python generate_briefing.py --input paper.pdf --output briefing.pptx --max-chars 3000 --overlap 250
```

生成成功后，终端会输出：

```text
briefing.pptx generated successfully.
```

## Prompt 设计思路

- 让 LLM 扮演资深 AI+硬件协同设计技术分析师。
- 不直接对整篇论文做最终总结。
- 先做 NIL 预处理。
- NIL = Named terms + Important actions + Literals/Data。
- NIL 先保留核心专有名词、关键动词关系、重要数据。
- 再合并 NIL 结果。
- 最后生成符合固定 JSON Schema 的 briefing JSON。
- 通过 JSON Schema 约束输出。
- 通过 JSON repair 机制提高稳定性。
- 通过语义提取覆盖任务要求，而不是写死关键词。

## 文本切分策略

- 使用段落累积切分。
- 保留章节标题。
- 超长段落按句子边界切分。
- 相邻 chunk 保留 overlap。
- 每个 chunk 保留 `chunk_id`，方便追踪来源。

## 遇到的问题与解决方案

- PDF 文本较长，所以需要 chunking。
- 直接总结容易丢失术语和数据，所以增加 NIL 预处理。
- LLM 输出可能不是合法 JSON，所以增加 JSON 修复机制。
- PPT 文字可能溢出，所以使用卡片式布局、文本压缩和自动换行。
- 论文技术术语多，所以英文术语尽量保留。

## 输出 PPT 结构

生成的 `briefing.pptx` 为 16:9，共三页：

1. 标题页 / 核心愿景页
2. 三大抽象层级页：Hardware × Algorithm × Application
3. 技术演进时间轴页：近期到远期
