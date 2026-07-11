# AI+硬件协同设计愿景自动化简报生成器

本工具读取 PDF 或 TXT 论文，经过文档加载、chunking、NIL 预处理与合并、最终 JSON 提取和 JSON repair，生成三页 `briefing.pptx`。LLM 调用使用 OpenAI Python SDK，并支持通用 OpenAI-compatible API。

## 安装

建议使用 Python 3.10+：

```powershell
pip install -r requirements.txt
```

## API 配置与运行

配置优先级如下：

- API Key：`--api-key`、`API_KEY`、`OPENAI_API_KEY`、`DEEPSEEK_API_KEY`、运行时隐藏输入。
- Base URL：`--base-url`、`API_BASE_URL`；均未提供时使用 OpenAI 官方接口。
- Model：`--model`、`API_MODEL`；均未提供时使用 `gpt-4o-mini`。

### 推荐方式一：运行时隐藏输入 API Key

```powershell
python generate_briefing.py --input paper.pdf --output briefing.pptx --base-url https://api.deepseek.com --model deepseek-chat
```

程序会提示 `No API key found. Please enter your API key:`，输入时不会显示明文，也不会把 Key 保存到文件。

### 推荐方式二：使用环境变量

Windows PowerShell：

```powershell
$env:API_KEY="your_api_key_here"
$env:API_BASE_URL="https://api.deepseek.com"
$env:API_MODEL="deepseek-chat"

python generate_briefing.py --input paper.pdf --output briefing.pptx
```

### OpenAI 官方接口示例

```powershell
$env:OPENAI_API_KEY="your_openai_api_key_here"
$env:API_MODEL="gpt-4o-mini"

python generate_briefing.py --input paper.pdf --output briefing.pptx
```

不设置 Base URL 时，程序使用 OpenAI 官方接口。也可以用 `--api-key` 传入 Key，但命令行参数可能被 shell 历史或进程列表记录，因此更推荐隐藏输入或环境变量。

其他运行示例：

默认 chunk 参数为 `--max-chars 7000` 和 `--overlap 300`。用户仍可通过命令行指定其他值：

```powershell
python generate_briefing.py --input paper.txt --output briefing.pptx
python generate_briefing.py --input paper.pdf --output briefing.pptx --max-chars 5000 --overlap 300
```

## 安全提醒

- 不要把 API Key 写进代码。
- 不要把 API Key 上传到 GitHub。
- 不要在 README 中填写真实 API Key。
- 不要截图泄露 API Key。
- 如果怀疑 API Key 泄露，应立刻在平台删除旧 key 并重新创建。

## 处理流程

程序保持以下流程：文档加载 → chunking → NIL 预处理 → merge NIL → 最终 JSON 提取 → JSON repair（需要时）→ PPT 自动生成。

生成的 PPT 为 16:9 三页结构：

1. 标题页 / 核心愿景页
2. Hardware × Algorithm × Application 三大抽象层级页
3. 近期到远期的技术演进时间轴页
