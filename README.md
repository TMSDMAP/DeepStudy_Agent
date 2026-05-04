# 深学讲义 Agent

一个面向课堂教学场景的多 Agent 讲义生成项目。它把“教研拆解、教学设计、互动设计、题库生成、总编整合、质量校验”串成一条完整工作流，用于生成更接近真实课堂使用的讲义，而不只是一个问答助手。

## 项目定位

本项目主要解决三个问题：

- 将教学主题自动扩展为结构完整的课堂讲义
- 支持教师上传临时资料与长期个人知识库参与生成
- 支持导出为适合课堂使用的 HTML、PPT、DOCX 等格式

## 核心能力

- 多 Agent 工作流：
  - 教研 Agent
  - 教学设计 Agent
  - 互动设计 Agent
  - 题库 Agent
  - 总编 Agent
  - 质检 Agent
- 教师资料注入：
  - 本次临时资料解析
  - 长期个人知识库检索
- 多格式导出：
  - HTML
  - PPTX
  - DOCX
- 多模型接入：
  - DeepSeek 文本生成
  - DashScope 多模态解析
- 桌面化运行：
  - 基于 Streamlit
  - 支持 PyInstaller 打包为 Windows 可执行程序

## 技术栈

- Python
- Streamlit
- LangGraph
- LangChain
- OpenAI-compatible API
- python-docx
- python-pptx
- PyInstaller

## 仓库结构

```text
.
├─ app.py                   # Streamlit 主应用
├─ agents.py                # 多 Agent 定义、模型接入、质检逻辑
├─ workflow.py              # LangGraph 工作流编排
├─ state.py                 # 工作流共享状态
├─ launcher_streamlit.py    # 打包后的桌面启动入口
├─ fill_ai_report.py        # 作品报告辅助填充脚本
├─ deepstudy_agent.spec     # PyInstaller 打包配置
├─ requirements.txt         # 依赖列表
├─ .env.example             # 环境变量模板
└─ hooks/                   # PyInstaller 自定义 hooks
```

## 运行方式

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置环境变量

复制模板文件：

```bash
copy .env.example .env
```

按需填写以下配置：

```env
# 文本生成：DeepSeek
LLM_API_KEY=your_llm_api_key
LLM_API_BASE=https://api.deepseek.com
LLM_MODEL=deepseek-v4-flash
LLM_THINKING_ENABLED=true
LLM_THINKING_TYPE=enabled
LLM_REASONING_EFFORT=low
LLM_MAX_TOKENS=32768

# 多模态解析：DashScope
VISION_API_KEY=your_vision_api_key
VISION_API_BASE=https://dashscope.aliyuncs.com/compatible-mode/v1
VISION_MODEL=qwen-vl-plus
```

### 3. 启动应用

```bash
streamlit run app.py
```

## 打包

项目已包含：

- `launcher_streamlit.py`
- `deepstudy_agent.spec`
- `hooks/hook-workflow.py`

可用于 PyInstaller 打包 Windows 可执行文件。

## 适用场景

- 中学或大学课堂讲义生成
- 教师备课与课件辅助
- 课程 demo / 比赛作品展示
- 基于私有资料的教学内容再组织

## 当前仓库不包含

为了保证安全和仓库整洁，以下内容默认不随源码上传：

- 真实 `.env`
- API Key
- 本地长期知识库数据
- 打包后的 exe 发布目录
- 本地虚拟环境
- 临时调试文件

## License

本项目当前使用 `MIT License`。
