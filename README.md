# Thesis MCP Server — 毕业论文格式审查 MCP 工具

为 [Nexent](https://github.com/ModelEngine-Group/nexent) 平台提供 Word 文档处理能力，支持论文格式规范审查、内容完整性检查和自动修正。

## 功能

提供 3 个 MCP 工具，供 Nexent 智能体调用：

| 工具 | 功能 | 输入 |
|------|------|------|
| `extract_doc_structure` | 提取 Word 文档完整结构、格式和内容 | 本地路径 或 HTTP URL |
| `convert_doc_format` | .doc ↔ .docx 格式互转（基于 LibreOffice） | 本地路径 或 HTTP URL |
| `apply_corrections` | 应用格式修正，生成新版 .docx 并返回下载链接 | 本地路径 或 HTTP URL + 修正指令 JSON |

## 快速开始

### 1. 环境要求

- Python 3.10+
- LibreOffice（用于 .doc ↔ .docx 转换，仅 Windows 需要单独安装）

### 2. 安装依赖

```bash
# 创建虚拟环境（推荐）
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装
pip install fastmcp python-docx requests boto3
```

### 3. 配置（可选）

```bash
# 复制配置文件模板
cp .env.example .env

# 编辑 .env，按需修改
```

**环境变量说明：**

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `THESIS_MCP_PORT` | `8899` | 服务监听端口 |
| `THESIS_OUTPUT_DIR` | `./output` | 修正文件输出目录 |
| `MINIO_ENDPOINT` | 空（不启用） | MinIO/S3 地址，不配则无下载链接 |
| `MINIO_ACCESS_KEY` | 空 | MinIO Access Key |
| `MINIO_SECRET_KEY` | 空 | MinIO Secret Key |
| `MINIO_BUCKET` | `nexent` | 存储桶名称 |
| `MINIO_EXTERNAL_HOST` | 空 | 替换 localhost 的外部 IP/域名 |

### 4. 启动

```bash
./start_thesis_mcp.sh

# 或者直接运行
python3 thesis_mcp_server.py
```

启动后服务监听 `http://0.0.0.0:8899/mcp`。

### 5. 在 Nexent 中注册

1. 登录 Nexent → Agent 页面 → MCP Config
2. 添加 MCP Server：
   - **Service Name**: `thesis-reviewer`
   - **MCP URL**: `http://<你的IP>:8899/mcp`
3. 点击 Refresh Tools，确认发现 3 个工具

### 6. 创建智能体

在 Nexent 中创建 Agent：

| 字段 | 建议值 |
|------|--------|
| 名称 | `thesis_format_reviewer` |
| 显示名称 | `毕业论文格式规范审查` |
| 模型 | deepseek-v4-pro 或其他大语言模型 |
| 工具 | 勾选 `extract_doc_structure`、`convert_doc_format`、`apply_corrections` |

**系统提示词（Duty Prompt）** 参见下方 [Agent 提示词](#agent-提示词) 章节。

---

## 使用方式

### 对话示例

```
用户: 请审查这篇论文
     模板: /data/论文模板.doc
     论文: [上传文件]

智能体:
  1. convert_doc_format("论文模板.doc") → 论文模板.docx
  2. extract_doc_structure("论文模板.docx") → 规范标准 JSON
  3. extract_doc_structure("学生论文.docx") → 论文结构 JSON
  4. LLM 对比分析 → 生成审查报告

用户: 请修正所有格式问题

智能体:
  1. 生成 corrections_json
  2. 展示修改清单 → 用户确认
  3. apply_corrections(论文, corrections) → 下载链接
```

### 文件输入说明

**支持两种方式：**

1. **本地路径**：服务器上的绝对路径
   ```
   /home/user/thesis.docx
   ```

2. **HTTP URL**：Nexent 上传文件时自动生成
   ```
   http://localhost:5013/api/nb/v1/file/fetch?presigned_url=...
   ```

工具会自动检测并处理，无需区分。

### 同一会话复用模板

同一个对话中，首次加载模板后，后续审查新论文时 Agent 会自动复用模板标准，无需重复加载。

---

## Agent 提示词

将以下提示词配置到 Nexent Agent 的 Duty Prompt：

```markdown
# Role（角色）
你是一位专业的毕业论文审查专家。你的职责是全面审查学生提交的毕业论文，
检查其格式规范性和内容完整性。

# Core Workflow（核心工作流程）

## 第一步：加载规范标准
1. 如果文件是 .doc 格式，先调用 convert_doc_format 转换为 .docx
2. 调用 extract_doc_structure 读取模板文档
3. 从模板中提取审查标准：格式标准（字体、字号、行距、页边距、标题样式等）、
   内容标准（章节要求、主题覆盖、章节顺序）、特殊元素要求

## 第二步：审查论文
1. 调用 extract_doc_structure 读取学生论文
2. 逐项对比模板标准，检查格式合规和内容完整性

## 第三步：生成审查报告
📋 总体评价 | ❌ 格式问题清单 | ⚠️ 内容问题 | ✅ 合规项

## 第四步：修正论文（用户确认后执行）
1. 根据审查报告生成 corrections_json
2. 展示修改清单，获得用户确认
3. 调用 apply_corrections 执行修正
4. 返回下载链接（24小时有效）

# Rules（重要原则）
- 严格以模板为标准，不自行放宽标准
- 每条问题必须指明具体位置（章节 + 段落）
- 修正前必须展示清单并获得确认
- .doc 文件先用 convert_doc_format 转换
- 论文中存在模板外的内容，标记为「超出规范范围」不扣分
- 同一会话中已加载的模板标准自动复用
- 修正完成后，向用户展示 apply_corrections 返回的 download_url

# File Handling
- file_path 参数支持本地路径和 HTTP URL，直接传入即可
```

---

## 项目结构

```
terminal/
├── thesis_mcp_server.py    # MCP Server 主程序
├── start_thesis_mcp.sh     # 启动脚本
├── requirements.txt        # Python 依赖
├── .env.example            # 配置模板
├── output/                 # 修正文件输出目录（自动创建）
└── README.md               # 本文件
```

---

## Make It Portable（让别人也能用）

其他用户部署此 MCP 的步骤：

```bash
# 1. 下载代码
git clone <repo> && cd terminal

# 2. 安装依赖
python3 -m venv venv && source venv/bin/activate
pip install fastmcp python-docx requests boto3

# 3. 安装 LibreOffice（仅 .doc 转换需要）
# Ubuntu/Debian: sudo apt install libreoffice
# macOS:         brew install libreoffice
# Windows:       下载安装 LibreOffice，确保 soffice 在 PATH 中

# 4. 配置（可选）
cp .env.example .env
# 编辑 .env，设置 THESIS_MCP_PORT 等

# 5. 启动
./start_thesis_mcp.sh

# 6. 在 Nexent 中注册
# MCP URL: http://<本机IP>:8899/mcp
```

---

## 常见问题

**Q: 没有 MinIO 能用吗？**
A: 能。不配置 `MINIO_ENDPOINT` 则修正文件仅保存本地，不生成下载链接。Agent 会告知用户文件保存路径。

**Q: 提示 "LibreOffice conversion failed"？**
A: 确保已安装 LibreOffice 且 `soffice` 命令在 PATH 中。仅处理 .doc 文件需要。

**Q: 如何修改监听端口？**
A: 设置环境变量 `THESIS_MCP_PORT=8080` 或在 `.env` 中配置。

**Q: 发布 Agent 后别人能用吗？**
A: 能。其他人只需在他们的 Nexent 中注册你的 MCP URL 即可。Agent 配置（提示词+工具引用）会随发布分发，但 MCP Server 需要单独部署。
