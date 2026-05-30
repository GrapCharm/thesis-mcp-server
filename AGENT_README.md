# 毕业论文格式规范审查

基于 Nexent 平台的智能论文审查助手，支持 Word 文档格式规范性检查和内容完整性审查。

## 功能特性

- **格式审查**：自动检测字体、字号、行距、页边距、标题样式、首行缩进等格式问题
- **内容审查**：对比规范模板，检查章节完整性、主题覆盖度、结构层级
- **自动修正**：根据审查结果一键修正格式问题，生成规范版 Word 文档
- **多格式支持**：兼容 `.doc` 和 `.docx` 格式，自动转换

## 使用效果

| 审查前 | 审查后 |
|--------|--------|
| 上传论文 + 规范模板 | 逐条列出格式问题和内容缺失 |
| 一键修正 | 下载符合规范的修正版论文 |

## 前置要求

部署本智能体前，需要先部署 MCP Server：

> 🔗 **MCP Server 仓库**：https://github.com/GrapCharm/thesis-mcp-server

### MCP Server 部署步骤

```bash
# 1. 克隆仓库
git clone https://github.com/GrapCharm/thesis-mcp-server.git
cd thesis-mcp-server

# 2. 安装依赖
pip install fastmcp python-docx requests

# 3. 安装 LibreOffice（仅 .doc 转换需要）
# Ubuntu/Debian: sudo apt install libreoffice
# macOS:         brew install libreoffice
# Windows:       从 https://www.libreoffice.org 下载安装

# 4. 启动服务
python3 thesis_mcp_server.py
```

服务启动后监听 `http://0.0.0.0:8899/mcp`。

### 在 Nexent 中连接 MCP

1. 登录 Nexent，进入 **Agent 页面 → MCP Config**
2. 添加 MCP Server：
   - **Service Name**：`thesis-reviewer`
   - **MCP URL**：`http://<你的服务器IP>:8899/mcp`
3. 点击 **Refresh Tools**，确认发现 3 个工具：
   - `extract_doc_structure` — 提取文档结构
   - `convert_doc_format` — 格式互转
   - `apply_corrections` — 应用修正

### 导入智能体

1. 在 Nexent 中导入本智能体
2. 将智能体的模型切换为你平台上的大语言模型
3. 确认工具已关联上述 3 个 MCP 工具
4. 开始使用

## 使用方式

### 对话示例

```
用户: 请审查这篇论文
     模板路径: /data/论文模板.doc
     论文: [上传文件]

智能体:
  → 加载模板规范标准
  → 提取论文结构和格式
  → 对比审查
  → 输出审查报告（格式问题 + 内容问题 + 合规项）

用户: 请修正所有格式问题

智能体:
  → 展示修改清单
  → 用户确认
  → 生成修正版文档
```

### 文件说明

- **模板文件**：论文格式规范文档（.doc/.docx），供智能体提取审查标准
- **论文文件**：待审查的学生论文（.doc/.docx），支持上传或本地路径
- **修正输出**：修正后的文件保存在 MCP Server 的 `output/` 目录

## 系统提示词

智能体的审查逻辑基于以下提示词，可在 Nexent Agent 的 Duty Prompt 中配置：

```markdown
# Role（角色）
你是一位专业的毕业论文审查专家。审查论文格式规范性和内容完整性。

# Core Workflow
1. 加载规范模板 → 提取格式标准和内容要求
2. 读取学生论文 → 提取结构和格式元数据
3. 逐项对比 → 生成审查报告（格式问题 + 内容问题 + 合规项）
4. 用户确认后 → 应用格式修正

# Rules
- 严格以模板为标准，不自行放宽
- 每条问题指明具体位置
- 修正前展示清单并获得确认
- .doc 文件先用 convert_doc_format 转换
```

## 许可证

MIT License
