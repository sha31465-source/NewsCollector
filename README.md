# 📰 NewsCollector V36 - 智能资讯采集与分析系统

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Active-success)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

**一个现代化的智能资讯采集、分类、分析与报告生成工具**

[功能特性](#功能特性) • [快速开始](#快速开始) • [使用指南](#使用指南) • [配置说明](#配置说明) • [常见问题](#常见问题)

</div>

---

## 📋 项目简介

**NewsCollector V36** 是一个功能强大的资讯采集与分析系统，能够从 RSS 源自动采集资讯、智能分类、提取关键信息，并生成结构化的 Word/TXT 报告，CSV目录。系统采用现代化的左右分栏 UI 设计，界面简洁直观，操作便捷高效。需要什么信息取决于加入了什么rss源，以及添加了什么关键词，目前配置11个rss源。

### 🌟 核心亮点

- **🎨 现代化静态分栏 UI**：左侧导航栏 + 右侧内容区，视觉清晰，操作流畅
- **🤖 智能分类引擎**：基于关键字匹配，自动将资讯分配到不同类别
- **📊 自动报告生成**：一键导出 Word 或 TXT 格式的结构化报告
- **🧹 智能清理功能**：一键清理不匹配关键字的资讯，保持数据库整洁
- **💎 源代码提取**：自动从资讯中提取代码片段，方便技术分析
- **⚙️ 灵活的配置管理**：支持自定义 RSS 源、分类关键字、API 配置

---

## ✨ 功能特性

### 🎯 核心功能

| 功能 | 说明 |
|------|------|
| **🚀 全流程运行** | 一键完成采集 → 处理 → 报告生成 |
| **📡 RSS 采集** | 从多个 RSS 源并行采集最新资讯 |
| **🧠 智能处理** | 自动去重、分类、提取关键信息 |
| **📄 报告生成** | 支持 Word 和 TXT 两种格式导出 |
| **🗂️ 分类管理** | 自定义分类关键字，精准匹配 |
| **🔗 RSS 管理** | 灵活添加、删除、管理 RSS 订阅源 |
| **🧹 智能清理** | **一键清理不相关资讯**（V36 新增）|
| **💻 代码提取** | 自动提取资讯中的源代码片段 |
| **⚙️ 系统设置** | 个性化配置输出格式、API 等参数 |

### 🆕 V36 新增特性

- ✨ **全新 UI 设计**：180px 固定宽度左侧导航栏，深灰色背景
- ✨ **完整按钮文字**：无缩写，清晰易懂（如 `[全] 全流程运行`）
- ✨ **设置按钮沉底**：通过弹性占位符固定在左下角
- ✨ **纯静态布局**：移除所有动画效果，极大提升稳定性和性能
- ✨ **智能清理功能**：一键标记不匹配关键字的资讯为"已忽略"
- ✨ **代码重构**：简化架构，更加易维护

---

## 🚀 快速开始

### 环境要求

- **操作系统**: Windows 7/10/11
- **Python**: 3.8 或更高版本
- **内存**: 至少 2GB RAM
- **硬盘**: 至少 500MB 可用空间
- **网络**: 需要互联网连接（采集 RSS 资讯）

### 安装步骤

#### 方法一：使用已编译的 EXE 文件（推荐）

1. 下载 [`NewsCollector_V36.exe`](dist/NewsCollector_V36.exe)
2. 双击运行，首次使用自动初始化数据库

#### 方法二：从源码运行

```bash
# 1. 克隆仓库
git clone https://github.com/yourusername/NewsCollector.git
cd NewsCollector

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python main_gui_standalone.py
```

### 依赖安装

请使用以下命令安装所有必需的依赖包：

```bash
pip install -r requirements.txt
```

或在虚拟环境中安装（推荐）：

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境（Windows）
venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

---

## 📖 使用指南

### 快速上手

1. **启动程序**
   ```
   双击 NewsCollector_V36.exe
   ```

2. **配置 RSS 源**
   - 点击左侧 `[源] RSS源` 按钮
   - 添加您需要的 RSS 订阅地址
   - 推荐源：
     ```
     https://www.pythontab.com/html/news/rss.xml
     https://openai.com/news/rss.xml
     https://www.python.org/news/rss.xml
     ```

3. **配置分类关键字**
   - 点击左侧 `[关] 关键字` 按钮
   - 为每个分类添加关键字
   - 示例：
     ```
     AI技术：人工智能, AI, 机器学习, 深度学习, ChatGPT
     Python：Python, Django, Flask, NumPy, Pandas
     前端：JavaScript, Vue, React, Angular, TypeScript
     ```

4. **开始采集**
   - **方式一（推荐）**：点击 `[全] 全流程运行` 一键完成所有步骤
   - **方式二**：点击 `[采] 仅采集` 只采集资讯

5. **查看结果**
   - 日志区域：实时显示运行状态和进度
   - 状态栏：显示统计数据（总数、待处理、已处理）
   - 输出文件夹：查看生成的报告（`输出/docs/`）

### 左侧导航栏按钮说明

| 按钮 | 功能 |
|------|------|
| **[全] 全流程运行** | 自动执行：采集 → 处理 → 报告生成 |
| **[采] 仅采集** | 从 RSS 源采集最新资讯到数据库 |
| **[处] 仅处理** | 处理数据库中的待处理资讯（分类、提取） |
| **[报] 生成报告** | 将已处理资讯导出为 Word/TXT 报告 |
| **[源] RSS源** | 管理 RSS 订阅源（添加、删除、查看） |
| **[关] 关键字** | 管理分类关键字（添加、删除、编辑） |
| **[提] 提取源代码** | 从资讯中提取代码片段到 `输出/scripts/` |
| **[清] 清理不相关资讯** | **将不匹配关键字的资讯标记为"已忽略"** |
| **[设] 设置** | 系统设置（API配置、输出格式等） |

---

## 🗄️ 数据库初始化

系统使用 SQLite 数据库存储数据，首次运行会自动创建 `news_v3.db` 文件。

### 核心表结构

#### news_items 表（资讯主表）

```sql
CREATE TABLE IF NOT EXISTS news_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,              -- 资讯标题
    link TEXT NOT NULL UNIQUE,        -- 资讯链接
    description TEXT,                 -- 资讯描述
    content TEXT,                     -- 完整内容
    pub_date TEXT,                    -- 发布日期
    source TEXT,                      -- RSS源名称
    category TEXT,                    -- 分类（AI技术、Python、前端等）
    is_processed INTEGER DEFAULT 0,   -- 是否已处理：0-待处理，1-已处理
    status TEXT DEFAULT 'pending',    -- 状态：pending/processed/ignored
    collected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 采集时间
    processed_at TIMESTAMP,           -- 处理时间
    summary TEXT,                     -- 摘要（如有）
    code_blocks TEXT                  -- 提取的代码块（JSON格式）
);

-- 索引
CREATE INDEX IF NOT EXISTS idx_link ON news_items(link);
CREATE INDEX IF NOT EXISTS idx_status ON news_items(status);
CREATE INDEX IF NOT EXISTS idx_category ON news_items(category);
CREATE INDEX IF NOT EXISTS idx_is_processed ON news_items(is_processed);
```

#### rss_sources 表（RSS源配置）

```sql
CREATE TABLE IF NOT EXISTS rss_sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,               -- 源名称
    url TEXT NOT NULL UNIQUE,         -- RSS链接
    category TEXT,                    -- 默认分类
    enabled INTEGER DEFAULT 1,        -- 是否启用：1-启用，0-禁用
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

#### keywords 表（分类关键字）

```sql
CREATE TABLE IF NOT EXISTS keywords (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT NOT NULL,           -- 分类名称
    keyword TEXT NOT NULL,            -- 关键字
    enabled INTEGER DEFAULT 1         -- 是否启用
);
```

### 数据库文件位置

- **开发环境**: 与源代码同级目录
- **EXE运行**: 与 EXE 文件同级目录

---

## ⚙️ 配置说明

### prompts_config.yaml 配置文件

首次运行会在程序目录创建 `prompts_config.yaml` 配置文件：

```yaml
# API配置（可选，用于智能摘要和翻译）
api:
  api_url: "https://open.bigmodel.cn/api/paas/v4/"
  api_key: ""  # 请在此处填写您的API密钥
  model: "glm-4-flash"

# 分类配置
categories:
  - name: "AI技术"
    keywords: ["人工智能", "AI", "机器学习", "深度学习"]
  - name: "Python"
    keywords: ["Python", "Django", "Flask", "NumPy"]
  - name: "前端"
    keywords: ["JavaScript", "Vue", "React", "Angular"]

# 输出配置
output:
  format: "txt"  # txt 或 word
  save_raw: true
  extract_code: true
```

### 输出目录结构

程序运行后会在当前目录创建以下文件夹：

```
NewsCollector_V36/
├── news_v3.db              # 数据库文件
├── prompts_config.yaml     # 配置文件
├── NewsCollector_V36.exe   # 主程序
└── 输出/                   # 自动生成
    ├── docs/              # Word/TXT报告
    ├── logs/              # 运行日志
    ├── scripts/           # 提取的源代码
    ├── images/            # 下载的图片
    └── other/             # 其他附件
```

---

## 🧹 智能清理功能说明

### 功能介绍

**`[清] 清理不相关资讯`** 是 V36 版本新增的核心功能，用于清理数据库中不匹配任何分类关键字的资讯。

### 工作原理

1. **扫描待处理资讯**：查询所有 `status='pending'` 且 `is_processed=0` 的资讯
2. **关键字匹配**：遍历所有已启用的分类关键字
3. **标记为已忽略**：对于不匹配任何关键字的资讯，更新为 `status='ignored', is_processed=1`
4. **保护相关资讯**：已匹配关键字的资讯不会被清理

### 使用场景

- ✅ 定期清理数据库，移除不相关资讯
- ✅ 专注于特定主题的资讯
- ✅ 提高报告质量和相关性

### 注意事项

- ⚠️ 清理操作**不可逆**，请谨慎使用
- ⚠️ 建议先配置好关键字，再执行清理
- ⚠️ 已处理的资讯不会被清理

---

## 🔧 常见问题

### Q1: 程序无法启动？

**A**: 请确保：
- Windows 7/10/11 系统
- 已安装 Python 3.8+（源码运行）
- 有网络连接（采集 RSS 需要）

### Q2: 采集失败或没有新资讯？

**A**: 请检查：
- RSS 源地址是否正确
- 网络连接是否正常
- 该 RSS 源是否有新内容发布

### Q3: 如何备份我的数据？

**A**: 定期备份以下文件：
- `news_v3.db`（数据库，包含所有采集的资讯）
- `prompts_config.yaml`（配置文件）

### Q4: 找不到生成的报告？

**A**: 报告保存在程序目录的 `输出/docs/` 文件夹中

### Q5: 清理不相关资讯后能恢复吗？

**A**: 不能。清理操作会将资讯标记为 `status='ignored'`，建议清理前备份数据库。

### Q6: 如何添加新的分类？

**A**:
1. 点击 `[关] 关键字` 按钮
2. 输入新的分类名称（如 "区块链"）
3. 添加对应的关键字（如 "区块链", "Bitcoin", "以太坊"）
4. 保存后重新采集即可

---

## 📦 依赖列表

详见 [`requirements.txt`](requirements.txt)：

```
customtkinter==5.2.1    # 现代化GUI框架
pyyaml==6.0.1           # YAML配置文件解析
python-docx==1.1.0      # Word文档生成
feedparser==6.0.10      # RSS解析
lxml==5.1.0             # XML/HTML处理
urllib3==2.0.0          # HTTP请求
```

---

## 📄 开源许可

本项目采用 **MIT License** 开源许可证。

详见 [LICENSE](LICENSE) 文件。

---

## 🤝 贡献指南

欢迎贡献代码、报告 Bug 或提出新功能建议！

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

---

## 📞 联系方式

- **Issues**: [GitHub Issues](https://github.com/yourusername/NewsCollector/issues)
- **Email**: your.email@example.com

---

## 🙏 致谢

感谢以下开源项目：

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - 现代化 GUI 框架
- [Feedparser](https://github.com/kurtmckee/feedparser) - RSS 解析库
- [python-docx](https://github.com/python-openxml/python-docx) - Word 文档生成

---

## 📈 更新日志

### [3.6.0] - 2026-05-21

#### 新增
- ✨ 全新左右分栏 UI 设计
- ✨ 按钮文字完整显示，无缩写
- ✨ 设置按钮沉底显示
- ✨ 智能清理不相关资讯功能

#### 优化
- 🚀 移除所有动画效果，极大提升稳定性
- 🚀 代码重构，架构更简洁
- 🚀 性能优化，启动速度提升 50%

#### 修复
- 🐛 修复按钮卡顿问题
- 🐛 修复内存泄漏问题

---

<div align="center">

**Made with ❤️ by NewsCollector Team**

⭐ 如果这个项目对您有帮助，请给我们一个 Star！

</div>
