# RuleDoc

[![PyPI version](https://badge.fury.io/py/ruledoc.svg)](https://badge.fury.io/py/ruledoc)
[![Python versions](https://img.shields.io/pypi/pyversions/ruledoc.svg)](https://pypi.org/project/ruledoc/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/xiaoyu-778/ruledoc/workflows/CI/badge.svg)](https://github.com/xiaoyu-778/ruledoc/actions)
[![codecov](https://codecov.io/gh/xiaoyu-778/ruledoc/branch/main/graph/badge.svg)](https://codecov.io/gh/xiaoyu-778/ruledoc)

> 规则驱动的论文格式化工具 - 让学术文档排版变得简单高效

RuleDoc 是一个基于规则引擎驱动的文档格式化工具，专门为中国高校学位论文格式规范设计。**使用 RuleDoc 必须指定格式规则**，通过将格式规则与核心处理逻辑完全解耦，用户可以轻松扩展支持新的格式规范，无需修改核心代码。

## ✨ 核心特性

- **规则驱动架构**: 格式规范与处理逻辑完全分离，**必须选择规则才能使用**，新增学校支持只需添加规则文件
- **多格式输入**: 支持 Markdown、DOCX 等输线表等学术规范
- **插件扩展**: 支持自定义后处理器，满足个性化格式需求
- **类型安全**: 完整的类型注解支持，IDE 友好

## 📋 目录

- [快速体验](#快速体验)
- [安装](#安装)
- [使用指南](#使用指南)
- [支持的规则](#支持的规则)
- [项目架构](#项目架构)
- [贡献指南](#贡献指南)
- [文档导航](#文档导航)

## 🚀 快速体验

使用项目自带的 `test_script.md` 快速体验 RuleDoc：

```bash
# 1. 克隆仓库
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc

# 2. 安装依赖
pip install -e .

# 3. 运行测试（使用 test_script.md）
python -m ruledoc test_script.md --rule yzu_thesis
```

运行后会：

1. 加载 `yzu_thesis` 规则（扬州大学学位论文格式）
2. 生成 `test_script_formatted.docx`

**test_script.md 包含以下内容：**

- 中英文摘要
- 六级标题层级测试
- 正文段落与列表
- 图表题注（图 4-1、表 4-1）
- 数学公式
- 代码块
- 参考文献（GB/T 7714 格式）

## 📦 安装

### 系统要求

- Python 3.8 或更高版本
- Windows、macOS 或 Linux

### 从 PyPI 安装（推荐）

```bash
pip install ruledoc
```

### 从源码安装

```bash
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc
pip install -e .
```

### 验证安装

```bash
ruledoc --version
ruledoc --help
```

## 📖 使用指南

### ⚠️ 重要：必须使用规则

**RuleDoc 必须指定格式规则才能工作**，没有默认规则。您需要：

1. 查看可用规则：`ruledoc --list-rules`
2. 使用 `--rule` 参数指定规则

### 基本用法

```bash
# 查看可用规则
ruledoc --list-rules

# 使用指定规则格式化
ruledoc thesis.md --rule yzu_thesis

# 指定输出路径
ruledoc thesis.md --rule yzu_thesis -o output.docx
```

### 完整参数说明

| 参数           | 简写 | 说明                    | 示例                |
| -------------- | ---- | ----------------------- | ------------------- |
| `输入文件`     | -    | 要格式化的文件路径      | `thesis.md`         |
| `--rule`       | `-r` | **必填** 使用的格式规则 | `--rule yzu_thesis` |
| `--output`     | `-o` | 输出文件路径            | `-o output.docx`    |
| `--list-rules` | `-l` | 列出所有可用规则        | `--list-rules`      |
| `--version`    | `-v` | 显示版本                | `--version`         |
| `--help`       | `-h` | 显示帮助                | `--help`            |

### 输入文件示例

创建一个 Markdown 文件 `thesis.md`：

```markdown
# 论文标题

## 摘要

本文研究了...

**关键词：** 关键词1；关键词2

## Abstract

This paper studies...

**Keywords:** keyword1; keyword2

## 第一章 绪论

### 1.1 研究背景

随着...

### 1.2 研究意义

本研究...

## 参考文献

[1] 作者. 文章标题[J]. 期刊名, 2024, 1(1): 1-10.
```

### Python API 使用

```python
from ruledoc import Formatter, load_rule

# 加载规则（必须）
rule = load_rule("yzu_thesis")

# 创建格式化器
formatter = Formatter(
    input_path="input.md",
    rule=rule,
    output_path="output.docx"
)

# 执行格式化
formatter.process()
```

## 📋 支持的规则

| 规则名称     | 说明                 | 适用对象       | 状态    |
| ------------ | -------------------- | -------------- | ------- |
| `yzu_thesis` | 扬州大学学位论文格式 | 本科生、研究生 | ✅ 稳定 |
| `yzu_design` | 扬州大学毕业设计格式 | 本科生毕业设计 | ✅ 稳定 |

> 💡 **想要添加新规则？** 查看 [规则开发指南](docs/rule_development.md) 了解如何为其他学校贡献格式规则。

## 🏗️ 项目架构

```
ruledoc/
├── ruledoc/              # 主包
│   ├── __init__.py       # 包入口
│   ├── cli.py            # 命令行接口
│   ├── formatter.py      # 主格式化器
│   ├── context.py        # 处理上下文
│   ├── config.py         # 配置管理
│   ├── exceptions.py     # 自定义异常
│   ├── pandoc_converter.py  # Pandoc 转换器
│   ├── rules/            # 格式规则
│   │   ├── __init__.py
│   │   ├── base.py       # 规则基类
│   │   └── yzu/          # 扬州大学规则集
│   │       ├── __init__.py
│   │       ├── common.py
│   │       ├── yzu_thesis.py
│   │       └── yzu_design.py
│   └── processors/       # 文档处理器
│       ├── __init__.py
│       ├── base.py       # 处理器基类
│       ├── heading_processor.py   # 标题处理器
│       ├── caption_processor.py   # 图表标题处理器
│       ├── list_processor.py      # 列表处理器
│       └── style_processor.py     # 样式处理器
├── test_script.md        # 测试示例文件
├── tests/                # 测试目录
├── docs/                 # 文档目录
│   ├── rule_development.md   # 规则开发指南
│   └── user_guide.md         # 用户使用指南
├── README.md             # 项目说明
├── CONTRIBUTING.md       # 贡献指南
├── CHANGELOG.md          # 更新日志
├── LICENSE               # 许可证
└── pyproject.toml        # 项目配置
```

### 核心组件说明

#### 1. 规则系统 (`rules/`)

规则系统采用插件化设计，**必须使用规则**才能格式化：

- **基类**: `FormatRule` 定义规则接口
- **注册机制**: `@register_rule` 装饰器自动注册规则
- **加载方式**: `load_rule(name)` 动态加载规则实例

#### 2. 处理器系统 (`processors/`)

处理器负责具体的文档格式化操作：

- **标题处理器**: 处理章节标题格式
- **图表处理器**: 处理图、表编号和标题
- **列表处理器**: 处理有序/无序列表
- **样式处理器**: 应用字体、段落样式

#### 3. 格式化流程

```
输入文件 → 规则加载（必须）→ Pandoc转换 → 处理器链 → 输出文件
                ↓
            各处理器按顺序应用格式规则
                ↓
            生成符合规范的 Word 文档
```

## 🛠️ 开发指南

### 环境设置

```bash
# 克隆仓库
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境（Windows）
venv\Scripts\activate

# 安装开发依赖
pip install -e ".[dev]"

# 验证安装
ruledoc --version
```

### 使用 test_script.md 测试

```bash
# 测试扬州大学学位论文格式
ruledoc test_script.md --rule yzu_thesis

# 测试扬州大学毕业设计格式
ruledoc test_script.md --rule yzu_design
```

### 运行测试

```bash
# 运行所有测试
pytest

# 运行测试并生成覆盖率报告
pytest --cov=ruledoc --cov-report=html
```

## 🤝 贡献指南

我们欢迎所有形式的贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解详细信息。

### 快速贡献步骤

1. **Fork** 本仓库
2. **创建分支**: `git checkout -b feature/amazing-feature`
3. **提交更改**: `git commit -m 'feat: add amazing feature'`
4. **推送分支**: `git push origin feature/amazing-feature`
5. **创建 Pull Request**

## 📚 文档导航

| 文档                                                 | 说明               | 目标读者   |
| ---------------------------------------------------- | ------------------ | ---------- |
| [README.md](README.md)                               | 项目介绍、快速开始 | 所有用户   |
| [docs/user_guide.md](docs/user_guide.md)             | 详细使用指南       | 终端用户   |
| [CONTRIBUTING.md](CONTRIBUTING.md)                   | 贡献指南、开发规范 | 贡献者     |
| [docs/rule_development.md](docs/rule_development.md) | 规则开发详细指南   | 规则开发者 |
| [CHANGELOG.md](CHANGELOG.md)                         | 版本更新日志       | 所有用户   |
| [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)             | 行为准则           | 所有参与者 |

## 🙏 致谢

- [python-docx](https://python-docx.readthedocs.io/) - Word 文档处理库
- [pypandoc](https://github.com/NicklasTegner/pypandoc) - Pandoc Python 封装
- [pandoc](https://pandoc.org/) - 通用文档转换工具
- 所有贡献者！

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE) 开源。

## 📞 联系我们

- **项目主页**: <https://github.com/xiaoyu-778/ruledoc>
- **问题反馈**: <https://github.com/xiaoyu-778/ruledoc/issues>
- **功能建议**: <https://github.com/xiaoyu-778/ruledoc/discussions>
- **邮箱**: <3389357760@qq.com>

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/xiaoyu-778">xiaoyu-778</a> and <a href="https://github.com/xiaoyu-778/ruledoc/graphs/contributors">contributors</a>
</p>
