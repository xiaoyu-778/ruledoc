# RuleDoc

[![PyPI version](https://badge.fury.io/py/ruledoc.svg)](https://badge.fury.io/py/ruledoc)
[![Python versions](https://img.shields.io/pypi/pyversions/ruledoc.svg)](https://pypi.org/project/ruledoc/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/xiaoyu-778/ruledoc/workflows/CI/badge.svg)](https://github.com/xiaoyu-778/ruledoc/actions)
[![codecov](https://codecov.io/gh/xiaoyu-778/ruledoc/branch/main/graph/badge.svg)](https://codecov.io/gh/xiaoyu-778/ruledoc)

> 规则驱动的论文格式化工具 - 让学术文档排版变得简单高效

RuleDoc 是一个基于规则引擎驱动的文档格式化工具，支持多种高校论文格式规范。通过将格式规则与核心处理逻辑完全解耦，用户可以轻松扩展支持新的格式规范。

## ✨ 核心特性

- **规则驱动**: 格式规范与处理逻辑完全分离，易于扩展
- **多格式支持**: 支持 Markdown、DOCX 等输入格式（通过 Pandoc 转换）
- **智能处理**: 自动识别标题层级、图表编号、列表格式
- **插件架构**: 支持自定义后处理器，满足个性化需求
- **类型安全**: 完整的类型注解支持

## 📦 安装

### 从 PyPI 安装

```bash
pip install ruledoc
```

### 从源码安装

```bash
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc
pip install -e .
```

### 开发安装

```bash
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc
pip install -e ".[dev]"
```

## 🚀 快速开始

### 基本使用

```bash
# 使用默认规则格式化 Markdown 文件
ruledoc input.md

# 指定输出路径
ruledoc input.md -o output.docx

# 使用特定规则
ruledoc input.md --rule yzu_thesis

# 格式化 DOCX 文件（重新应用格式规则）
ruledoc input.docx --rule yzu_thesis

# 列出所有可用规则
ruledoc --list-rules
```

### Python API 使用

```python
from ruledoc import Formatter, load_rule

# 加载规则
rule = load_rule("yzu_thesis")

# 创建格式化器
formatter = Formatter("input.md", rule, "output.docx")

# 执行格式化
formatter.process()
```

## 📋 支持的规则

| 规则名称 | 说明 | 状态 |
|---------|------|------|
| `yzu_thesis` | 扬州大学学位论文格式 | ✅ 稳定 |
| `yzu_design` | 扬州大学毕业设计格式 | ✅ 稳定 |

> 欢迎贡献更多学校的格式规则！

## 🏗️ 架构设计

```
ruledoc/
├── processors/          # 文档处理器
│   ├── base.py         # 处理器基类
│   ├── heading_processor.py   # 标题处理器
│   ├── caption_processor.py   # 图表标题处理器
│   ├── list_processor.py      # 列表处理器
│   └── style_processor.py     # 样式处理器
├── rules/              # 格式规则
│   ├── base.py         # 规则基类
│   └── yzu/           # 扬州大学规则集
│       ├── common.py   # 通用规则
│       ├── yzu_thesis.py    # 学位论文规则
│       └── yzu_design.py    # 毕业设计规则
├── context.py          # 处理上下文
├── formatter.py        # 主格式化器
└── cli.py             # 命令行接口
```

## 🛠️ 开发指南

### 环境设置

```bash
# 克隆仓库
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装开发依赖
pip install -e ".[dev]"
```

### 运行测试

```bash
# 运行所有测试
pytest

# 运行测试并生成覆盖率报告
pytest --cov=ruledoc --cov-report=html

# 运行特定测试
pytest tests/test_cli.py -v
```

### 代码规范

```bash
# 代码格式化
black ruledoc tests
isort ruledoc tests

# 类型检查
mypy ruledoc

# 代码检查
ruff check ruledoc tests
```

## 🤝 贡献指南

我们欢迎所有形式的贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解详细信息。

### 贡献步骤

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 创建 Pull Request

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE) 开源。

## 🙏 致谢

- [python-docx](https://python-docx.readthedocs.io/) - Word 文档处理
- [pypandoc](https://github.com/NicklasTegner/pypandoc) - Pandoc 封装
- 所有贡献者！

## 📚 文档导航

| 文档 | 说明 |
|------|------|
| [README.md](README.md) | 项目介绍和快速开始 |
| [CONTRIBUTING.md](CONTRIBUTING.md) | 贡献指南和开发规范 |
| [docs/rule_development.md](docs/rule_development.md) | 规则开发详细指南 |
| [CHANGELOG.md](CHANGELOG.md) | 版本更新日志 |
| [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) | 行为准则 |

## 📞 联系我们

- 项目主页: https://github.com/xiaoyu-778/ruledoc
- 问题反馈: https://github.com/xiaoyu-778/ruledoc/issues
- 邮箱: 3389357760@qq.com

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/xiaoyu-778">xiaoyu-778</a>
</p>
