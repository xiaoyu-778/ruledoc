# 贡献指南

感谢您对 RuleDoc 项目的关注！我们欢迎所有形式的贡献，包括但不限于：

- 报告 Bug
- 提出新功能建议
- 改进文档
- 提交代码修复
- 添加新的格式规则

## 目录

- [行为准则](#行为准则)
- [如何贡献](#如何贡献)
  - [报告 Bug](#报告-bug)
  - [提出功能建议](#提出功能建议)
  - [提交代码](#提交代码)
- [开发环境设置](#开发环境设置)
- [代码规范](#代码规范)
- [提交规范](#提交规范)
- [审查流程](#审查流程)

## 行为准则

本项目遵循 [Contributor Covenant](https://www.contributor-covenant.org/) 行为准则。参与本项目即表示您同意遵守此准则。

## 如何贡献

### 报告 Bug

在提交 Bug 报告之前，请先：

1. 搜索现有 issues，确认该问题未被报告过
2. 使用最新版本测试，确认问题仍然存在

提交 Bug 报告时，请使用 [Bug 报告模板](.github/ISSUE_TEMPLATE/bug_report.md)，并包含以下信息：

- 问题描述
- 复现步骤
- 期望行为
- 实际行为
- 环境信息（Python 版本、操作系统等）
- 相关代码或日志

### 提出功能建议

我们欢迎新功能建议！请使用 [功能建议模板](.github/ISSUE_TEMPLATE/feature_request.md)，并描述：

- 功能解决的问题
- 期望的解决方案
- 替代方案（如有）
- 其他上下文信息

### 提交代码

#### 1. Fork 仓库

点击 GitHub 页面右上角的 "Fork" 按钮。

#### 2. 克隆您的 Fork

```bash
git clone https://github.com/YOUR_USERNAME/ruledoc.git
cd ruledoc
```

#### 3. 添加上游仓库

```bash
git remote add upstream https://github.com/xiaoyu-778/ruledoc.git
```

#### 4. 创建分支

```bash
git checkout -b feature/your-feature-name
# 或
git checkout -b fix/your-bug-fix
```

#### 5. 进行更改

- 编写清晰的代码
- 添加或更新测试
- 更新相关文档

#### 6. 提交更改

```bash
git add .
git commit -m "feat: add your feature description"
```

#### 7. 推送到您的 Fork

```bash
git push origin feature/your-feature-name
```

#### 8. 创建 Pull Request

在 GitHub 上创建 Pull Request，并填写 PR 模板。

## 开发环境设置

### 前提条件

- Python 3.8 或更高版本
- Git
- Pandoc（可选，用于测试文档转换）

### 设置步骤

```bash
# 克隆仓库
git clone https://github.com/xiaoyu-778/ruledoc.git
cd ruledoc

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Linux/macOS:
source venv/bin/activate
# Windows:
venv\Scripts\activate

# 安装开发依赖
pip install -e ".[dev]"

# 验证安装
ruledoc --version
```

## 代码规范

### Python 代码风格

我们使用以下工具保持代码一致性：

- **Black**: 代码格式化
- **isort**: 导入排序
- **ruff**: 代码检查
- **mypy**: 类型检查

### 运行代码检查

```bash
# 格式化代码
black ruledoc tests
isort ruledoc tests

# 类型检查
mypy ruledoc

# 代码检查
ruff check ruledoc tests

# 运行所有检查
black --check ruledoc tests && isort --check-only ruledoc tests && ruff check ruledoc tests && mypy ruledoc
```

### 测试

```bash
# 运行所有测试
pytest

# 运行测试并生成覆盖率报告
pytest --cov=ruledoc --cov-report=html

# 运行特定测试文件
pytest tests/test_cli.py -v

# 运行特定测试函数
pytest tests/test_cli.py::test_main_function -v
```

### 代码规范要点

1. **类型注解**: 所有函数参数和返回值都应添加类型注解
2. **文档字符串**: 公共 API 必须包含 Google 风格的文档字符串
3. **异常处理**: 使用自定义异常类，提供有意义的错误信息
4. **日志记录**: 使用 `logging` 模块，避免 `print`

## 提交规范

我们使用 [Conventional Commits](https://www.conventionalcommits.org/) 规范：

```
<type>(<scope>): <subject>

<body>

<footer>
```

### 类型 (Type)

- **feat**: 新功能
- **fix**: Bug 修复
- **docs**: 文档更新
- **style**: 代码格式（不影响功能的更改）
- **refactor**: 代码重构
- **perf**: 性能优化
- **test**: 测试相关
- **chore**: 构建过程或辅助工具的变动

### 示例

```
feat(rules): add support for YZU design thesis format

Add new rule set for Yangzhou University graduation design
thesis format with specific requirements for cover page
and table of contents.

Closes #123
```

```
fix(processor): handle empty headings gracefully

Previously, empty headings would cause an IndexError.
Now they are skipped with a warning logged.

Fixes #456
```

## 审查流程

1. **自动检查**: CI 会自动运行测试和代码检查
2. **代码审查**: 维护者会审查代码质量和设计
3. **反馈处理**: 根据反馈进行修改
4. **合并**: 通过审查后，维护者会合并 PR

### 审查标准

- 代码符合项目风格
- 包含适当的测试
- 文档已更新
- 提交信息清晰规范
- 不引入回归问题

## 添加新规则

如果您想为新的学校或格式添加支持，请参考详细的 [规则开发指南](docs/rule_development.md)。

### 快速步骤

1. **创建规则文件**：在 `ruledoc/rules/` 下创建新的规则目录
   ```
   ruledoc/rules/
   └── your_school/
       ├── __init__.py
       └── your_thesis.py
   ```

2. **实现规则类**：继承 `FormatRule` 基类，实现必需的方法
   ```python
   from ruledoc.rules.base import FormatRule, register_rule

   @register_rule
   class YourThesisRule(FormatRule):
       @property
       def name(self):
           return 'your_school_thesis'
       
       # 实现其他必需方法...
   ```

3. **添加测试**：创建 `tests/rules/test_your_school.py`

4. **提交 PR**：遵循 [提交规范](#提交规范)

### 规则开发要点

- **命名规范**：`{学校缩写}_{文档类型}`，如 `yzu_thesis`, `pku_design`
- **必需实现**：`name`, `description`, `get_page_settings`, `get_font_settings`, `get_heading_format`
- **可选扩展**：段落类型检测、特殊格式化、页眉页脚等

详细文档：[docs/rule_development.md](docs/rule_development.md)

## 获取帮助

如果您在贡献过程中需要帮助：

- 查看 [文档](https://github.com/xiaoyu-778/ruledoc#readme)
- 在 [Discussions](https://github.com/xiaoyu-778/ruledoc/discussions) 中提问
- 发送邮件至 3389357760@qq.com

## 许可证

通过贡献代码，您同意您的贡献将在 [MIT 许可证](LICENSE) 下发布。

---

再次感谢您的贡献！🎉
