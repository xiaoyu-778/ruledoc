# 规则开发指南

本文档详细介绍如何为 RuleDoc 贡献新的学校格式规则。

## 目录

- [快速开始](#快速开始)
- [规则架构](#规则架构)
- [最小规则示例](#最小规则示例)
- [规则详解](#规则详解)
- [测试规则](#测试规则)
- [提交规则](#提交规则)

## 快速开始

### 1. 理解规则注册机制

RuleDoc 使用装饰器模式实现规则注册。当你编写新规则时，需要使用 `@register_rule` 装饰器：

```python
from ruledoc.rules.base import FormatRule, register_rule

@register_rule  # ← 这个装饰器会自动注册规则
class YourRule(FormatRule):
    ...
```

**注册机制说明：**
- `@register_rule` 装饰器会将规则类添加到全局注册表 `_registered_rules`
- 注册时使用规则的 `name` 属性作为键
- 用户可以通过 `load_rule('rule_name')` 加载已注册的规则
- 支持别名映射，如 `'yzu'` → `'yzu_thesis'`

**为什么需要注册？**
1. **自动发现**：无需手动导入规则类，框架自动发现所有注册规则
2. **动态加载**：可以通过字符串名称动态加载规则
3. **扩展性**：新规则只需添加装饰器即可集成到系统

### 2. 创建规则文件

在 `ruledoc/rules/` 目录下创建新的规则目录和文件：

```
ruledoc/rules/
└── your_school/
    ├── __init__.py
    ├── common.py          # 共享配置（可选）
    └── your_thesis.py     # 规则实现
```

### 3. 最小规则示例

```python
"""
XX大学论文格式规则

基于《XX大学本科生毕业论文格式要求》实现。
"""

from typing import Dict, Tuple
from ruledoc.rules.base import FormatRule, register_rule


# 使用 @register_rule 装饰器注册规则
# 这会将规则自动添加到全局规则注册表，使其可以通过 load_rule() 加载
@register_rule
class XXThesisRule(FormatRule):
    """
    XX大学学位论文格式规则
    
    符合XX大学本科生毕业论文格式要求。
    """
    
    @property
    def name(self) -> str:
        """规则唯一标识名"""
        return 'xx_thesis'
    
    @property
    def description(self) -> str:
        """规则描述"""
        return 'XX大学学位论文格式'
    
    def get_page_settings(self) -> Dict[str, float]:
        """
        页面设置（单位：厘米）
        
        返回:
            - top_margin: 上边距
            - bottom_margin: 下边距
            - left_margin: 左边距
            - right_margin: 右边距
            - gutter: 装订线
        """
        return {
            'top_margin': 2.54,
            'bottom_margin': 2.54,
            'left_margin': 3.17,
            'right_margin': 3.17,
            'gutter': 0.0,
        }
    
    def get_font_settings(self) -> Dict[str, str]:
        """
        字体设置
        
        返回:
            - body_font: 正文字体
            - heading_font: 标题字体
            - caption_font: 题注字体
        """
        return {
            'body_font': '宋体',
            'heading_font': '黑体',
            'caption_font': '宋体',
        }
    
    def get_heading_format(self, level: int) -> Tuple[str, str, float]:
        """
        获取标题格式
        
        Args:
            level: 标题层级 (1-4)
            
        Returns:
            元组 (字体名称, 对齐方式, 字号)
            - 对齐方式: 'left', 'center', 'right'
            - 字号: Word 字号值（如 18.0 表示小二号）
        """
        formats = {
            1: ('黑体', 'center', 18.0),  # 一级标题：黑体居中，小二号
            2: ('黑体', 'left', 15.0),    # 二级标题：黑体左对齐，小三号
            3: ('黑体', 'left', 12.0),    # 三级标题：黑体左对齐，小四号
            4: ('黑体', 'left', 12.0),    # 四级标题：黑体左对齐，小四号
        }
        return formats.get(level, ('宋体', 'left', 12.0))
```

### 3. 注册规则

使用 `@register_rule` 装饰器自动注册规则：

```python
from ruledoc.rules.base import register_rule

@register_rule
class XXThesisRule(FormatRule):
    ...
```

## 规则架构

### 类继承关系

```
FormatRule (ABC)
    ├── 抽象方法（必须实现）
    │   ├── name
    │   ├── description
    │   ├── get_page_settings
    │   ├── get_font_settings
    │   └── get_heading_format
    ├── 可重写方法
    │   ├── detect_paragraph_type
    │   ├── format_paragraph
    │   ├── get_header_footer_settings
    │   └── ...
    └── 工具方法
        ├── get_validated_page_settings
        └── ...
```

### 规则加载机制

1. 规则类使用 `@register_rule` 装饰器注册
2. 通过 `load_rule('rule_name')` 加载规则实例
3. 支持别名映射（如 'yzu' → 'yzu_thesis'）

## 规则详解

### 必需实现的方法

#### 1. `name` 属性

```python
@property
def name(self) -> str:
    return 'xx_thesis'  # 小写字母+下划线
```

**命名规范**：
- 使用小写字母和下划线
- 格式：`{学校缩写}_{文档类型}`
- 示例：`yzu_thesis`, `pku_design`, `thu_report`

#### 2. `description` 属性

```python
@property
def description(self) -> str:
    return 'XX大学学位论文格式'  # 中文描述
```

#### 3. `get_page_settings()`

```python
def get_page_settings(self) -> Dict[str, float]:
    return {
        'top_margin': 2.2,      # 上边距 (cm)
        'bottom_margin': 2.2,   # 下边距 (cm)
        'left_margin': 2.5,     # 左边距 (cm)
        'right_margin': 2.0,    # 右边距 (cm)
        'gutter': 0.5,          # 装订线 (cm)
    }
```

#### 4. `get_font_settings()`

```python
def get_font_settings(self) -> Dict[str, str]:
    return {
        'body_font': '宋体',
        'heading_font': '黑体',
        'caption_font': '宋体',
        'english_font': 'Times New Roman',
    }
```

#### 5. `get_heading_format()`

```python
def get_heading_format(self, level: int) -> Tuple[str, str, float]:
    """
    Args:
        level: 1=一级标题, 2=二级标题, 3=三级标题, 4=四级标题
        
    Returns:
        (字体, 对齐方式, 字号)
    """
    formats = {
        1: ('黑体', 'center', 18.0),  # 小二号
        2: ('黑体', 'left', 15.0),    # 小三号
        3: ('黑体', 'left', 12.0),    # 小四号
        4: ('黑体', 'left', 12.0),    # 小四号
    }
    return formats.get(level, ('宋体', 'left', 12.0))
```

**字号对照表**：

| 中文字号 | 磅值 (pt) | 用途 |
|---------|----------|------|
| 初号 | 42.0 | 封面大标题 |
| 小初 | 36.0 | - |
| 一号 | 26.0 | - |
| 小一 | 24.0 | - |
| 二号 | 22.0 | - |
| 小二 | 18.0 | 论文标题 |
| 三号 | 16.0 | - |
| 小三 | 15.0 | 一级标题 |
| 四号 | 14.0 | 二级标题 |
| 小四 | 12.0 | 正文、三级标题 |
| 五号 | 10.5 | 题注 |
| 小五 | 9.0 | 页眉页脚 |

### 可选重写的方法

#### 段落类型检测

```python
def detect_paragraph_type(self, para, ctx) -> str:
    """
    检测段落类型
    
    返回类型：
    - 'title': 论文标题
    - 'heading': 章节标题
    - 'abstract': 摘要内容
    - 'references': 参考文献
    - 'caption': 图表题注
    - 'body': 正文
    - 'empty': 空段落
    """
    text = para.text.strip()
    
    # 检测摘要
    if '摘要' in text and len(text) < 10:
        return 'abstract_heading'
    
    # 检测参考文献
    if text.startswith('[') and text[1:2].isdigit():
        return 'references'
    
    # 默认返回正文
    return 'body'
```

#### 段落格式化

```python
def format_paragraph(self, para, para_type: str, ctx) -> None:
    """
    根据段落类型应用格式
    """
    if para_type == 'title':
        self._format_title(para)
    elif para_type == 'heading':
        self._format_heading(para, ctx)
    elif para_type == 'abstract':
        self._format_abstract(para)
    else:
        self._format_body(para)
```

#### 页眉页脚设置

```python
def get_header_footer_settings(self) -> Dict:
    return {
        'header_text': 'XX大学本科生毕业论文',
        'header_font': '宋体',
        'header_font_size': 9,
        'footer_type': 'page_number',  # 或 'none'
    }
```

## 完整示例：YZU 规则

参考 `ruledoc/rules/yzu/yzu_thesis.py` 了解完整实现，包括：

- 特殊标题检测（摘要、参考文献等）
- 题注格式化（图、表）
- 参考文献自动编号
- 表格三线表格式
- 公式居中编号
- 交叉引用处理

## 测试规则

### 1. 创建测试文件

```python
# tests/rules/test_xx_thesis.py

import pytest
from ruledoc.rules import load_rule


def test_xx_thesis_rule_loads():
    """测试规则可以正常加载"""
    rule = load_rule('xx_thesis')
    assert rule is not None
    assert rule.name == 'xx_thesis'


def test_page_settings():
    """测试页面设置"""
    rule = load_rule('xx_thesis')
    settings = rule.get_page_settings()
    
    assert 'top_margin' in settings
    assert settings['top_margin'] > 0


def test_heading_format():
    """测试标题格式"""
    rule = load_rule('xx_thesis')
    
    font, align, size = rule.get_heading_format(1)
    assert font == '黑体'
    assert align == 'center'
    assert size == 18.0
```

### 2. 运行测试

```bash
pytest tests/rules/test_xx_thesis.py -v
```

### 3. 手动测试

```bash
# 安装开发版本
pip install -e .

# 测试新规则
ruledoc test.md --rule xx_thesis
```

## 提交规则

### 1. 准备提交

- [ ] 规则实现完整
- [ ] 包含单元测试
- [ ] 通过所有测试
- [ ] 代码符合项目风格

### 2. 提交信息规范

```
feat(rules): add XX University thesis format

Add new rule set for XX University graduation thesis format.

Features:
- Page margins: 2.54cm all sides
- Heading format: Heiti, centered for level 1
- Special handling for abstract and references

Closes #123
```

### 3. 创建 PR

1. Fork 仓库
2. 创建特性分支：`git checkout -b feat/xx-thesis-rule`
3. 提交更改：`git commit -m "feat(rules): add XX University thesis format"`
4. 推送到 Fork：`git push origin feat/xx-thesis-rule`
5. 在 GitHub 创建 Pull Request

## 常见问题

### Q: 如何处理不同学院的不同要求？

A: 可以创建子规则或使用配置参数：

```python
class XXEngineeringThesisRule(XXThesisRule):
    @property
    def name(self):
        return 'xx_engineering'
```

### Q: 如何支持双语（中英文）格式？

A: 在 `detect_paragraph_type` 中检测语言：

```python
def detect_paragraph_type(self, para, ctx):
    text = para.text.strip().lower()
    if text in ('摘要', 'abstract'):
        return 'abstract_heading'
```

### Q: 如何调试规则？

A: 使用日志输出：

```python
import logging

logger = logging.getLogger(__name__)

def detect_paragraph_type(self, para, ctx):
    text = para.text.strip()
    logger.debug(f"检测段落: {text[:20]}...")
    # ...
```

## 参考资源

- [扬州大学规则示例](../ruledoc/rules/yzu/yzu_thesis.py)
- [FormatRule 基类](../ruledoc/rules/base.py)
- [python-docx 文档](https://python-docx.readthedocs.io/)

---

如有问题，欢迎在 [Discussions](https://github.com/xiaoyu-778/ruledoc/discussions) 中提问！
