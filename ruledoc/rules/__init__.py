"""规则模块

包含规则注册机制和所有学校格式规则。

使用方式:
    from ruledoc.rules import register_rule, load_rule, list_available_rules

    @register_rule
    class MyRule(FormatRule):
        ...

    rule = load_rule('myrule')

目录结构:
    rules/
    ├── __init__.py
    ├── base.py
    ├── yzu/                    # 扬州大学规则目录
    │   ├── __init__.py
    │   ├── common.py
    │   ├── yzu_thesis.py
    │   └── yzu_design.py
    └── pku/                    # 北京大学规则目录 (未来扩展)
        ├── __init__.py
        ├── common.py
        ├── pku_thesis.py
        └── pku_design.py
"""

import os
from typing import List

from ruledoc.rules.base import (
    FormatRule,
    _registered_rules,
    list_available_rules,
    load_rule,
    register_rule,
)


def _auto_discover_rules() -> List[str]:
    """自动发现并导入规则模块

    扫描 rules 目录下的：
    1. .py 文件（除 __init__.py 和 base.py）
    2. 子目录（包含 __init__.py 的目录）

    自动导入以触发 @register_rule 装饰器。

    Returns:
        已导入的规则模块名称列表
    """
    rules_dir = os.path.dirname(__file__)
    discovered = []

    for item in os.listdir(rules_dir):
        item_path = os.path.join(rules_dir, item)

        if os.path.isfile(item_path) and item.endswith(".py"):
            if item in ("__init__.py", "base.py"):
                continue
            module_name = item[:-3]
            try:
                __import__(f"ruledoc.rules.{module_name}", fromlist=[module_name])
                discovered.append(module_name)
            except ImportError as e:
                import logging

                logging.warning(f"无法导入规则模块 {module_name}: {e}")

        elif os.path.isdir(item_path):
            init_file = os.path.join(item_path, "__init__.py")
            if os.path.exists(init_file):
                try:
                    __import__(f"ruledoc.rules.{item}", fromlist=[item])
                    discovered.append(item)
                except ImportError as e:
                    import logging

                    logging.warning(f"无法导入规则包 {item}: {e}")

    return discovered


_auto_discover_rules()

__all__ = [
    "FormatRule",
    "register_rule",
    "load_rule",
    "list_available_rules",
    "_registered_rules",
]
