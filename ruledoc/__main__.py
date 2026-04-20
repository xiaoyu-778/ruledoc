"""CLI 入口点 - 支持 python -m ruledoc

用法:
    python -m ruledoc input.md --rule yzu
    python -m ruledoc --list-rules
"""

import sys

# 使用相对导入，确保在包内运行时能正确找到模块
from .cli import main

if __name__ == "__main__":
    sys.exit(main())
