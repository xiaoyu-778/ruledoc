"""CLI 入口点 - 支持 python -m ruledoc

用法:
    python -m ruledoc input.md --rule yzu
    python -m ruledoc --list-rules
"""

import sys

from ruledoc.cli import main

if __name__ == "__main__":
    sys.exit(main())
