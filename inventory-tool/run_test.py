"""標準ライブラリ unittest でテストを実行（追加テスト依存なし）。"""

import sys
import unittest
from pathlib import Path


if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    sys.path.insert(0, str(base))
    suite = unittest.defaultTestLoader.discover(str(base / "tests"))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    raise SystemExit(0 if result.wasSuccessful() else 1)
