"""スナップショット取得 → 分析 → 出力/ へ書き出し。

使い方: python3 scripts/run.py [YYYY-MM-DD]   (省略時は今日の日付)
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from analyze import main  # noqa: E402
from snapshot import take_snapshot  # noqa: E402

if __name__ == "__main__":
    day = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    print("snapshot:", take_snapshot(day))
    main(day)
