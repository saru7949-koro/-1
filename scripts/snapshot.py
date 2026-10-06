"""原本/ を 作業/snapshots/<日付>/ にコピーする。原本は読み取りのみ。

コピー前後で原本の SHA-256 を照合し、原本が変わっていないことを確認する。
"""
import hashlib
import shutil
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "原本"


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def take_snapshot(day=None):
    day = day or date.today().isoformat()
    dst = ROOT / "作業" / "snapshots" / day
    dst.mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in SRC.iterdir() if p.is_file() and not p.name.startswith("."))
    before = {p.name: sha256(p) for p in files}
    for p in files:
        target = dst / p.name
        if not target.exists() or sha256(target) != before[p.name]:
            shutil.copy2(p, target)
            target.chmod(0o644)
    after = {p.name: sha256(p) for p in files}
    if before != after:
        sys.exit("原本のハッシュが変化しました。処理を中止します。")
    (dst / "MANIFEST.sha256").write_text(
        "".join(f"{h}  {n}\n" for n, h in before.items()), encoding="utf-8")
    return dst


if __name__ == "__main__":
    print(take_snapshot(sys.argv[1] if len(sys.argv) > 1 else None))
