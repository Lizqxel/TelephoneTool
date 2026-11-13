"""本番と同じCTI集計処理を手動実行する。本人のみのチェックを事前に設定する。"""
import argparse
from dataclasses import asdict
from datetime import datetime
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from services.call_statistics import StatisticsStore
from services.cti_statistics_service import CtiStatisticsCollector


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, help="Optional SQLite path to save the snapshot")
    parser.add_argument("--unit", choices=("seconds", "minutes"), default="seconds")
    parser.add_argument("--clear-after-save", action="store_true", help="Clear CTI history after a successful database save")
    args = parser.parse_args()
    if args.clear_after_save and not args.database:
        parser.error("--clear-after-save requires --database")
    collector = CtiStatisticsCollector()
    snapshot = collector.collect(args.unit)
    if args.database:
        StatisticsStore(args.database).save(snapshot)
    if args.clear_after_save:
        collector.clear_history()
    output = ROOT / "logs" / "cti_inspection" / ("statistics_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    output.mkdir(parents=True)
    data = asdict(snapshot)
    data["acquired_at"] = snapshot.acquired_at.isoformat()
    (output / "snapshot.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if collector.clear_history_info is not None:
        (output / "clear_history.json").write_text(json.dumps(collector.clear_history_info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(output), "counters": snapshot.counters,
                      "history_rows": snapshot.history_rows,
                      "clear_history": collector.clear_history_info,
                      "talk_seconds": sum(snapshot.talk_seconds.values())}, ensure_ascii=True))


if __name__ == "__main__":
    main()
