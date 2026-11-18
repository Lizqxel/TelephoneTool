"""ランキング・コピー・グラフに共通の集計。時間は秒のまま計算する。"""
from dataclasses import dataclass

from services.call_statistics import format_duration


@dataclass
class RankingRow:
    staff_id: str
    name: str
    calls: int | None
    dialogues: int | None
    seconds: float | None

    @property
    def complete(self):
        return all(value is not None for value in (self.calls, self.dialogues, self.seconds))


@dataclass
class Report:
    rows: list
    people: int
    calls: int
    dialogues: int
    seconds: float

    @property
    def average_dialogue(self):
        return self.seconds / self.dialogues if self.people and self.dialogues else None


def ranking(members, minute=None):
    rows = []
    for member in members:
        slots = list(member.intervals.values()) if minute is None else [member.intervals.get(minute, (None,) * 4)]

        def total(index):
            values = [slot[index] for slot in slots]
            return sum(values) if values and all(value is not None for value in values) else None

        rows.append(RankingRow(member.staff_id, member.name, total(0), total(1), total(3)))
    rows.sort(key=lambda row: (row.seconds is None, -(row.seconds or 0), row.name, row.staff_id))
    complete = [row for row in rows if row.complete]
    return Report(rows, len(complete), sum(row.calls for row in complete),
                  sum(row.dialogues for row in complete), sum(row.seconds for row in complete))


def slot_label(minute):
    return f"{minute // 60:02}:{minute % 60:02}"


def period_label(minute):
    return "取得済み全日" if minute is None else f"{slot_label(minute)}～{slot_label(minute + 30)}"


def summaries(report):
    count = report.people
    whole = (f"【全体】人数 {count}人（対象 {len(report.rows)}人・未取得/未判定 {len(report.rows) - count}人）"
             f"｜コール数 {report.calls if count else '—'}｜対話数 {report.dialogues if count else '—'}"
             f"｜対話時間 {format_duration(report.seconds) if count else '—'}"
             f"｜1件あたり平均 {format_duration(report.average_dialogue) if report.average_dialogue is not None else '—'}")
    average = (f"【1人あたり平均】コール数 {report.calls / count:.2f}｜対話数 {report.dialogues / count:.2f}"
               f"｜対話時間 {format_duration(report.seconds / count)}") if count else "【1人あたり平均】コール数 —｜対話数 —｜対話時間 —"
    return whole, average


def row_text(row):
    return [row.name, "未取得" if row.calls is None else str(row.calls),
            "未取得/未判定" if row.dialogues is None else str(row.dialogues), format_duration(row.seconds)]


def report_text(report, day, unit, minute):
    # 改行やタブを含む氏名でも、貼り付け後の列を崩さない。
    clean = lambda value: str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")
    lines = [f"日付：{clean(day)}　GPユニット：{clean(unit)}　時間帯：{period_label(minute)}", "氏名\tコール数\t対話数\t対話時間"]
    lines += ["\t".join(clean(value) for value in row_text(row)) for row in report.rows]
    lines += [*summaries(report), "対話時間＝見込＋対話有NGの通話時間。平均・全体は3項目すべて取得済みの人が対象。取得途中の時間帯は途中実績。"]
    return "\n".join(lines)
