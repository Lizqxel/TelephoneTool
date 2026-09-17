"""J:COM住所候補の段階照合。Qt/Seleniumには依存しない。"""

from __future__ import annotations

from dataclasses import dataclass
import re
import unicodedata
from typing import List, Optional, Sequence, Tuple


NEXT_WITH_CURRENT = "next_with_current"
MISSING_ADDRESS = "missing_address"


def split_address_components(address: str) -> Tuple[str, ...]:
    """既存提供判定と同じ都道府県→市区町村→町名→番号の順で分離する。

    建物名の長音符を壊さず、J:COMの段階候補に使う成分だけを返す。
    """
    value = unicodedata.normalize("NFKC", address or "")
    value = re.sub(r"[\s\u3000]", "", value)
    prefecture_match = re.match(r"^(東京都|北海道|(?:京都|大阪)府|.+?県)", value)
    if prefecture_match:
        value = value[prefecture_match.end():]
    # 政令指定都市は「千葉市若葉区」までを行政区画として先に消費する。
    # 汎用の「最初の市区町村」を先に使うと「千葉市」で途切れてしまう。
    city_match = re.match(r"^(.+?市.+?区|.+?郡.+?[町村]|.+?[市区町村])", value)
    if city_match:
        value = value[city_match.end():]

    components: List[str] = []
    chome_match = re.match(r"^(.*?)(\d+丁目)", value)
    if chome_match:
        town, chome = chome_match.groups()
        if town:
            components.append(town)
        components.append(chome)
        value = value[chome_match.end():]
    else:
        town_match = re.match(r"^([^\d]+?)(?=\d|$)", value)
        if town_match:
            town = town_match.group(1)
            if town:
                components.append(town)
            value = value[town_match.end():]

    while value:
        value = value.lstrip("-‐‑‒–—―−")
        number_match = re.match(r"^(\d+)(丁目|番地|番|号)?", value)
        if not number_match:
            break
        number, suffix = number_match.groups()
        components.append(number + ("丁目" if suffix == "丁目" else ""))
        value = value[number_match.end():]
        if not suffix and not value.startswith(("-", "‐", "‑", "‒", "–", "—", "―", "−")):
            break

    if value:
        room_match = re.match(r"^(.+?)(\d+(?:号室|号))$", value)
        if room_match:
            building, room = room_match.groups()
            if building:
                components.append(building)
            components.append(room)
        else:
            components.append(value)
    return tuple(component for component in components if component)


def normalize_address_for_match(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = value.casefold().replace("ヶ", "ケ")
    value = re.sub(r"[\s\u3000,，]", "", value)
    # 長音符は建物名の意味を持つため変更しない。住所番号の区切りだけ統一する。
    value = re.sub(r"(?<=\d)[‐‑‒–—―−-](?=\d)", "-", value)
    return value


def classify_special_candidate(text: str) -> str:
    normalized = normalize_address_for_match(text)
    if "表示中の住所" in normalized and "次へ" in normalized:
        return NEXT_WITH_CURRENT
    if (
        "住所・物件がない" in normalized
        or ("住所および物件が見つからない" in normalized)
        or ("該当する住所がない" in normalized)
    ):
        return MISSING_ADDRESS
    return ""


def _tokens(value: str) -> List[str]:
    value = normalize_address_for_match(value)
    # 数字は必ず独立トークン化し、3と13、206番と206号室を混同しない。
    return re.findall(r"\d+|丁目|番地?|号棟?|号室|室|[a-z]+|[^\d\W]+|[-]", value)


def _candidate_component_value(text: str) -> str:
    """J:COMが候補に付ける市区町村・件数表示を照合用に除く。"""
    value = normalize_address_for_match(text)
    return re.sub(r"の物件\(?\d+件\)?$", "", value)


def _matching_hint_end(value: str, hints: Sequence[str], start: int) -> Optional[int]:
    """valueの末尾に連続する住所成分があれば、消費後位置を返す。"""
    best_end = None
    for end in range(start + 1, len(hints) + 1):
        combined = "".join(hints[start:end])
        variants = (
            combined,
            f"{combined}番",
            f"{combined}番地",
            f"{combined}号",
            f"{combined}号棟",
            f"{combined}号室",
        )
        allow_suffix = not combined.isdigit()
        if any(
            value == variant or (allow_suffix and value.endswith(variant))
            for variant in variants
        ):
            best_end = end
    return best_end


def _matching_hint_span(
    value: str, hints: Sequence[str], start: int
) -> Optional[Tuple[int, int]]:
    """J:COMが町名を暗黙に消費する場合も含め、最初の一致範囲を返す。"""
    for hint_start in range(start, len(hints)):
        end = _matching_hint_end(value, hints, hint_start)
        if end is not None:
            return hint_start, end
    return None


def remaining_address(full_address: str, selected_address: str) -> str:
    if selected_address:
        hints = [normalize_address_for_match(part) for part in split_address_components(full_address)]
        start = 0
        for part in selected_address.split("|"):
            value = _candidate_component_value(part)
            span = _matching_hint_span(value, hints, start)
            if span is not None:
                start = span[1]
        if start:
            return "".join(hints[start:])
    full = normalize_address_for_match(full_address)
    selected = normalize_address_for_match(selected_address)
    full_tokens = _tokens(full)
    selected_tokens = _tokens(selected)
    if selected_tokens:
        selected_index = 0
        last_match = -1
        for index, token in enumerate(full_tokens):
            if token == selected_tokens[selected_index]:
                selected_index += 1
                last_match = index
                if selected_index == len(selected_tokens):
                    return "".join(full_tokens[last_match + 1:])
    if selected and full.startswith(selected):
        return full[len(selected):]
    if selected:
        index = full.find(selected)
        if index >= 0:
            return full[index + len(selected):]
    return full


def confirmed_address_matches(
    input_address: str,
    confirmed_address: str,
    component_hints: Sequence[str] = (),
) -> bool:
    """J:COMが建物名や0号を補った確定住所を、入力成分と照合する。"""
    input_normalized = normalize_address_for_match(input_address)
    confirmed_normalized = normalize_address_for_match(confirmed_address)
    if input_normalized == confirmed_normalized:
        return True

    hints = tuple(component_hints) or split_address_components(input_address)
    if not hints or not confirmed_normalized:
        return False

    confirmed_numbers = re.findall(r"\d+", confirmed_normalized)
    number_index = 0
    for hint in hints:
        normalized = normalize_address_for_match(hint)
        if normalized.isdigit():
            try:
                number_index = confirmed_numbers.index(normalized, number_index) + 1
            except ValueError:
                return False
        elif normalized not in confirmed_normalized:
            return False
    return True


@dataclass(frozen=True)
class MatchDecision:
    candidate_id: Optional[str]
    reason: str
    requires_user: bool
    approximate_from: str = ""
    approximate_to: str = ""

    @property
    def no_match(self) -> bool:
        return self.candidate_id is None and "ありません" in self.reason


def choose_address_candidate(
    full_address: str,
    selected_address: str,
    candidates: Sequence,
    component_hints: Sequence[str] = (),
) -> MatchDecision:
    """現在段階の完全一致を優先し、欠けた数値だけ最も近い候補を選ぶ。"""
    remainder = remaining_address(full_address, selected_address)
    remainder_norm = normalize_address_for_match(remainder)
    remainder_tokens = _tokens(remainder_norm)
    while remainder_tokens and remainder_tokens[0] in ("丁目", "番", "番地", "号"):
        remainder_tokens.pop(0)
    matches = []

    if component_hints:
        selected_parts = [
            normalize_address_for_match(part)
            for part in selected_address.split("|") if part
        ]
        start_index = 0
        normalized_hints = [normalize_address_for_match(part) for part in component_hints]
        for selected_part in selected_parts:
            span = _matching_hint_span(
                _candidate_component_value(selected_part), normalized_hints, start_index
            )
            if span is not None:
                start_index = span[1]
        indexed_matches = []
        for candidate in candidates:
            text = getattr(candidate, "text", "")
            special = getattr(candidate, "special_action", "") or classify_special_candidate(text)
            if special:
                continue
            normalized = _candidate_component_value(text)
            span = _matching_hint_span(normalized, normalized_hints, start_index)
            if span is not None and (span[0] == start_index or not selected_parts):
                indexed_matches.append(
                    (span[0], span[1], getattr(candidate, "candidate_id", None))
                )
                continue

            # 号棟候補だけは建物名の後ろの番号を照合する。任意の建物名中の
            # 数字を番地・号の完全一致とみなすと、異なる住所を選んでしまう。
            if (
                start_index < len(normalized_hints)
                and normalized_hints[start_index].isdigit()
                and normalized.endswith("号棟")
            ):
                numeric_tokens = re.findall(r"\d+", normalized)
                if normalized_hints[start_index] in numeric_tokens:
                    indexed_matches.append(
                        (start_index, start_index + 1, getattr(candidate, "candidate_id", None))
                    )
        if indexed_matches:
            # 同じ町域から「町域のみ」と「町域＋丁目」の両方が提示される。
            # 最初に一致する成分を優先し、その中では入力を最も多く消費する
            # 候補を選ぶ。これにより麻溝台と麻溝台7丁目を曖昧扱いしない。
            earliest = min(start for start, _, _ in indexed_matches)
            earliest_matches = [
                item for item in indexed_matches if item[0] == earliest
            ]
            longest_end = max(end for _, end, _ in earliest_matches)
            earliest_ids = [
                candidate_id
                for _, end, candidate_id in earliest_matches
                if end == longest_end
            ]
            if len(earliest_ids) == 1:
                return MatchDecision(
                    earliest_ids[0],
                    "提供判定方式で分離した住所成分に最長一致",
                    False,
                )
            return MatchDecision(None, "同じ住所成分の候補が複数あります", True)

    if component_hints and selected_address:
        # 番地・号など、現在の数値成分だけを近似する。建物名中の数字や
        # 次段階の数字を拾わず、同距離なら小さい数値を採る。
        if start_index < len(normalized_hints):
            expected = normalized_hints[start_index]
            number_match = re.fullmatch(r"(\d+)(丁目|番地|番|号室|号)?", expected)
            if number_match:
                target = int(number_match.group(1))
                expected_suffix = number_match.group(2) or ""
                numeric_candidates = []
                for candidate in candidates:
                    text = getattr(candidate, "text", "")
                    if getattr(candidate, "special_action", "") or classify_special_candidate(text):
                        continue
                    normalized = _candidate_component_value(text)
                    match = re.fullmatch(r"(\d+)(丁目|番地|番|号|号室)?", normalized)
                    if not match:
                        continue
                    suffix = match.group(2) or ""
                    if bool(suffix == "丁目") != bool(expected_suffix == "丁目"):
                        continue
                    if expected_suffix in ("号", "号室") and suffix not in ("号", "号室"):
                        continue
                    if expected_suffix in ("番", "番地") and suffix not in ("番", "番地"):
                        continue
                    number = int(match.group(1))
                    numeric_candidates.append((
                        abs(number - target), number,
                        getattr(candidate, "candidate_id", None), text,
                    ))
                if numeric_candidates:
                    numeric_candidates.sort(key=lambda item: (item[0], item[1]))
                    best = numeric_candidates[0]
                    if sum(item[:2] == best[:2] for item in numeric_candidates) == 1:
                        return MatchDecision(
                            best[2], "指定の数値候補がないため最も近い値を選択", False,
                            component_hints[start_index], best[3],
                        )
        # 次成分が候補にない場合、後続の号・部屋番号へ飛ばさない。
        next_actions = [
            candidate for candidate in candidates
            if (
                getattr(candidate, "special_action", "")
                or classify_special_candidate(getattr(candidate, "text", ""))
            ) == NEXT_WITH_CURRENT
        ]
        if len(next_actions) == 1 and not remaining_address(full_address, selected_address):
            return MatchDecision(
                getattr(next_actions[0], "candidate_id", None),
                "入力住所の成分を選び終えたため表示中の住所で次へ",
                False,
            )
        return MatchDecision(None, "現在段階に一致する候補がありません", True)

    for candidate in candidates:
        text = getattr(candidate, "text", "")
        candidate_id = getattr(candidate, "candidate_id", None)
        special = getattr(candidate, "special_action", "") or classify_special_candidate(text)
        if special:
            continue
        normalized = normalize_address_for_match(text)
        if not normalized:
            continue
        candidate_tokens = _tokens(normalized)
        if not candidate_tokens:
            continue

        # 一段目は都道府県・市区町村がサイト側で既に消費されているため
        # 入力内の一意成分を許す。二段目以降は未消費部分の先頭だけを見る。
        exact_prefix = remainder_norm.startswith(normalized)
        token_prefix = remainder_tokens[: len(candidate_tokens)] == candidate_tokens
        initial_component_match = not selected_address and any(
            remainder_tokens[index:index + len(candidate_tokens)] == candidate_tokens
            for index in range(0, max(0, len(remainder_tokens) - len(candidate_tokens) + 1))
        )
        if exact_prefix or token_prefix or initial_component_match:
            matches.append(candidate_id)

    if len(matches) == 1:
        return MatchDecision(matches[0], "未消費住所の先頭成分と一意に一致", False)
    if not matches:
        return MatchDecision(None, "現在段階に一致する候補がありません", True)
    return MatchDecision(None, "現在段階に一致する候補が複数あります", True)
