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

    @property
    def no_match(self) -> bool:
        return self.candidate_id is None and "ありません" in self.reason


def choose_address_candidate(
    full_address: str,
    selected_address: str,
    candidates: Sequence,
    component_hints: Sequence[str] = (),
) -> MatchDecision:
    """現在段階に完全一致する一意候補だけを自動選択する。"""
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
            if span is not None:
                indexed_matches.append((span[0], getattr(candidate, "candidate_id", None)))
                continue

            # 物件名候補の中に次の番号成分が含まれる経路がある。
            # 数値トークンの完全一致だけを認め、3と13は区別する。
            if start_index < len(normalized_hints) and normalized_hints[start_index].isdigit():
                numeric_tokens = re.findall(r"\d+", normalized)
                if normalized_hints[start_index] in numeric_tokens:
                    indexed_matches.append((start_index, getattr(candidate, "candidate_id", None)))
        if indexed_matches:
            earliest = min(index for index, _ in indexed_matches)
            earliest_ids = [candidate_id for index, candidate_id in indexed_matches if index == earliest]
            if len(earliest_ids) == 1:
                return MatchDecision(earliest_ids[0], "提供判定方式で分離した住所成分と一致", False)
            return MatchDecision(None, "同じ住所成分の候補が複数あります", True)

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
