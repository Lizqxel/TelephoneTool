"""J:COM料金シミュレーションの入出力モデルと入力検証。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from typing import List, Optional, Tuple
from zoneinfo import ZoneInfo
import re
import unicodedata
import uuid


JST = ZoneInfo("Asia/Tokyo")


class ResidenceType(str, Enum):
    DETACHED = "sdu"
    APARTMENT = "mdu"

    @property
    def label(self) -> str:
        return "戸建住宅" if self is ResidenceType.DETACHED else "集合住宅"


class SimulationStatus(str, Enum):
    SUCCESS = "success"
    PARTIAL = "partial"
    CANCELLED = "cancelled"
    ADDRESS_NOT_FOUND = "address_not_found"
    OUT_OF_AREA = "out_of_area"
    COURSE_UNAVAILABLE = "course_unavailable"
    COMMUNICATION_ERROR = "communication_error"
    EXTRACTION_ERROR = "extraction_error"
    ERROR = "error"


class AgeBracket(str, Enum):
    UNDER_22 = "22歳以下"
    UNDER_26 = "26歳以下"
    OVER_27 = "27歳以上"


def normalize_postal_code(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    normalized = re.sub(r"[-\s]", "", normalized)
    if not re.fullmatch(r"\d{7}", normalized):
        raise ValueError("郵便番号は7桁で入力してください。")
    return normalized


def parse_birth_date(value: str) -> date:
    if not value:
        raise ValueError("生年月日を入力してください。")
    try:
        parts = [int(part) for part in re.split(r"[/-]", value.strip())]
        if len(parts) != 3:
            raise ValueError
        result = date(*parts)
    except (TypeError, ValueError):
        raise ValueError("生年月日が正しい日付ではありません。") from None
    today = datetime.now(JST).date()
    if result > today:
        raise ValueError("未来の生年月日は指定できません。")
    return result


def calculate_full_age(birth_date: date, today: Optional[date] = None) -> int:
    current = today or datetime.now(JST).date()
    if birth_date > current:
        raise ValueError("未来の生年月日は指定できません。")
    return current.year - birth_date.year - (
        (current.month, current.day) < (birth_date.month, birth_date.day)
    )


def age_bracket_for(age: int) -> AgeBracket:
    if age < 0:
        raise ValueError("年齢が不正です。")
    if age <= 22:
        return AgeBracket.UNDER_22
    if age <= 26:
        return AgeBracket.UNDER_26
    return AgeBracket.OVER_27


@dataclass(frozen=True)
class JcomSearchCriteria:
    request_id: str
    generation: int
    product: str
    postal_code: str
    address_original: str
    address_normalized: str
    birth_date: Optional[date]
    age: int
    age_bracket: AgeBracket
    residence_type: ResidenceType
    address_components: Tuple[str, ...] = ()
    service: str = "ネットのみ"
    usage_status: str = "利用していない"
    course: str = "光(N) 1Gコース"

    @classmethod
    def create(
        cls,
        generation: int,
        postal_code: str,
        address: str,
        birth_date_text: str,
        residence_type: ResidenceType,
    ) -> "JcomSearchCriteria":
        from utils.jcom_address_matcher import (
            normalize_address_for_match,
            split_address_components,
        )

        original = (address or "").strip()
        if not original:
            raise ValueError("住所を入力してください。")
        normalized_birth_date = (birth_date_text or "").strip()
        if normalized_birth_date:
            birth_date = parse_birth_date(normalized_birth_date)
            age = calculate_full_age(birth_date)
        else:
            # J:COMの年齢選択画面では、未入力時は27歳以上を既定にする。
            birth_date = None
            age = 27
        return cls(
            request_id=uuid.uuid4().hex,
            generation=generation,
            product="J:COM",
            postal_code=normalize_postal_code(postal_code),
            address_original=original,
            address_normalized=normalize_address_for_match(original),
            birth_date=birth_date,
            age=age,
            age_bracket=age_bracket_for(age),
            residence_type=residence_type,
            address_components=split_address_components(original),
        )


@dataclass(frozen=True)
class AddressCandidate:
    candidate_id: str
    text: str
    special_action: str = ""


@dataclass(frozen=True)
class AddressCandidateRequest:
    request_id: str
    generation: int
    stage_id: str
    selected_address: str
    remaining_address: str
    candidates: List[AddressCandidate]
    selection_kind: str = "building"


@dataclass
class DiscountLine:
    name: str
    amount_yen: Optional[int] = None
    period_text: str = ""
    raw_text: str = ""


@dataclass(frozen=True)
class AddressApproximation:
    requested: str
    selected: str


@dataclass
class JcomSimulationResult:
    request_id: str
    generation: int
    acquired_at: datetime
    input_address: str
    confirmed_address: str
    residence_type: ResidenceType
    age: int
    calculated_age_bracket: AgeBracket
    applied_age_bracket: Optional[AgeBracket]
    age_selection_used: bool
    selected_service: str
    line_type: str
    course: str
    birth_date_provided: bool = True
    contract_period: str = ""
    next_month_price_yen: Optional[int] = None
    discount_period_text: str = ""
    base_price_yen: Optional[int] = None
    discounts: List[DiscountLine] = field(default_factory=list)
    benefits_text: str = ""
    notes_text: str = ""
    source_url: str = ""
    raw_text: str = ""
    screenshot_path: Optional[str] = None
    screenshot_pending: bool = False
    status: SimulationStatus = SimulationStatus.ERROR
    partial_address: bool = False
    address_approximations: List[AddressApproximation] = field(default_factory=list)
    error_message: str = ""

    @property
    def successful(self) -> bool:
        return self.status in (SimulationStatus.SUCCESS, SimulationStatus.PARTIAL)

    def display_text(self) -> str:
        def money(value: Optional[int]) -> str:
            return "未取得" if value is None else f"{value:,}円（税込）"

        applied = (
            self.applied_age_bracket.value
            if self.age_selection_used and self.applied_age_bracket
            else "この経路では年齢選択なし"
        )
        lines = []
        if self.status is SimulationStatus.COURSE_UNAVAILABLE:
            lines.append("判定: 判定不能（光(N) 1Gコース未確認）")
        if self.address_approximations:
            lines.append("【注意】指定の番地・号がなく、近い候補の住所で検索しました。入力住所の料金ではありません。")
        age_condition = (
            f"入力済み生年月日から判定: {self.age}歳／{self.calculated_age_bracket.value}"
            if self.birth_date_provided
            else "生年月日未入力のため自動選択: 27歳以上"
        )
        lines.extend((
            f"対象サービス: {self.selected_service or '未取得'}",
            f"回線種別: {self.line_type or '未取得'}",
            f"コース: {self.course or '未取得'}",
            f"月額（加入翌月）: {money(self.next_month_price_yen)}",
            f"割引期間: {self.discount_period_text or '未取得'}",
            f"基本料金: {money(self.base_price_yen)}",
            f"契約期間: {self.contract_period or '未取得'}",
            f"住宅区分: {self.residence_type.label}",
            age_condition,
            f"サイト適用条件: {applied}",
        ))
        if self.address_approximations:
            lines.append("近い住所として選択した候補:")
            lines.extend(
                f"・指定 {item.requested} → 選択 {item.selected}"
                for item in self.address_approximations
            )
        elif self.partial_address:
            lines.append("住所確認: 住所の一部未確認")
        if self.partial_address or self.address_approximations or self.status is SimulationStatus.COURSE_UNAVAILABLE:
            lines.extend((
                f"入力住所: {self.input_address}",
                f"サイト確定住所: {self.confirmed_address or '未取得'}",
            ))
        if self.discounts:
            lines.append("\n割引明細")
            lines.extend(f"・{item.raw_text or item.name}" for item in self.discounts)
        if self.benefits_text:
            lines.extend(["\n特典・条件", self.benefits_text])
        if self.notes_text:
            lines.extend(["\n注意事項", self.notes_text])
        if self.error_message:
            lines.extend(["\n取得状態", self.error_message])
        return "\n".join(lines)
