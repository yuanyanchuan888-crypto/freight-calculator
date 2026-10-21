"""Validated input and result models for freight calculations."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Literal


BillingMode = Literal["per_piece", "shipment_total"]
DimensionUnit = Literal["cm", "in"]
WeightUnit = Literal["kg", "lb"]
MAX_PIECES_PER_LINE = 1_000_000
MAX_INPUT_MAGNITUDE = Decimal("1000000")


class FreightCalculationError(ValueError):
    """Raised when an input cannot be used for a freight calculation."""


def as_decimal(value: object, label: str) -> Decimal:
    try:
        number = value if isinstance(value, Decimal) else Decimal(str(value).strip())
    except (InvalidOperation, ValueError, AttributeError):
        raise FreightCalculationError(f"{label}必须是有效数字。") from None
    if not number.is_finite():
        raise FreightCalculationError(f"{label}必须是有限数字。")
    return number


@dataclass(frozen=True)
class CargoLine:
    name: str
    pieces: int
    length: Decimal
    width: Decimal
    height: Decimal
    dimension_unit: DimensionUnit
    actual_weight_per_piece: Decimal
    weight_unit: WeightUnit

    def __post_init__(self) -> None:
        name = str(self.name).strip()
        if not name:
            raise FreightCalculationError("货物名称不能为空。")
        if len(name) > 200:
            raise FreightCalculationError("货物名称不能超过200个字符。")
        object.__setattr__(self, "name", name)

        if isinstance(self.pieces, bool) or not isinstance(self.pieces, int):
            raise FreightCalculationError("件数必须是正整数。")
        if self.pieces < 1:
            raise FreightCalculationError("件数必须至少为1。")
        if self.pieces > MAX_PIECES_PER_LINE:
            raise FreightCalculationError(
                f"单行件数不能超过{MAX_PIECES_PER_LINE:,}。"
            )

        if self.dimension_unit not in ("cm", "in"):
            raise FreightCalculationError("尺寸单位只能是厘米(cm)或英寸(in)。")
        if self.weight_unit not in ("kg", "lb"):
            raise FreightCalculationError("重量单位只能是公斤(kg)或磅(lb)。")

        for field_name, label in (
            ("length", "长度"),
            ("width", "宽度"),
            ("height", "高度"),
            ("actual_weight_per_piece", "单件实际重量"),
        ):
            number = as_decimal(getattr(self, field_name), label)
            if number <= 0:
                raise FreightCalculationError(f"{label}必须大于0。")
            if number > MAX_INPUT_MAGNITUDE:
                raise FreightCalculationError(f"{label}不能超过1,000,000。")
            object.__setattr__(self, field_name, number)


@dataclass(frozen=True)
class CalculationSettings:
    volume_divisor: Decimal = Decimal("6000")
    billing_mode: BillingMode = "shipment_total"
    rounding_increment_kg: Decimal | None = None

    def __post_init__(self) -> None:
        divisor = as_decimal(self.volume_divisor, "体积重除数")
        if divisor <= 0:
            raise FreightCalculationError("体积重除数必须大于0。")
        object.__setattr__(self, "volume_divisor", divisor)

        if self.billing_mode not in ("per_piece", "shipment_total"):
            raise FreightCalculationError("计费口径无效。")

        if self.rounding_increment_kg is not None:
            increment = as_decimal(self.rounding_increment_kg, "进位单位")
            if increment <= 0:
                raise FreightCalculationError("进位单位必须大于0，或选择不进位。")
            if increment > MAX_INPUT_MAGNITUDE:
                raise FreightCalculationError("进位单位不能超过1,000,000 kg。")
            object.__setattr__(self, "rounding_increment_kg", increment)
