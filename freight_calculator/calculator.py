"""Explicit freight volume-weight and chargeable-weight rules."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING
from typing import Iterable

from .models import CargoLine, CalculationSettings, FreightCalculationError, as_decimal


CM_PER_INCH = Decimal("2.54")
KG_PER_LB = Decimal("0.45359237")
CM3_PER_M3 = Decimal("1000000")


@dataclass(frozen=True)
class LineCalculation:
    cargo: CargoLine
    volume_m3: Decimal
    actual_weight_kg_per_piece: Decimal
    actual_weight_kg: Decimal
    volumetric_weight_kg_per_piece: Decimal
    volumetric_weight_kg: Decimal
    per_piece_chargeable_kg: Decimal


@dataclass(frozen=True)
class CalculationResult:
    lines: tuple[LineCalculation, ...]
    total_volume_m3: Decimal
    total_actual_weight_kg: Decimal
    total_volumetric_weight_kg: Decimal
    per_piece_chargeable_weight_kg: Decimal
    shipment_total_chargeable_weight_kg: Decimal
    selected_chargeable_weight_kg: Decimal
    settings: CalculationSettings


def _dimensions_cm(cargo: CargoLine) -> tuple[Decimal, Decimal, Decimal]:
    factor = CM_PER_INCH if cargo.dimension_unit == "in" else Decimal("1")
    return cargo.length * factor, cargo.width * factor, cargo.height * factor


def _weight_kg(value: Decimal, unit: str) -> Decimal:
    return value * KG_PER_LB if unit == "lb" else value


def round_up_weight(weight_kg: Decimal, increment_kg: Decimal | None) -> Decimal:
    """Round positive weight upward to a user-selected increment; None means exact."""
    weight_kg = as_decimal(weight_kg, "重量")
    if weight_kg < 0:
        raise FreightCalculationError("重量不能为负数。")
    if increment_kg is None:
        return weight_kg
    increment_kg = as_decimal(increment_kg, "进位单位")
    if increment_kg <= 0:
        raise FreightCalculationError("进位单位必须大于0，或选择不进位。")
    return (weight_kg / increment_kg).to_integral_value(rounding=ROUND_CEILING) * increment_kg


def calculate(
    cargo_lines: Iterable[CargoLine],
    settings: CalculationSettings | None = None,
) -> CalculationResult:
    settings = settings or CalculationSettings()
    lines = tuple(cargo_lines)
    if not lines:
        raise FreightCalculationError("请至少添加一行货物资料。")

    calculated: list[LineCalculation] = []
    for cargo in lines:
        length_cm, width_cm, height_cm = _dimensions_cm(cargo)
        volume_per_piece_m3 = length_cm * width_cm * height_cm / CM3_PER_M3
        actual_per_piece_kg = _weight_kg(cargo.actual_weight_per_piece, cargo.weight_unit)
        volumetric_per_piece_kg = (
            length_cm * width_cm * height_cm / settings.volume_divisor
        )
        per_piece_chargeable = round_up_weight(
            max(actual_per_piece_kg, volumetric_per_piece_kg),
            settings.rounding_increment_kg,
        ) * cargo.pieces

        calculated.append(
            LineCalculation(
                cargo=cargo,
                volume_m3=volume_per_piece_m3 * cargo.pieces,
                actual_weight_kg_per_piece=actual_per_piece_kg,
                actual_weight_kg=actual_per_piece_kg * cargo.pieces,
                volumetric_weight_kg_per_piece=volumetric_per_piece_kg,
                volumetric_weight_kg=volumetric_per_piece_kg * cargo.pieces,
                per_piece_chargeable_kg=per_piece_chargeable,
            )
        )

    total_volume = sum((line.volume_m3 for line in calculated), Decimal("0"))
    total_actual = sum((line.actual_weight_kg for line in calculated), Decimal("0"))
    total_volumetric = sum(
        (line.volumetric_weight_kg for line in calculated), Decimal("0")
    )
    per_piece_total = sum(
        (line.per_piece_chargeable_kg for line in calculated), Decimal("0")
    )
    shipment_total = round_up_weight(
        max(total_actual, total_volumetric),
        settings.rounding_increment_kg,
    )
    selected = (
        per_piece_total
        if settings.billing_mode == "per_piece"
        else shipment_total
    )

    return CalculationResult(
        lines=tuple(calculated),
        total_volume_m3=total_volume,
        total_actual_weight_kg=total_actual,
        total_volumetric_weight_kg=total_volumetric,
        per_piece_chargeable_weight_kg=per_piece_total,
        shipment_total_chargeable_weight_kg=shipment_total,
        selected_chargeable_weight_kg=selected,
        settings=settings,
    )
