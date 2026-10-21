from decimal import Decimal
import unittest

from freight_calculator.calculator import calculate, round_up_weight
from freight_calculator.models import (
    CargoLine,
    CalculationSettings,
    FreightCalculationError,
)


class CalculatorTests(unittest.TestCase):
    def test_cm_kg_multiple_pieces(self):
        cargo = CargoLine(
            "Box", 2, Decimal("50"), Decimal("40"), Decimal("30"), "cm",
            Decimal("20"), "kg",
        )
        result = calculate([cargo], CalculationSettings(Decimal("5000")))
        self.assertEqual(result.total_volume_m3, Decimal("0.120"))
        self.assertEqual(result.total_actual_weight_kg, Decimal("40"))
        self.assertEqual(result.total_volumetric_weight_kg, Decimal("24"))
        self.assertEqual(result.selected_chargeable_weight_kg, Decimal("40"))

    def test_inches_and_pounds_convert_to_metric(self):
        cargo = CargoLine(
            "Sample", 1, Decimal("10"), Decimal("10"), Decimal("10"), "in",
            Decimal("2"), "lb",
        )
        result = calculate([cargo], CalculationSettings(Decimal("5000")))
        self.assertEqual(result.lines[0].actual_weight_kg_per_piece, Decimal("0.90718474"))
        self.assertEqual(result.total_volume_m3, Decimal("0.016387064"))
        self.assertEqual(
            result.total_volumetric_weight_kg,
            Decimal("16.387064") / Decimal("5"),
        )

    def test_per_piece_and_shipment_rounding_are_distinct(self):
        lines = [
            CargoLine(
                "A", 1, Decimal("1"), Decimal("1"), Decimal("1"), "cm",
                Decimal("0.4"), "kg",
            ),
            CargoLine(
                "B", 1, Decimal("1"), Decimal("1"), Decimal("1"), "cm",
                Decimal("0.4"), "kg",
            ),
        ]
        piece_settings = CalculationSettings(
            Decimal("6000"), "per_piece", Decimal("1")
        )
        shipment_settings = CalculationSettings(
            Decimal("6000"), "shipment_total", Decimal("1")
        )
        self.assertEqual(calculate(lines, piece_settings).selected_chargeable_weight_kg, Decimal("2"))
        self.assertEqual(
            calculate(lines, shipment_settings).selected_chargeable_weight_kg,
            Decimal("1"),
        )

    def test_rounding_uses_ceiling_increment(self):
        cargo = CargoLine(
            "Box", 1, Decimal("10"), Decimal("10"), Decimal("10"), "cm",
            Decimal("2.51"), "kg",
        )
        result = calculate(
            [cargo],
            CalculationSettings(Decimal("5000"), "shipment_total", Decimal("0.5")),
        )
        self.assertEqual(result.selected_chargeable_weight_kg, Decimal("3.0"))

    def test_multiple_lines_aggregate_each_measure_and_apply_selected_mode(self):
        lines = [
            CargoLine(
                "Carton A", 2, Decimal("50"), Decimal("40"), Decimal("30"), "cm",
                Decimal("9"), "kg",
            ),
            CargoLine(
                "Carton B", 1, Decimal("20"), Decimal("20"), Decimal("20"), "cm",
                Decimal("5"), "kg",
            ),
        ]
        settings = CalculationSettings(
            Decimal("5000"), "per_piece", Decimal("0.5")
        )
        result = calculate(lines, settings)
        self.assertEqual(result.total_volume_m3, Decimal("0.128"))
        self.assertEqual(result.total_actual_weight_kg, Decimal("23"))
        self.assertEqual(result.total_volumetric_weight_kg, Decimal("25.6"))
        self.assertEqual(result.per_piece_chargeable_weight_kg, Decimal("29"))
        self.assertEqual(result.shipment_total_chargeable_weight_kg, Decimal("26.0"))
        self.assertEqual(result.selected_chargeable_weight_kg, Decimal("29"))

    def test_exact_rounding_boundary_is_not_rounded_up_again(self):
        self.assertEqual(round_up_weight(Decimal("2.5"), Decimal("0.5")), Decimal("2.5"))
        self.assertEqual(round_up_weight(Decimal("2.49"), Decimal("0.5")), Decimal("2.5"))
        self.assertEqual(round_up_weight(Decimal("2.51"), Decimal("0.5")), Decimal("3.0"))

    def test_round_up_weight_rejects_invalid_values(self):
        for weight, increment in (
            (Decimal("NaN"), None),
            (Decimal("1"), Decimal("0")),
            (Decimal("1"), Decimal("-0.5")),
            (Decimal("1"), Decimal("Infinity")),
        ):
            with self.subTest(weight=weight, increment=increment):
                with self.assertRaises(FreightCalculationError):
                    round_up_weight(weight, increment)

    def test_custom_divisor_is_applied(self):
        cargo = CargoLine(
            "Box", 1, Decimal("60"), Decimal("50"), Decimal("40"), "cm",
            Decimal("10"), "kg",
        )
        result_5000 = calculate([cargo], CalculationSettings(Decimal("5000")))
        result_6000 = calculate([cargo], CalculationSettings(Decimal("6000")))
        self.assertEqual(result_5000.total_volumetric_weight_kg, Decimal("24"))
        self.assertEqual(result_6000.total_volumetric_weight_kg, Decimal("20"))

    def test_rejects_invalid_dimensions_units_and_divisors(self):
        with self.assertRaises(FreightCalculationError):
            CargoLine("Box", 1, Decimal("0"), Decimal("1"), Decimal("1"), "cm",
                      Decimal("1"), "kg")
        with self.assertRaises(FreightCalculationError):
            CargoLine("Box", 1, Decimal("1"), Decimal("1"), Decimal("1"), "ft",
                      Decimal("1"), "kg")
        with self.assertRaises(FreightCalculationError):
            CalculationSettings(Decimal("0"))
        with self.assertRaisesRegex(FreightCalculationError, "不能超过"):
            CargoLine("Box", 1_000_001, Decimal("1"), Decimal("1"), Decimal("1"), "cm",
                      Decimal("1"), "kg")
        with self.assertRaises(FreightCalculationError):
            calculate([])

    def test_rejects_non_finite_and_invalid_model_inputs(self):
        invalid_cargo_values = (
            {"length": Decimal("NaN")},
            {"width": Decimal("Infinity")},
            {"actual_weight_per_piece": Decimal("-0.01")},
            {"pieces": True},
            {"pieces": 1.5},
            {"name": "   "},
        )
        for override in invalid_cargo_values:
            with self.subTest(override=override):
                args = dict(
                    name="Box", pieces=1, length=Decimal("1"), width=Decimal("1"),
                    height=Decimal("1"), dimension_unit="cm",
                    actual_weight_per_piece=Decimal("1"), weight_unit="kg",
                )
                args.update(override)
                with self.assertRaises(FreightCalculationError):
                    CargoLine(**args)

        for settings in (
            {"volume_divisor": Decimal("NaN")},
            {"volume_divisor": Decimal("Infinity")},
            {"billing_mode": "other"},
            {"rounding_increment_kg": Decimal("NaN")},
        ):
            with self.subTest(settings=settings):
                args = {
                    "volume_divisor": Decimal("6000"),
                    "billing_mode": "shipment_total",
                    "rounding_increment_kg": None,
                }
                args.update(settings)
                with self.assertRaises(FreightCalculationError):
                    CalculationSettings(**args)


if __name__ == "__main__":
    unittest.main()
