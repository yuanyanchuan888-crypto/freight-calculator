from decimal import Decimal
import tkinter as tk
import unittest
from tkinter import TclError

from freight_calculator.models import CargoLine
from freight_calculator.ui import FreightCalculatorApp


class UserInterfaceSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.root = tk.Tk()
        except TclError as exc:
            raise unittest.SkipTest(f"Tk display is unavailable: {exc}") from exc
        cls.root.withdraw()
        cls.app = FreightCalculatorApp(cls.root)

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "root"):
            cls.root.destroy()

    def test_layout_calculation_and_custom_settings_controls(self):
        app = self.app
        self.assertEqual(len(app.tree["columns"]), 12)
        self.assertEqual(app.root.title(), "国际物流计费计算器 | International Freight Calculator")

        app.lines = [
            CargoLine(
                "UI smoke carton", 2, Decimal("50"), Decimal("40"), Decimal("30"),
                "cm", Decimal("8"), "kg",
            )
        ]
        app._recalculate_safely()
        self.root.update_idletasks()
        row = app.tree.item(app.tree.get_children()[0], "values")
        self.assertEqual(row[0], "1")
        self.assertEqual(row[1], "UI smoke carton")
        self.assertEqual(app.metric_vars["actual"].get(), "16")
        self.assertEqual(app.metric_vars["volume"].get(), "0.12")

        app.divisor_var.set("自定义")
        app._sync_settings_fields()
        self.assertEqual(str(app.custom_divisor_entry.cget("state")), "normal")
        app.divisor_var.set("6000")
        app._sync_settings_fields()
        self.assertEqual(str(app.custom_divisor_entry.cget("state")), "disabled")


if __name__ == "__main__":
    unittest.main()
