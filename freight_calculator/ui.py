"""Chinese Tkinter interface for the local freight calculator."""

from __future__ import annotations

import tkinter as tk
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .calculator import CalculationResult, calculate
from .excel_io import read_workbook, write_results, write_template
from .models import CargoLine, CalculationSettings, FreightCalculationError


MODE_LABELS = {
    "逐件计算（每件取大后合计）": "per_piece",
    "整票汇总（合计后取大）": "shipment_total",
}
MODE_NAMES = {
    "per_piece": "逐件计算（每件取大后合计）",
    "shipment_total": "整票汇总（合计后取大）",
}


def _display(value: Decimal, places: int = 4) -> str:
    text = f"{value:,.{places}f}"
    return text.rstrip("0").rstrip(".") if "." in text else text


class FreightCalculatorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("国际物流计费计算器 | International Freight Calculator")
        self.root.geometry("1440x900")
        self.root.minsize(1180, 760)
        self.root.configure(background="#F2F5F8")
        self.lines: list[CargoLine] = []
        self.last_result: CalculationResult | None = None

        self._configure_style()
        self._build_layout()
        self._refresh_table()
        self._update_summary()

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("TFrame", background="#F2F5F8")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat", borderwidth=0)
        style.configure("Panel.TLabelframe", background="#FFFFFF", borderwidth=1,
                        relief="solid", padding=10)
        style.configure("Panel.TLabelframe.Label", background="#FFFFFF",
                        foreground="#173B57", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TLabel", background="#F2F5F8", foreground="#233746")
        style.configure("Muted.TLabel", background="#F2F5F8", foreground="#617384",
                        font=("Microsoft YaHei UI", 9))
        style.configure("Hint.TLabel", background="#FFFFFF", foreground="#728292",
                        font=("Microsoft YaHei UI", 9))
        style.configure("HeaderTitle.TLabel", background="#173B57", foreground="#FFFFFF",
                        font=("Microsoft YaHei UI", 19, "bold"))
        style.configure("HeaderSub.TLabel", background="#173B57", foreground="#D8E5EE",
                        font=("Microsoft YaHei UI", 10))
        style.configure("Section.TLabel", background="#F2F5F8", foreground="#173B57",
                        font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("SectionHint.TLabel", background="#F2F5F8", foreground="#728292",
                        font=("Microsoft YaHei UI", 9))
        style.configure("MetricLabel.TLabel", background="#FFFFFF", foreground="#617384",
                        font=("Microsoft YaHei UI", 9))
        style.configure("MetricValue.TLabel", background="#FFFFFF", foreground="#173B57",
                        font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Primary.TButton", background="#147D83", foreground="#FFFFFF",
                        padding=(13, 8), font=("Microsoft YaHei UI", 10, "bold"),
                        borderwidth=0)
        style.map("Primary.TButton", background=[("active", "#0F696E"), ("disabled", "#91B8BA")],
                  foreground=[("disabled", "#F5FAFA")])
        style.configure("TButton", padding=(10, 7), borderwidth=0)
        style.map("TButton", background=[("active", "#DCE8EE")])
        style.configure("TEntry", padding=(7, 5), fieldbackground="#FFFFFF")
        style.configure("TCombobox", padding=(6, 5), fieldbackground="#FFFFFF")
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                        foreground="#263746", rowheight=32, font=("Microsoft YaHei UI", 9),
                        borderwidth=0)
        style.configure("Treeview.Heading", background="#E7EEF3", foreground="#173B57",
                        font=("Microsoft YaHei UI", 9, "bold"), padding=(8, 9), relief="flat")
        style.map("Treeview", background=[("selected", "#D7EBED")],
                  foreground=[("selected", "#173B57")])

    def _build_layout(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(3, weight=1)

        header = tk.Frame(self.root, background="#173B57", padx=26, pady=17)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        tk.Label(header, text="国际物流计费计算器", bg="#173B57", fg="#FFFFFF",
                 font=("Microsoft YaHei UI", 19, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="International Freight Calculator  ·  本地计算，数据不会自动上传",
            bg="#173B57", fg="#D8E5EE", font=("Microsoft YaHei UI", 9),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        tk.Label(header, text="LOCAL WORKSPACE", bg="#27536E", fg="#EAF3F8",
                 padx=11, pady=6, font=("Segoe UI", 8, "bold")).grid(
                     row=0, column=1, rowspan=2, sticky="e", padx=(18, 0))

        settings = ttk.Frame(self.root, padding=(22, 13, 22, 7))
        settings.grid(row=1, column=0, sticky="ew")
        settings.grid_columnconfigure(0, weight=4)
        settings.grid_columnconfigure(1, weight=2)

        self.divisor_var = tk.StringVar(value="6000")
        self.custom_divisor_var = tk.StringVar(value="5000")
        self.mode_var = tk.StringVar(value=MODE_NAMES["shipment_total"])
        self.rounding_var = tk.StringVar(value="不进位")
        self.custom_rounding_var = tk.StringVar(value="0.5")
        self.default_dimension_var = tk.StringVar(value="cm")
        self.default_weight_var = tk.StringVar(value="kg")

        rules = ttk.LabelFrame(settings, text="计费规则", style="Panel.TLabelframe",
                               padding=(13, 10))
        rules.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        rules.grid_columnconfigure(4, weight=1)
        ttk.Label(rules, text="体积重除数").grid(row=0, column=0, sticky="w", padx=(0, 8))
        divisor_box = ttk.Combobox(rules, textvariable=self.divisor_var, width=9,
                                   values=("5000", "6000", "自定义"), state="readonly")
        divisor_box.grid(row=0, column=1, sticky="w", padx=(0, 7))
        self.custom_divisor_entry = ttk.Entry(rules, textvariable=self.custom_divisor_var, width=9)
        self.custom_divisor_entry.grid(row=0, column=2, sticky="w", padx=(0, 18))

        ttk.Label(rules, text="计费口径").grid(row=0, column=3, sticky="w", padx=(0, 8))
        mode_box = ttk.Combobox(rules, textvariable=self.mode_var, width=25,
                                values=tuple(MODE_LABELS), state="readonly")
        mode_box.grid(row=0, column=4, sticky="ew")

        ttk.Label(rules, text="进位单位").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(9, 0))
        rounding_box = ttk.Combobox(
            rules, textvariable=self.rounding_var, width=9,
            values=("不进位", "0.1 kg", "0.5 kg", "1 kg", "自定义"), state="readonly",
        )
        rounding_box.grid(row=1, column=1, sticky="w", padx=(0, 7), pady=(9, 0))
        self.custom_rounding_entry = ttk.Entry(rules, textvariable=self.custom_rounding_var, width=9)
        self.custom_rounding_entry.grid(row=1, column=2, sticky="w", padx=(0, 18), pady=(9, 0))
        ttk.Label(rules, text="规则请以具体承运渠道的书面确认为准。",
                  style="Hint.TLabel").grid(row=1, column=3, columnspan=2, sticky="w", pady=(9, 0))

        units = ttk.LabelFrame(settings, text="新增货物默认单位", style="Panel.TLabelframe",
                               padding=(13, 10))
        units.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        units.grid_columnconfigure(1, weight=1)
        units.grid_columnconfigure(3, weight=1)
        ttk.Label(units, text="尺寸").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._unit_combo(units, self.default_dimension_var, ("cm", "in"), 1, 0)
        ttk.Label(units, text="重量").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(9, 0))
        self._unit_combo(units, self.default_weight_var, ("kg", "lb"), 1, 1)
        ttk.Label(units, text="仅影响后续新增的货物行。",
                  style="Hint.TLabel").grid(row=2, column=0, columnspan=4, sticky="w", pady=(9, 0))

        divisor_box.bind("<<ComboboxSelected>>", lambda _event: self._sync_settings_fields())
        rounding_box.bind("<<ComboboxSelected>>", lambda _event: self._sync_settings_fields())
        for variable in (
            self.divisor_var, self.custom_divisor_var, self.mode_var,
            self.rounding_var, self.custom_rounding_var,
        ):
            variable.trace_add("write", lambda *_args: self._recalculate_safely())
        self._sync_settings_fields()

        toolbar = ttk.Frame(self.root, padding=(22, 5, 22, 11))
        toolbar.grid(row=2, column=0, sticky="new")
        ttk.Label(toolbar, text="货物明细", style="Section.TLabel").pack(side="left", padx=(0, 6))
        ttk.Label(toolbar, text="录入每种货物的件数、外箱尺寸和单件实际重量。",
                  style="SectionHint.TLabel").pack(side="left", padx=(0, 12))
        ttk.Button(toolbar, text="新增货物", style="Primary.TButton",
                   command=self._add_line).pack(side="left", padx=4)
        ttk.Button(toolbar, text="编辑选中", command=self._edit_line).pack(side="left", padx=4)
        ttk.Button(toolbar, text="删除选中", command=self._delete_line).pack(side="left", padx=4)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=9)
        ttk.Button(toolbar, text="导入 Excel", command=self._import_excel).pack(side="left", padx=4)
        ttk.Button(toolbar, text="导出 Excel", command=self._export_excel).pack(side="left", padx=4)
        ttk.Button(toolbar, text="导出模板", command=self._export_template).pack(side="left", padx=4)
        ttk.Button(toolbar, text="重新计算", style="Primary.TButton",
                   command=self._calculate).pack(side="right", padx=4)

        table_area = ttk.Frame(self.root, padding=(22, 0, 22, 13))
        table_area.grid(row=3, column=0, sticky="nsew")
        table_area.grid_columnconfigure(0, weight=1)
        table_area.grid_rowconfigure(0, weight=1)
        columns = (
            "no", "name", "pieces", "length", "width", "height", "dimension_unit",
            "actual", "weight_unit", "volume", "volumetric", "per_piece",
        )
        self.tree = ttk.Treeview(table_area, columns=columns, show="headings", height=12,
                                 selectmode="browse")
        headings = {
            "no": ("序号", 54, "center"),
            "name": ("货物名称", 190, "w"),
            "pieces": ("件数", 68, "center"),
            "length": ("长", 75, "e"),
            "width": ("宽", 75, "e"),
            "height": ("高", 75, "e"),
            "dimension_unit": ("尺寸单位", 82, "center"),
            "actual": ("单件实际重量", 120, "e"),
            "weight_unit": ("重量单位", 82, "center"),
            "volume": ("行体积 (m³)", 112, "e"),
            "volumetric": ("行体积重 (kg)", 125, "e"),
            "per_piece": ("逐件计费重量 (kg)", 144, "e"),
        }
        for key, (title, width, anchor) in headings.items():
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=width, anchor=anchor,
                             stretch=(key in ("name", "per_piece")))
        self.tree.tag_configure("even", background="#F7FAFC")
        self.tree.tag_configure("odd", background="#FFFFFF")
        self.tree.tag_configure("placeholder", foreground="#8A99A6", background="#F7FAFC")
        ybar = ttk.Scrollbar(table_area, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(table_area, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", lambda _event: self._edit_line())
        self.tree.bind("<Delete>", lambda _event: self._delete_line())
        self.tree.bind("<Insert>", lambda _event: self._add_line())
        self.root.bind("<Control-Return>", lambda _event: self._calculate())

        summary_frame = ttk.Frame(self.root, padding=(22, 0, 22, 11))
        summary_frame.grid(row=4, column=0, sticky="ew")
        for col in range(6):
            summary_frame.grid_columnconfigure(col, weight=1)
        self.metric_vars = {
            "volume": tk.StringVar(value="—"),
            "actual": tk.StringVar(value="—"),
            "volumetric": tk.StringVar(value="—"),
            "piece": tk.StringVar(value="—"),
            "shipment": tk.StringVar(value="—"),
            "selected": tk.StringVar(value="—"),
        }
        metric_titles = (
            ("volume", "总体积", "m³"),
            ("actual", "总实际重量", "kg"),
            ("volumetric", "总体积重", "kg"),
            ("piece", "逐件口径计费重量", "kg"),
            ("shipment", "整票汇总计费重量", "kg"),
            ("selected", "当前计费重量", "kg"),
        )
        for index, (key, title, unit) in enumerate(metric_titles):
            card = ttk.Frame(summary_frame, style="Card.TFrame", padding=(12, 10))
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 6, 0))
            ttk.Label(card, text=title, style="MetricLabel.TLabel").pack(anchor="w")
            ttk.Label(card, textvariable=self.metric_vars[key], style="MetricValue.TLabel").pack(anchor="w", pady=(5, 0))
            ttk.Label(card, text=unit, style="MetricLabel.TLabel").pack(anchor="w")

        footer = ttk.Frame(self.root, padding=(22, 0, 22, 14))
        footer.grid(row=5, column=0, sticky="ew")
        self.status_var = tk.StringVar(value="尚未录入货物。选择参数后添加货物即可开始。")
        footer.grid_columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(footer, text="快捷键：Insert 新增  ·  Delete 删除  ·  Ctrl+Enter 计算",
                  style="Muted.TLabel").grid(row=0, column=1, sticky="e")

    @staticmethod
    def _unit_combo(
        parent,
        variable: tk.StringVar,
        values: tuple[str, ...],
        column: int,
        row: int = 0,
    ) -> None:
        ttk.Combobox(parent, textvariable=variable, width=7, values=values,
                     state="readonly").grid(row=row, column=column, sticky="w", padx=(0, 8))

    def _sync_settings_fields(self) -> None:
        self.custom_divisor_entry.configure(
            state="normal" if self.divisor_var.get() == "自定义" else "disabled"
        )
        self.custom_rounding_entry.configure(
            state="normal" if self.rounding_var.get() == "自定义" else "disabled"
        )
        self._recalculate_safely()

    def _settings(self) -> CalculationSettings:
        divisor_text = (
            self.custom_divisor_var.get()
            if self.divisor_var.get() == "自定义"
            else self.divisor_var.get()
        )
        try:
            divisor = Decimal(divisor_text.strip())
        except (InvalidOperation, AttributeError):
            raise FreightCalculationError("请填写有效的体积重除数。") from None

        mode = MODE_LABELS.get(self.mode_var.get())
        if mode is None:
            raise FreightCalculationError("请选择有效的计费口径。")

        rounding_label = self.rounding_var.get()
        rounding_map = {
            "不进位": None,
            "0.1 kg": Decimal("0.1"),
            "0.5 kg": Decimal("0.5"),
            "1 kg": Decimal("1"),
        }
        if rounding_label == "自定义":
            try:
                increment = Decimal(self.custom_rounding_var.get().strip())
            except (InvalidOperation, AttributeError):
                raise FreightCalculationError("请填写有效的进位单位。") from None
        elif rounding_label in rounding_map:
            increment = rounding_map[rounding_label]
        else:
            raise FreightCalculationError("请选择有效的进位单位。")
        return CalculationSettings(divisor, mode, increment)  # type: ignore[arg-type]

    def _refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        result = None
        if self.lines:
            try:
                result = calculate(self.lines, self._settings())
            except FreightCalculationError:
                pass
        if not self.lines:
            self.tree.insert(
                "", "end",
                values=("", "尚未添加货物", "", "", "", "", "", "", "", "", "", ""),
                tags=("placeholder",),
            )
            return
        by_index = {id(item.cargo): item for item in result.lines} if result else {}
        for index, line in enumerate(self.lines, start=1):
            calc = by_index.get(id(line))
            self.tree.insert(
                "", "end",
                values=(
                    index, line.name, line.pieces, _display(line.length), _display(line.width),
                    _display(line.height), line.dimension_unit,
                    _display(line.actual_weight_per_piece), line.weight_unit,
                    _display(calc.volume_m3) if calc else "—",
                    _display(calc.volumetric_weight_kg) if calc else "—",
                    _display(calc.per_piece_chargeable_kg) if calc else "—",
                ),
                tags=("even" if index % 2 == 0 else "odd",),
            )

    def _update_summary(self, result: CalculationResult | None = None) -> None:
        result = result or self.last_result
        if not result:
            for value in self.metric_vars.values():
                value.set("—")
            return
        values = {
            "volume": result.total_volume_m3,
            "actual": result.total_actual_weight_kg,
            "volumetric": result.total_volumetric_weight_kg,
            "piece": result.per_piece_chargeable_weight_kg,
            "shipment": result.shipment_total_chargeable_weight_kg,
            "selected": result.selected_chargeable_weight_kg,
        }
        for key, value in values.items():
            self.metric_vars[key].set(_display(value))

    def _recalculate_safely(self) -> None:
        if not hasattr(self, "tree"):
            return
        if not self.lines:
            self.last_result = None
            self._update_summary()
            if hasattr(self, "status_var"):
                self.status_var.set("尚未录入货物。参数为示例设置，请按承运渠道确认。")
            self._refresh_table()
            return
        try:
            self.last_result = calculate(self.lines, self._settings())
        except FreightCalculationError as exc:
            self.last_result = None
            self.status_var.set(str(exc))
            self._update_summary()
            self._refresh_table()
            return
        self._update_summary(self.last_result)
        self.status_var.set(
            f"已按 {self.last_result.settings.volume_divisor} cm³/kg 和"
            f"“{MODE_NAMES[self.last_result.settings.billing_mode]}”规则计算；"
            "以上为计算示例，具体账单以承运渠道确认结果为准。"
        )
        self._refresh_table()

    def _calculate(self) -> None:
        try:
            result = calculate(self.lines, self._settings())
        except FreightCalculationError as exc:
            messagebox.showerror("无法计算", str(exc), parent=self.root)
            return
        self.last_result = result
        self._update_summary(result)
        self._refresh_table()

    def _add_line(self) -> None:
        self._line_dialog(None)

    def _selected_index(self) -> int | None:
        selection = self.tree.selection()
        if not selection:
            return None
        item = self.tree.index(selection[0])
        return item if item < len(self.lines) else None

    def _edit_line(self) -> None:
        index = self._selected_index()
        if index is None:
            messagebox.showinfo("选择货物", "请先选择一行有效货物。", parent=self.root)
            return
        self._line_dialog(index)

    def _delete_line(self) -> None:
        index = self._selected_index()
        if index is None:
            messagebox.showinfo("选择货物", "请先选择一行有效货物。", parent=self.root)
            return
        del self.lines[index]
        self._refresh_table()
        self._recalculate_safely()

    def _line_dialog(self, edit_index: int | None) -> None:
        existing = self.lines[edit_index] if edit_index is not None else None
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑货物" if existing else "新增货物")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(background="#F2F5F8")
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=(22, 18))
        frame.grid(row=0, column=0)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(3, weight=1)
        ttk.Label(
            frame,
            text="按同一包装规格录入货物。件数填写此规格的总箱数。",
            style="Muted.TLabel",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 13))

        values = {
            "name": tk.StringVar(value=existing.name if existing else ""),
            "pieces": tk.StringVar(value=str(existing.pieces) if existing else "1"),
            "length": tk.StringVar(value=_display(existing.length) if existing else ""),
            "width": tk.StringVar(value=_display(existing.width) if existing else ""),
            "height": tk.StringVar(value=_display(existing.height) if existing else ""),
            "dimension_unit": tk.StringVar(
                value=existing.dimension_unit if existing else self.default_dimension_var.get()
            ),
            "actual": tk.StringVar(
                value=_display(existing.actual_weight_per_piece) if existing else ""
            ),
            "weight_unit": tk.StringVar(
                value=existing.weight_unit if existing else self.default_weight_var.get()
            ),
        }
        def make_widget(key: str, kind: str, width: int = 18) -> ttk.Widget:
            if kind == "entry":
                return ttk.Entry(frame, textvariable=values[key], width=width)
            choices = ("cm", "in") if kind == "dimension" else ("kg", "lb")
            return ttk.Combobox(
                frame, textvariable=values[key], values=choices, state="readonly", width=width,
            )

        def add_field(row: int, column: int, label: str, key: str, kind: str,
                      span: int = 1) -> ttk.Widget:
            ttk.Label(frame, text=label, style="Muted.TLabel").grid(
                row=row, column=column, sticky="w", padx=(0, 10), pady=7,
            )
            widget = make_widget(key, kind, 24 if span > 1 else 16)
            widget.grid(row=row, column=column + 1, columnspan=span, sticky="ew",
                        padx=(0, 17 if column == 1 else 0), pady=7)
            return widget

        name_entry = add_field(1, 0, "货物名称", "name", "entry", span=3)
        add_field(2, 0, "件数", "pieces", "entry")
        add_field(2, 2, "尺寸单位", "dimension_unit", "dimension")
        add_field(3, 0, "长度", "length", "entry")
        add_field(3, 2, "宽度", "width", "entry")
        add_field(4, 0, "高度", "height", "entry")
        add_field(4, 2, "单件实际重量", "actual", "entry")
        add_field(5, 0, "重量单位", "weight_unit", "weight")
        ttk.Label(
            frame, text="单件外箱尺寸；实际重量按每件填写。", style="Hint.TLabel",
        ).grid(row=5, column=2, columnspan=2, sticky="w")

        def save() -> None:
            try:
                pieces_text = values["pieces"].get().strip()
                if not pieces_text.isdigit():
                    raise FreightCalculationError("件数必须是正整数。")
                line = CargoLine(
                    name=values["name"].get(),
                    pieces=int(pieces_text),
                    length=Decimal(values["length"].get().strip()),
                    width=Decimal(values["width"].get().strip()),
                    height=Decimal(values["height"].get().strip()),
                    dimension_unit=values["dimension_unit"].get(),  # type: ignore[arg-type]
                    actual_weight_per_piece=Decimal(values["actual"].get().strip()),
                    weight_unit=values["weight_unit"].get(),  # type: ignore[arg-type]
                )
            except (FreightCalculationError, InvalidOperation, ValueError) as exc:
                messagebox.showerror(
                    "输入有误",
                    str(exc) if isinstance(exc, FreightCalculationError) else "尺寸和重量请输入有效数字。",
                    parent=dialog,
                )
                return
            if edit_index is None:
                self.lines.append(line)
            else:
                self.lines[edit_index] = line
            self._refresh_table()
            self._recalculate_safely()
            dialog.destroy()

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=4, sticky="e", pady=(15, 0))
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(button_frame, text="保存", style="Primary.TButton", command=save).pack(side="right")
        dialog.bind("<Return>", lambda _event: save())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        name_entry.focus_set()

    def _import_excel(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="导入货物资料",
            filetypes=(("Excel 工作簿", "*.xlsx"),),
        )
        if not path:
            return
        try:
            lines, settings = read_workbook(path)
            if settings and not messagebox.askyesno(
                "导入计算参数",
                "工作簿包含计算参数。是否同时使用其中的除数、计费口径和进位设置？",
                parent=self.root,
            ):
                settings = None
        except FreightCalculationError as exc:
            messagebox.showerror("导入失败", str(exc), parent=self.root)
            return
        self.lines = lines
        if settings:
            self._set_settings(settings)
        self._refresh_table()
        self._recalculate_safely()
        self.status_var.set(f"已导入 {len(lines)} 行货物资料。请核对单位和计算参数。")

    def _set_settings(self, settings: CalculationSettings) -> None:
        if settings.volume_divisor in (Decimal("5000"), Decimal("6000")):
            self.divisor_var.set(str(int(settings.volume_divisor)))
        else:
            self.divisor_var.set("自定义")
            self.custom_divisor_var.set(_display(settings.volume_divisor))
        self.mode_var.set(MODE_NAMES[settings.billing_mode])
        increments = {
            Decimal("0.1"): "0.1 kg",
            Decimal("0.5"): "0.5 kg",
            Decimal("1"): "1 kg",
        }
        if settings.rounding_increment_kg is None:
            self.rounding_var.set("不进位")
        elif settings.rounding_increment_kg in increments:
            self.rounding_var.set(increments[settings.rounding_increment_kg])
        else:
            self.rounding_var.set("自定义")
            self.custom_rounding_var.set(_display(settings.rounding_increment_kg))
        self._sync_settings_fields()

    def _export_excel(self) -> None:
        try:
            result = calculate(self.lines, self._settings())
        except FreightCalculationError as exc:
            messagebox.showerror("无法导出", str(exc), parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出计算结果",
            defaultextension=".xlsx",
            initialfile="freight-calculation.xlsx",
            filetypes=(("Excel 工作簿", "*.xlsx"),),
        )
        if not path:
            return
        try:
            write_results(path, self.lines, result)
        except (OSError, FreightCalculationError) as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)
            return
        messagebox.showinfo("导出完成", f"结果已保存：\n{Path(path).name}", parent=self.root)

    def _export_template(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出 Excel 导入模板",
            defaultextension=".xlsx",
            initialfile="freight-calculator-template.xlsx",
            filetypes=(("Excel 工作簿", "*.xlsx"),),
        )
        if not path:
            return
        try:
            write_template(path)
        except (OSError, FreightCalculationError) as exc:
            messagebox.showerror("模板导出失败", str(exc), parent=self.root)
            return
        messagebox.showinfo("模板已保存", f"模板已保存：\n{Path(path).name}", parent=self.root)


def main() -> None:
    root = tk.Tk()
    FreightCalculatorApp(root)
    root.mainloop()
