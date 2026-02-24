from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import List, Optional

import ttkbootstrap as ttk
from ttkbootstrap.constants import E, W
from tkinter import messagebox

from invoice_splitter.models import Allocation
from invoice_splitter.rules.common import q2, validate_and_compute_allocations
from invoice_splitter.utils.money import parse_decimal_user_input


@dataclass
class SplitEditorResult:
    mode: str  # "percent" | "amount"
    allocations: List[Allocation]


class SplitEditorDialog(ttk.Toplevel):
    """
    Modal para configurar split personalizado:
    - Modo: percent o amount
    - Concepto por defecto igual al concepto general, editable por línea
    - Muestra 2 estados:
        1) suma/diferencia vs subtotal (siempre)
        2) CC/GL faltantes (siempre)
    - Validación:
        - abs(diff) > 0.01 -> error (no permite aceptar)
        - abs(diff) <= 0.01 -> permitido (se ajusta última línea)
    """

    # Column sizing (pixeles aproximados para alinear etiquetas con entries)
    COL_W = {
        "concept": 300,
        "percent": 80,
        "amount": 110,
        "cc": 90,
        "gl": 130,
        "del": 45,
    }

    def __init__(
        self,
        parent,
        subtotal: Decimal,
        default_concept: str,
        initial_mode: str = "percent",
        initial_allocations: Optional[List[Allocation]] = None,
        title: str = "Configurar split personalizado",
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        self.subtotal = q2(subtotal)
        self.default_concept = default_concept.strip() or "Concepto personalizado"
        self.result: Optional[SplitEditorResult] = None

        self.mode_var = ttk.StringVar(
            value=initial_mode if initial_mode in ("percent", "amount") else "percent"
        )

        # Cada row: dict de vars/widgets, indexado por posición
        self.rows: List[dict] = []

        self._build_ui()

        # Cargar filas iniciales
        if initial_allocations:
            for a in initial_allocations:
                self._add_row(
                    concept=a.concept or self.default_concept,
                    percent=str(a.percent) if a.percent is not None else "",
                    amount=str(a.amount) if a.amount is not None else "",
                    cc=str(a.cc),
                    gl=str(a.gl_account),
                )
        else:
            # Default: 1 fila coherente
            if self.mode_var.get() == "percent":
                self._add_row(self.default_concept, percent="100", amount="", cc="", gl="")
            else:
                self._add_row(
                    self.default_concept, percent="", amount=str(self.subtotal), cc="", gl=""
                )

        self._refresh_mode_ui()
        self._recalc_status()

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        pad = 10
        container = ttk.Frame(self, padding=pad)
        container.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            container, text=f"Subtotal factura: {self.subtotal}", font=("Segoe UI", 11, "bold")
        ).grid(row=0, column=0, columnspan=6, sticky=W, pady=(0, 8))

        # Mode selection
        mode_frame = ttk.Labelframe(container, text="Modo de split", padding=pad)
        mode_frame.grid(row=1, column=0, columnspan=6, sticky="we", pady=(0, 10))

        ttk.Radiobutton(
            mode_frame,
            text="Por porcentaje (%)",
            variable=self.mode_var,
            value="percent",
            command=self._on_mode_change,
        ).grid(row=0, column=0, sticky=W, padx=(0, 20))

        ttk.Radiobutton(
            mode_frame,
            text="Por valor",
            variable=self.mode_var,
            value="amount",
            command=self._on_mode_change,
        ).grid(row=0, column=1, sticky=W)

        # Grid principal (encabezados + filas) en el MISMO contenedor
        self.grid_frame = ttk.Labelframe(container, text="Líneas del split", padding=pad)
        self.grid_frame.grid(row=2, column=0, columnspan=6, sticky="we")

        # Configurar columnas (mismo ancho para headers y rows)
        self.grid_frame.grid_columnconfigure(0, minsize=self.COL_W["concept"])
        self.grid_frame.grid_columnconfigure(1, minsize=self.COL_W["percent"])
        self.grid_frame.grid_columnconfigure(2, minsize=self.COL_W["amount"])
        self.grid_frame.grid_columnconfigure(3, minsize=self.COL_W["cc"])
        self.grid_frame.grid_columnconfigure(4, minsize=self.COL_W["gl"])
        self.grid_frame.grid_columnconfigure(5, minsize=self.COL_W["del"])

        # Encabezados (fila 0)
        ttk.Label(self.grid_frame, text="Concepto", anchor="center").grid(
            row=0, column=0, sticky="ew", padx=2
        )
        self.h_percent = ttk.Label(self.grid_frame, text="%", anchor="center")
        self.h_percent.grid(row=0, column=1, sticky="ew", padx=2)
        self.h_amount = ttk.Label(self.grid_frame, text="Subtotal", anchor="center")
        self.h_amount.grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Label(self.grid_frame, text="CC", anchor="center").grid(
            row=0, column=3, sticky="ew", padx=2
        )
        ttk.Label(self.grid_frame, text="GL", anchor="center").grid(
            row=0, column=4, sticky="ew", padx=2
        )
        ttk.Label(self.grid_frame, text="", anchor="center").grid(
            row=0, column=5, sticky="ew", padx=2
        )

        # Botón agregar
        ttk.Button(
            container, text="➕ Agregar línea", bootstyle="secondary", command=self._add_row_default
        ).grid(row=3, column=0, sticky=W, pady=(10, 0))

        # Estados (2 advertencias)
        self.sum_status = ttk.Label(container, text="", bootstyle="secondary")
        self.sum_status.grid(row=4, column=0, columnspan=6, sticky="we", pady=(10, 0))

        self.fields_status = ttk.Label(container, text="", bootstyle="secondary")
        self.fields_status.grid(row=5, column=0, columnspan=6, sticky="we", pady=(5, 0))

        # Botones aceptar/cancelar
        actions = ttk.Frame(container)
        actions.grid(row=6, column=0, columnspan=6, sticky=E, pady=(12, 0))

        ttk.Button(actions, text="Aceptar", bootstyle="primary", command=self._on_accept).grid(
            row=0, column=0
        )
        ttk.Button(actions, text="Cancelar", bootstyle="secondary", command=self._on_cancel).grid(
            row=0, column=1, padx=(10, 0)
        )

    # ---------------- Rows management ----------------
    def _add_row_default(self) -> None:
        if self.mode_var.get() == "percent":
            self._add_row(self.default_concept, percent="0", amount="", cc="", gl="")
        else:
            self._add_row(self.default_concept, percent="", amount="0", cc="", gl="")
        self._refresh_mode_ui()
        self._recalc_status()

    def _add_row(self, concept: str, percent: str, amount: str, cc: str, gl: str) -> None:
        # row index visual en grid_frame: +1 por encabezados
        grid_row = len(self.rows) + 1

        concept_var = ttk.StringVar(value=concept)
        percent_var = ttk.StringVar(value=percent)
        amount_var = ttk.StringVar(value=amount)
        cc_var = ttk.StringVar(value=cc)
        gl_var = ttk.StringVar(value=gl)

        concept_entry = ttk.Entry(self.grid_frame, textvariable=concept_var)
        concept_entry.grid(row=grid_row, column=0, sticky="ew", padx=2, pady=2)

        percent_entry = ttk.Entry(self.grid_frame, textvariable=percent_var, justify="right")
        percent_entry.grid(row=grid_row, column=1, sticky="ew", padx=2, pady=2)

        amount_entry = ttk.Entry(self.grid_frame, textvariable=amount_var, justify="right")
        amount_entry.grid(row=grid_row, column=2, sticky="ew", padx=2, pady=2)

        cc_entry = ttk.Entry(self.grid_frame, textvariable=cc_var, justify="right")
        cc_entry.grid(row=grid_row, column=3, sticky="ew", padx=2, pady=2)

        gl_entry = ttk.Entry(self.grid_frame, textvariable=gl_var, justify="right")
        gl_entry.grid(row=grid_row, column=4, sticky="ew", padx=2, pady=2)

        del_btn = ttk.Button(
            self.grid_frame,
            text="🗑",
            bootstyle="danger",
            width=3,
            command=lambda i=len(self.rows): self._remove_row(i),
        )
        del_btn.grid(row=grid_row, column=5, sticky="ew", padx=2, pady=2)

        # Recalcular en vivo
        for w in (concept_entry, percent_entry, amount_entry, cc_entry, gl_entry):
            w.bind("<KeyRelease>", lambda _e: self._recalc_status())

        self.rows.append(
            dict(
                concept_var=concept_var,
                percent_var=percent_var,
                amount_var=amount_var,
                cc_var=cc_var,
                gl_var=gl_var,
                concept_entry=concept_entry,
                percent_entry=percent_entry,
                amount_entry=amount_entry,
                cc_entry=cc_entry,
                gl_entry=gl_entry,
                del_btn=del_btn,
            )
        )

    def _remove_row(self, index: int) -> None:
        if len(self.rows) <= 1:
            messagebox.showwarning("Split", "Debe existir al menos una línea.")
            return

        # destruir widgets de esa fila
        r = self.rows[index]
        for key in (
            "concept_entry",
            "percent_entry",
            "amount_entry",
            "cc_entry",
            "gl_entry",
            "del_btn",
        ):
            r[key].destroy()

        self.rows.pop(index)

        # Re-render: recolocar filas restantes en el grid (row = i+1)
        for i, rr in enumerate(self.rows):
            grid_row = i + 1
            rr["concept_entry"].grid_configure(row=grid_row)
            rr["percent_entry"].grid_configure(row=grid_row)
            rr["amount_entry"].grid_configure(row=grid_row)
            rr["cc_entry"].grid_configure(row=grid_row)
            rr["gl_entry"].grid_configure(row=grid_row)
            rr["del_btn"].grid_configure(row=grid_row)
            rr["del_btn"].configure(command=lambda idx=i: self._remove_row(idx))

        self._recalc_status()

    # ---------------- Mode / Status ----------------
    def _on_mode_change(self) -> None:
        # Defaults coherentes si solo hay una fila y está vacío
        if len(self.rows) == 1:
            r = self.rows[0]
            if self.mode_var.get() == "percent" and not r["percent_var"].get().strip():
                r["percent_var"].set("100")
            if self.mode_var.get() == "amount" and not r["amount_var"].get().strip():
                r["amount_var"].set(str(self.subtotal))

        self._refresh_mode_ui()
        self._recalc_status()

    def _refresh_mode_ui(self) -> None:
        mode = self.mode_var.get()
        if mode == "percent":
            self.h_percent.configure(state="normal")
            self.h_amount.configure(state="disabled")
        else:
            self.h_percent.configure(state="disabled")
            self.h_amount.configure(state="normal")

        for r in self.rows:
            if mode == "percent":
                r["percent_entry"].configure(state="normal")
                r["amount_entry"].configure(state="disabled")
            else:
                r["percent_entry"].configure(state="disabled")
                r["amount_entry"].configure(state="normal")

    def _compute_sum_status(self) -> tuple[str, str]:
        """
        Calcula estado de suma/diferencia aunque falten CC/GL.
        Para filas vacías:
          - percent vacío -> 0 (si es única fila -> 100)
          - amount vacío -> 0 (si es única fila -> subtotal)
        """
        mode = self.mode_var.get()
        single = len(self.rows) == 1

        allocs: List[Allocation] = []
        for r in self.rows:
            concept = r["concept_var"].get().strip() or self.default_concept

            if mode == "percent":
                pct_raw = r["percent_var"].get().strip()
                if single and pct_raw == "":
                    pct_raw = "100"
                pct_clean = pct_raw.replace("%", "").strip().replace(",", ".")
                pct = Decimal(pct_clean) if pct_clean else Decimal("0")
                allocs.append(Allocation(concept=concept, cc=0, gl_account=0, percent=pct))
            else:
                amt_raw = r["amount_var"].get().strip()
                if single and amt_raw == "":
                    amt = self.subtotal
                else:
                    amt = parse_decimal_user_input(amt_raw or "0", field_name="Valor de línea")
                allocs.append(Allocation(concept=concept, cc=0, gl_account=0, amount=amt))

        try:
            pairs = validate_and_compute_allocations(self.subtotal, mode, allocs)
            total = q2(sum(amt for _a, amt in pairs))
            diff = q2(self.subtotal - total)
            return (
                f"Suma líneas: {total} | Diferencia vs subtotal: {diff} (ajuste aplicado de 0.01 de ser necesario)",
                "success",
            )
        except Exception as e:
            return (f"⚠️ Suma/porcentajes inválidos: {e}", "warning")

    def _compute_fields_status(self) -> tuple[str, str]:
        missing = 0
        for r in self.rows:
            if not r["cc_var"].get().strip() or not r["gl_var"].get().strip():
                missing += 1

        if missing == 0:
            return ("CC/GL: OK (todas las líneas tienen CC y GL account).", "success")
        return (
            f"⚠️ CC/GL faltantes en {missing} línea(s). Debes completarlos para aceptar.",
            "warning",
        )

    def _recalc_status(self) -> None:
        msg_sum, style_sum = self._compute_sum_status()
        self.sum_status.configure(text=msg_sum, bootstyle=style_sum)

        msg_fields, style_fields = self._compute_fields_status()
        self.fields_status.configure(text=msg_fields, bootstyle=style_fields)

    # ---------------- Collect / Accept ----------------
    def _collect_allocations_strict(self) -> SplitEditorResult:
        mode = self.mode_var.get()
        allocs: List[Allocation] = []
        single = len(self.rows) == 1

        for r in self.rows:
            concept = r["concept_var"].get().strip() or self.default_concept

            cc_raw = r["cc_var"].get().strip()
            gl_raw = r["gl_var"].get().strip()
            if not cc_raw or not gl_raw:
                raise ValueError("Cada línea debe tener CC y GL account para aceptar.")

            cc = int(cc_raw)
            gl = int(gl_raw)

            if mode == "percent":
                pct_raw = r["percent_var"].get().strip()
                if single and pct_raw == "":
                    pct_raw = "100"
                pct_clean = pct_raw.replace("%", "").strip().replace(",", ".")
                pct = Decimal(pct_clean) if pct_clean else Decimal("0")
                allocs.append(Allocation(concept=concept, cc=cc, gl_account=gl, percent=pct))
            else:
                amt_raw = r["amount_var"].get().strip()
                if single and amt_raw == "":
                    amt = self.subtotal
                else:
                    amt = parse_decimal_user_input(amt_raw or "0", field_name="Valor de línea")
                allocs.append(Allocation(concept=concept, cc=cc, gl_account=gl, amount=amt))

        _ = validate_and_compute_allocations(self.subtotal, mode, allocs)
        return SplitEditorResult(mode=mode, allocations=allocs)

    def _on_accept(self) -> None:
        try:
            self.result = self._collect_allocations_strict()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Split inválido", str(e))

    def _on_cancel(self) -> None:
        self.result = None
        self.destroy()

    def show(self) -> Optional[SplitEditorResult]:
        self.wait_window()
        return self.result
