from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from typing import Dict, List, Optional, Any

from click import Path
import ttkbootstrap as ttk
from ttkbootstrap.constants import END, LEFT, W, X
from tkinter import messagebox
from tkinter.ttk import Treeview
from tkinter import filedialog
from pathlib import Path as PathlibPath

from invoice_splitter.config import get_settings
from invoice_splitter.excel.vendors import Vendor, load_vendors_from_table
from invoice_splitter.excel.writer import ExcelWriteError, apply_transaction
from invoice_splitter.models import InvoiceInput, LineItem, Allocation
from invoice_splitter.rules.registry import build_lines
from invoice_splitter.ui.split_editor import SplitEditorDialog
from invoice_splitter.utils.dates import UI_DATE_FORMAT, parse_ui_date, today
from invoice_splitter.utils.money import normalize_bill_number, parse_decimal_user_input, parse_iva
from invoice_splitter.config import get_settings, set_excel_path_user_config


# -----------------------------
# Vendor IDs
# -----------------------------
EIKON_ID = 1255097
AKROS_ID = 1255036
SIPBOX_ID = 1274957

PUNTONET_ID = 1261182
CIRION_ID = 1254926
MOVISTAR_ID = 1260177  # OTECEL (Movistar)

CLARO_ID = 1254902  # CONECEL / Claro

OTRO = "Otro (personalizado)"

# -----------------------------
# Concepts
# -----------------------------
EIKON_CONCEPTS = [
    "Infrastructure cloud (Monthly)",
    "Azure Consumptions (biannual)",
    "Maintenance and support (annual)",
    "Domains (annual)",
    OTRO,
]

AKROS_DEFAULT_CONCEPT = "Printers & Copiers"
SIPBOX_DEFAULT_CONCEPT = "Lenovo ThinkSmartHub + Stem speaker + POE switch"

PUNTONET_DEFAULT_CONCEPT = "40 MBPS"
CIRION_DEFAULT_CONCEPT = "Internet"
MOVISTAR_DEFAULT_CONCEPT = "10 lines DRP (4 lines 35 GB + 6 lines 53 GB)"

# CLARO – por radio
CLARO_SIPTRUNK_DEFAULT_CONCEPT = "Claro Siptrunk"
CLARO_SBC_DEFAULT_CONCEPT = "SBC in cloud"
CLARO_MOBILE_DEFAULT_CONCEPT = "2 lines 50 GB + 3 lines 20 GB"

CLARO_CONCEPTS_BY_TYPE = {
    "siptrunk": [CLARO_SIPTRUNK_DEFAULT_CONCEPT, OTRO],
    "sbc": [CLARO_SBC_DEFAULT_CONCEPT, OTRO],
    "mobile": [CLARO_MOBILE_DEFAULT_CONCEPT, OTRO],
}

# -----------------------------
# Defaults extras
# -----------------------------
DEFAULT_BANDWIDTH = "40"  # CIRION
DEFAULT_PHONE_LINES = "10"  # MOVISTAR

# CLARO Siptrunk
CLARO_SIPTRUNK_DEFAULT_BW = "50"
CLARO_SIPTRUNK_DEFAULT_CHANNELS = "10"

# CLARO SBC
CLARO_SBC_DEFAULT_SIPTRUNK_MBPS = "2"
CLARO_SBC_DEFAULT_LIC_QTY = "14"
CLARO_SBC_DEFAULT_SIPTRUNK_PRICE = "80"
CLARO_SBC_DEFAULT_LIC_PRICE = "266"

# CLARO Mobile
CLARO_MOBILE_DEFAULT_LINES_QTY = "5"


class MainWindow(ttk.Window):
    def __init__(self) -> None:
        super().__init__(themename="flatly")
        self.title("Invoice Splitter")
        self.geometry("1060x760")
        self.minsize(980, 620)

        # --- 1) Asegurar excel_path (file picker si no está configurado) ---
        try:
            self.settings = get_settings()
        except (ValueError, FileNotFoundError) as e:
            # Pedir al usuario el Excel (SharePoint sincronizado aparece como ruta local)
            path = filedialog.askopenfilename(
                title="Selecciona el archivo Excel de Invoice registers",
                filetypes=[("Excel files", "*.xlsx *.xlsm")],
            )
            if not path:
                messagebox.showerror("Configuración requerida", str(e))
                self.destroy()
                return

            set_excel_path_user_config(PathlibPath(path))
            # Reintentar
            try:
                self.settings = get_settings()
            except Exception as e2:
                messagebox.showerror("Error de configuración", str(e2))
                self.destroy()
                return

        self.excel_path = self.settings.excel_path
        self.backup_dir = self.settings.excel_path.parent / "invoice_splitter_backups"
        self.session_backup_path = None

        # --- 2) Cargar vendors desde Vendors_table ya con excel_path válido ---

        self.vendors: List[Vendor] = load_vendors_from_table(
            excel_path=str(self.excel_path),
            sheet_name=self.settings.vendors_sheet,
            table_name=self.settings.vendors_table,
        )

        self.vendor_by_name: Dict[str, Vendor] = {v.vendor_name: v for v in self.vendors}

        # Base vars
        self.vendor_var = ttk.StringVar(value=self.vendors[0].vendor_name if self.vendors else "")
        self.bill_var = ttk.StringVar(value="")
        self.subtotal_var = ttk.StringVar(value="")
        self.iva_var = ttk.StringVar(value=str(self.settings.default_iva))

        # EIKON vars
        self.eikon_concept_var = ttk.StringVar(value=EIKON_CONCEPTS[0])
        self.eikon_custom_concept_var = ttk.StringVar(value="")
        self.eikon_custom_cc_var = ttk.StringVar(value="")
        self.eikon_custom_gl_var = ttk.StringVar(value="")

        # Generic vars
        self.generic_concept_list_var = ttk.StringVar(value="")
        self.generic_custom_concept_var = ttk.StringVar(value="")
        self.generic_cc_var = ttk.StringVar(value="")
        self.generic_gl_var = ttk.StringVar(value="")

        # Extras vars (non-Claro)
        self.bandwidth_var = ttk.StringVar(value=DEFAULT_BANDWIDTH)  # CIRION
        self.phone_lines_var = ttk.StringVar(value=DEFAULT_PHONE_LINES)  # MOVISTAR

        # CLARO UI vars
        self.claro_service_type_var = ttk.StringVar(value="siptrunk")  # siptrunk|sbc|mobile
        self.claro_concept_var = ttk.StringVar(value=CLARO_SIPTRUNK_DEFAULT_CONCEPT)
        self.claro_custom_concept_var = ttk.StringVar(value="")
        self.claro_cc_var = ttk.StringVar(value="")
        self.claro_gl_var = ttk.StringVar(value="")

        # CLARO extras vars
        self.claro_siptrunk_bw_var = ttk.StringVar(value=CLARO_SIPTRUNK_DEFAULT_BW)
        self.claro_siptrunk_channels_var = ttk.StringVar(value=CLARO_SIPTRUNK_DEFAULT_CHANNELS)

        self.claro_sbc_siptrunk_mbps_var = ttk.StringVar(value=CLARO_SBC_DEFAULT_SIPTRUNK_MBPS)
        self.claro_sbc_lic_qty_var = ttk.StringVar(value=CLARO_SBC_DEFAULT_LIC_QTY)
        self.claro_sbc_siptrunk_price_var = ttk.StringVar(value=CLARO_SBC_DEFAULT_SIPTRUNK_PRICE)
        self.claro_sbc_lic_price_var = ttk.StringVar(value=CLARO_SBC_DEFAULT_LIC_PRICE)

        self.claro_mobile_lines_qty_var = ttk.StringVar(value=CLARO_MOBILE_DEFAULT_LINES_QTY)

        # Split state
        self.custom_alloc_mode: Optional[str] = None
        self.custom_allocations: List[Allocation] = []

        # Preview
        self.preview_lines: List[LineItem] = []

        # Tree autosize state (solo resize ventana)
        self._tree_cols: List[str] = []
        self._tree_col_minwidth: int = 50
        self._tree_col_weights: Dict[str, int] = {}
        self._tree_resizing: bool = False

        # Sorting state + indicator
        self._sort_state: Dict[str, bool] = {}
        self._sorted_col: Optional[str] = None
        self._base_headings: Dict[str, str] = {}

        # Soft validation (warnings)
        self._warnings: List[str] = []
        self.warnings_var = ttk.StringVar(value="")

        # --- Resumen de previsualización ---
        self.preview_lines_count_var = ttk.StringVar(value="0")
        self.preview_invoice_subtotal_var = ttk.StringVar(value="0.00")
        self.preview_sum_subtotal_var = ttk.StringVar(value="0.00")
        self.preview_diff_subtotal_var = ttk.StringVar(value="0.00")
        self.preview_sum_iva_var = ttk.StringVar(value="0.00")
        self.preview_sum_total_var = ttk.StringVar(value="0.00")
        self.preview_tables_var = ttk.StringVar(value="0")

        self._build_layout()
        self._apply_vendor_defaults_and_visibility()

    # ---------------- UI ----------------
    def _build_layout(self) -> None:
        pad = 10

        container = ttk.Frame(self, padding=pad)
        container.grid(row=0, column=0, sticky="nsew")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        container.grid_rowconfigure(2, weight=1)  # preview expands
        container.grid_columnconfigure(0, weight=1)

        title = ttk.Label(container, text="Ingreso de factura", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="ew", pady=(0, pad))

        # ---------- Form frame ----------
        form = ttk.Frame(container)
        form.grid(row=1, column=0, sticky="ew")
        form.grid_columnconfigure(0, weight=1)

        row1 = ttk.Frame(form)
        row1.grid(row=0, column=0, sticky="ew", pady=(0, pad))

        ttk.Label(row1, text="Proveedor (Vendor):").pack(side=LEFT, padx=(0, 8))
        self.vendor_combo = ttk.Combobox(
            row1,
            textvariable=self.vendor_var,
            values=[v.vendor_name for v in self.vendors],
            state="readonly",
            width=45,
        )
        self.vendor_combo.pack(side=LEFT, padx=(0, 25))
        self.vendor_combo.bind("<<ComboboxSelected>>", self._on_vendor_changed)

        ttk.Label(row1, text="Fecha:").pack(side=LEFT, padx=(0, 8))
        self.date_entry = ttk.DateEntry(
            row1, dateformat=UI_DATE_FORMAT, startdate=today(), bootstyle="primary", width=14
        )
        self.date_entry.pack(side=LEFT)

        row2 = ttk.Frame(form)
        row2.grid(row=1, column=0, sticky="ew", pady=(0, pad))
        ttk.Label(row2, text="Número de factura (9 dígitos):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(row2, textvariable=self.bill_var, width=20).pack(side=LEFT)

        row3 = ttk.Frame(form)
        row3.grid(row=2, column=0, sticky="ew", pady=(0, pad))
        ttk.Label(row3, text="Subtotal (USD):").pack(side=LEFT, padx=(0, 8))

        self.subtotal_entry = ttk.Entry(row3, textvariable=self.subtotal_var, width=20)
        self.subtotal_entry.pack(side=LEFT, padx=(0, 25))

        # ✅ Al salir del subtotal: normaliza CLARO->SBC y limpia preview
        self.subtotal_entry.bind("<FocusOut>", self._on_subtotal_focus_out)

        ttk.Label(row3, text="IVA (% o decimal):").pack(side=LEFT, padx=(0, 8))
        self.iva_entry = ttk.Entry(row3, textvariable=self.iva_var, width=12)
        self.iva_entry.pack(side=LEFT)

        # ✅ Al salir del IVA: limpia preview
        self.iva_entry.bind("<FocusOut>", self._on_iva_focus_out)

        # ---------- Service/Concept frame ----------
        self.vendor_specific_frame = ttk.Labelframe(form, text="Service/ concept", padding=pad)
        self.vendor_specific_frame.grid(row=3, column=0, sticky="ew", pady=(0, pad))
        self.vendor_specific_frame.grid_columnconfigure(0, weight=1)

        # EIKON block
        self.eikon_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.eikon_block, text="Service/ concept:").pack(side=LEFT, padx=(0, 8))
        self.eikon_combo = ttk.Combobox(
            self.eikon_block,
            textvariable=self.eikon_concept_var,
            values=EIKON_CONCEPTS,
            state="readonly",
            width=40,
        )
        self.eikon_combo.pack(side=LEFT)
        self.eikon_combo.bind("<<ComboboxSelected>>", self._on_eikon_concept_changed)

        self.eikon_custom_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.eikon_custom_block, text="Concepto personalizado:").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Entry(
            self.eikon_custom_block, textvariable=self.eikon_custom_concept_var, width=35
        ).pack(side=LEFT, padx=(0, 25))
        ttk.Label(self.eikon_custom_block, text="CC (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.eikon_custom_block, textvariable=self.eikon_custom_cc_var, width=10).pack(
            side=LEFT, padx=(0, 15)
        )
        ttk.Label(self.eikon_custom_block, text="GL (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.eikon_custom_block, textvariable=self.eikon_custom_gl_var, width=14).pack(
            side=LEFT
        )

        # Generic block
        self.generic_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.generic_block, text="Service/ concept:").pack(side=LEFT, padx=(0, 8))
        self.generic_combo = ttk.Combobox(
            self.generic_block,
            textvariable=self.generic_concept_list_var,
            values=[],
            state="readonly",
            width=40,
        )
        self.generic_combo.pack(side=LEFT)
        self.generic_combo.bind("<<ComboboxSelected>>", self._on_generic_concept_changed)

        self.generic_custom_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.generic_custom_block, text="Concepto personalizado:").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Entry(
            self.generic_custom_block, textvariable=self.generic_custom_concept_var, width=35
        ).pack(side=LEFT, padx=(0, 25))
        ttk.Label(self.generic_custom_block, text="CC (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.generic_custom_block, textvariable=self.generic_cc_var, width=10).pack(
            side=LEFT, padx=(0, 15)
        )
        ttk.Label(self.generic_custom_block, text="GL (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.generic_custom_block, textvariable=self.generic_gl_var, width=14).pack(
            side=LEFT
        )

        # Extras blocks (non-Claro)
        self.bandwidth_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.bandwidth_block, text="Bandwidth (MBPS):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.bandwidth_block, textvariable=self.bandwidth_var, width=8).pack(side=LEFT)

        self.phone_lines_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.phone_lines_block, text="Phone lines quantity:").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.phone_lines_block, textvariable=self.phone_lines_var, width=8).pack(
            side=LEFT
        )

        # --------- CLARO block (radio + concept + extras) ----------
        self.claro_block = ttk.Frame(self.vendor_specific_frame)

        rb_frame = ttk.Frame(self.claro_block)
        rb_frame.pack(fill=X, pady=(0, 8))

        ttk.Label(rb_frame, text="Servicio Claro:").pack(side=LEFT, padx=(0, 10))
        ttk.Radiobutton(
            rb_frame,
            text="Siptrunk",
            variable=self.claro_service_type_var,
            value="siptrunk",
            command=self._on_claro_service_changed,
        ).pack(side=LEFT, padx=(0, 10))
        ttk.Radiobutton(
            rb_frame,
            text="SBC",
            variable=self.claro_service_type_var,
            value="sbc",
            command=self._on_claro_service_changed,
        ).pack(side=LEFT, padx=(0, 10))
        ttk.Radiobutton(
            rb_frame,
            text="Mobile",
            variable=self.claro_service_type_var,
            value="mobile",
            command=self._on_claro_service_changed,
        ).pack(side=LEFT)

        concept_frame = ttk.Frame(self.claro_block)
        concept_frame.pack(fill=X)

        ttk.Label(concept_frame, text="Service/ concept:").pack(side=LEFT, padx=(0, 8))
        self.claro_concept_combo = ttk.Combobox(
            concept_frame,
            textvariable=self.claro_concept_var,
            values=CLARO_CONCEPTS_BY_TYPE["siptrunk"],
            state="readonly",
            width=40,
        )
        self.claro_concept_combo.pack(side=LEFT)
        self.claro_concept_combo.bind("<<ComboboxSelected>>", self._on_claro_concept_changed)

        self.claro_custom_block = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.claro_custom_block, text="Concepto personalizado:").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Entry(
            self.claro_custom_block, textvariable=self.claro_custom_concept_var, width=35
        ).pack(side=LEFT, padx=(0, 25))
        ttk.Label(self.claro_custom_block, text="CC (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.claro_custom_block, textvariable=self.claro_cc_var, width=10).pack(
            side=LEFT, padx=(0, 15)
        )
        ttk.Label(self.claro_custom_block, text="GL (1 línea):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.claro_custom_block, textvariable=self.claro_gl_var, width=14).pack(side=LEFT)

        self.claro_siptrunk_extras = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.claro_siptrunk_extras, text="Bandwidth (MBPS):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(
            self.claro_siptrunk_extras, textvariable=self.claro_siptrunk_bw_var, width=8
        ).pack(side=LEFT, padx=(0, 20))
        ttk.Label(self.claro_siptrunk_extras, text="Troncal SIP (channels):").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Entry(
            self.claro_siptrunk_extras, textvariable=self.claro_siptrunk_channels_var, width=8
        ).pack(side=LEFT)

        self.claro_sbc_extras = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.claro_sbc_extras, text="Siptrunk (MBPS):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(
            self.claro_sbc_extras, textvariable=self.claro_sbc_siptrunk_mbps_var, width=6
        ).pack(side=LEFT, padx=(0, 15))
        ttk.Label(self.claro_sbc_extras, text="Licences (Qty):").pack(side=LEFT, padx=(0, 8))
        ttk.Entry(self.claro_sbc_extras, textvariable=self.claro_sbc_lic_qty_var, width=6).pack(
            side=LEFT, padx=(0, 20)
        )

        ttk.Label(self.claro_sbc_extras, text="Siptrunk price:").pack(side=LEFT, padx=(0, 8))
        self.claro_sbc_siptrunk_price_entry = ttk.Entry(
            self.claro_sbc_extras, textvariable=self.claro_sbc_siptrunk_price_var, width=8
        )
        self.claro_sbc_siptrunk_price_entry.pack(side=LEFT, padx=(0, 15))
        self.claro_sbc_siptrunk_price_entry.bind("<FocusOut>", self._on_claro_sbc_price_focus_out)

        ttk.Label(self.claro_sbc_extras, text="Licences price:").pack(side=LEFT, padx=(0, 8))
        self.claro_sbc_lic_price_entry = ttk.Entry(
            self.claro_sbc_extras, textvariable=self.claro_sbc_lic_price_var, width=8
        )
        self.claro_sbc_lic_price_entry.pack(side=LEFT)
        self.claro_sbc_lic_price_entry.bind("<FocusOut>", self._on_claro_sbc_price_focus_out)

        self.claro_mobile_extras = ttk.Frame(self.vendor_specific_frame)
        ttk.Label(self.claro_mobile_extras, text="Phone lines quantity:").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Entry(
            self.claro_mobile_extras, textvariable=self.claro_mobile_lines_qty_var, width=8
        ).pack(side=LEFT)

        # Split controls (common)
        self.split_btn = ttk.Button(
            self.vendor_specific_frame,
            text="Configurar split personalizado…",
            bootstyle="info",
            command=self.on_open_split_dialog,
        )
        self.split_status = ttk.Label(self.vendor_specific_frame, text="", bootstyle="secondary")

        # ✅ Banner de advertencias (validación suave)
        self.warnings_label = ttk.Label(
            form,
            textvariable=self.warnings_var,
            bootstyle="warning",
            wraplength=950,
            justify="left",
        )
        # debajo del vendor_specific_frame (row 3) => row 4
        self.warnings_label.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        self.warnings_label.grid_remove()

        # ---------- Preview wrapper ----------
        preview_wrapper = ttk.Frame(container)
        preview_wrapper.grid(row=2, column=0, sticky="nsew")
        preview_wrapper.grid_rowconfigure(1, weight=1)
        preview_wrapper.grid_columnconfigure(0, weight=1)

        ttk.Label(
            preview_wrapper,
            text="Previsualización (líneas a insertar):",
            font=("Segoe UI", 11, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))

        preview_frame = ttk.Frame(preview_wrapper)
        preview_frame.grid(row=1, column=0, sticky="nsew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        self.v_scroll = ttk.Scrollbar(preview_frame, orient="vertical")
        self.h_scroll = ttk.Scrollbar(preview_frame, orient="horizontal")

        # ---------- Preview Summary ----------
        # (Debajo del Treeview, dentro del preview_wrapper)
        preview_wrapper.grid_rowconfigure(2, weight=0)

        summary_frame = ttk.Frame(preview_wrapper)
        summary_frame.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        summary_frame.grid_columnconfigure(0, weight=1)

        # Primera fila de resumen
        row_a = ttk.Frame(summary_frame)
        row_a.pack(fill="x")

        ttk.Label(row_a, text="Líneas:", font=("Segoe UI", 10, "bold")).pack(side=LEFT, padx=(0, 6))
        ttk.Label(row_a, textvariable=self.preview_lines_count_var).pack(side=LEFT, padx=(0, 16))

        ttk.Label(row_a, text="Tablas:", font=("Segoe UI", 10, "bold")).pack(side=LEFT, padx=(0, 6))
        ttk.Label(row_a, textvariable=self.preview_tables_var).pack(side=LEFT, padx=(0, 16))

        ttk.Label(row_a, text="Subtotal factura:", font=("Segoe UI", 10, "bold")).pack(
            side=LEFT, padx=(0, 6)
        )
        ttk.Label(row_a, textvariable=self.preview_invoice_subtotal_var).pack(side=LEFT)

        # Segunda fila de resumen
        row_b = ttk.Frame(summary_frame)
        row_b.pack(fill="x", pady=(4, 0))

        ttk.Label(row_b, text="Σ Subtotal assigned:", font=("Segoe UI", 10, "bold")).pack(
            side=LEFT, padx=(0, 6)
        )
        ttk.Label(row_b, textvariable=self.preview_sum_subtotal_var).pack(side=LEFT, padx=(0, 16))

        ttk.Label(row_b, text="Diferencia:", font=("Segoe UI", 10, "bold")).pack(
            side=LEFT, padx=(0, 6)
        )
        # diferencia en estilo (cambiaremos el color cuando no sea cero)
        self.preview_diff_label = ttk.Label(row_b, textvariable=self.preview_diff_subtotal_var)
        self.preview_diff_label.pack(side=LEFT, padx=(0, 16))

        ttk.Label(row_b, text="Σ IVA assigned:", font=("Segoe UI", 10, "bold")).pack(
            side=LEFT, padx=(0, 6)
        )
        ttk.Label(row_b, textvariable=self.preview_sum_iva_var).pack(side=LEFT, padx=(0, 16))

        ttk.Label(row_b, text="Σ Total assigned:", font=("Segoe UI", 10, "bold")).pack(
            side=LEFT, padx=(0, 6)
        )
        ttk.Label(row_b, textvariable=self.preview_sum_total_var).pack(side=LEFT)

        style = ttk.Style()
        style.configure("Preview.Treeview", rowheight=22)
        style.configure("Preview.Treeview.Heading", anchor="e")

        self._tree_cols = [
            "table",
            "date",
            "bill",
            "vendor",
            "concept",
            "cc",
            "gl",
            "sub",
            "iva",
            "iva_amt",
            "total",
        ]

        self.tree = Treeview(
            preview_frame,
            columns=tuple(self._tree_cols),
            show="headings",
            height=10,
            style="Preview.Treeview",
            yscrollcommand=self.v_scroll.set,
            xscrollcommand=self.h_scroll.set,
        )
        self.v_scroll.config(command=self.tree.yview)
        self.h_scroll.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        self._base_headings = {
            "table": "Tabla",
            "date": "Date",
            "bill": "Bill number",
            "vendor": "Vendor",
            "concept": "Service/ concept",
            "cc": "CC",
            "gl": "GL account",
            "sub": "Subtotal assigned",
            "iva": "% IVA",
            "iva_amt": "IVA assigned",
            "total": "Total assigned",
        }
        for c in self._tree_cols:
            self._set_tree_heading(c, self._base_headings[c])

        base_widths = {
            "table": 110,
            "date": 110,
            "bill": 120,
            "vendor": 200,
            "concept": 240,
            "cc": 100,
            "gl": 140,
            "sub": 140,
            "iva": 90,
            "iva_amt": 140,
            "total": 140,
        }
        self._tree_col_weights = {c: base_widths.get(c, 120) for c in self._tree_cols}
        for c in self._tree_cols:
            self.tree.column(
                c,
                anchor="e",
                width=base_widths.get(c, 120),
                stretch=True,
                minwidth=self._tree_col_minwidth,
            )

        self.tree.bind("<Configure>", self._on_tree_configure)
        self.tree.bind("<ButtonPress-1>", self._block_manual_tree_resize, add="+")

        # ---------- Buttons frame ----------
        buttons = ttk.Frame(container)
        buttons.grid(row=3, column=0, sticky="ew", pady=(pad, 0))
        buttons.grid_columnconfigure(0, weight=1)

        ttk.Button(
            buttons, text="Previsualizar", bootstyle="success", command=self.on_preview
        ).pack(side=LEFT)
        self.save_btn = ttk.Button(
            buttons, text="Guardar en Excel", bootstyle="primary", command=self.on_save
        )
        self.save_btn.pack(side=LEFT, padx=(10, 0))
        self.save_btn.configure(state="disabled")
        ttk.Button(buttons, text="Limpiar", bootstyle="secondary", command=self.on_clear).pack(
            side=LEFT, padx=(10, 0)
        )

    # ---------------- Soft validation helpers ----------------
    def _clear_warnings(self) -> None:
        self._warnings = []
        self.warnings_var.set("")
        self.warnings_label.grid_remove()

    def _add_warning(self, msg: str) -> None:
        self._warnings.append(msg)

    def _show_warnings(self) -> None:
        if not self._warnings:
            self.warnings_var.set("")
            self.warnings_label.grid_remove()
            return
        text = "⚠️ Advertencias (se aplicaron correcciones automáticas):\n- " + "\n- ".join(
            self._warnings
        )
        self.warnings_var.set(text)
        self.warnings_label.grid()

    def _soft_int_set(
        self,
        var: ttk.StringVar,
        *,
        field: str,
        default: int,
        min_value: int = 0,
        allow_zero: bool = True,
    ) -> int:
        s = (var.get() or "").strip()
        if not s:
            self._add_warning(f"{field} estaba vacío, se usó default={default}.")
            var.set(str(default))
            return default
        try:
            v = int(s)
        except Exception:
            self._add_warning(f"{field}='{s}' no es entero válido, se usó default={default}.")
            var.set(str(default))
            return default

        if not allow_zero and v == 0:
            self._add_warning(f"{field}=0 no permitido, se usó default={default}.")
            var.set(str(default))
            return default

        if v < min_value:
            self._add_warning(f"{field}={v} menor a {min_value}, se usó default={default}.")
            var.set(str(default))
            return default

        # si es válido, lo normalizamos (quita espacios)
        var.set(str(v))
        return v

    def _soft_decimal(
        self,
        value_str: str,
        *,
        field: str,
        default: Decimal,
    ) -> Decimal:
        s = (value_str or "").strip()
        if not s:
            self._add_warning(f"{field} estaba vacío, se usó default={default}.")
            return default
        try:
            return parse_decimal_user_input(s, field_name=field)
        except Exception:
            self._add_warning(f"{field}='{s}' inválido, se usó default={default}.")
            return default

    def _apply_decimal_to_var(self, var: ttk.StringVar, value: Decimal) -> None:
        var.set(f"{value:.2f}")

    # ---------------- UX bidireccional CLARO->SBC ----------------
    def _is_claro_sbc_context(self) -> bool:
        v = self._selected_vendor()
        return bool(
            v and v.vendor_id == CLARO_ID and self.claro_service_type_var.get().strip() == "sbc"
        )

    def _normalize_prices_to_subtotal_sign_soft(self) -> None:
        """
        Si CLARO->SBC:
        - Subtotal < 0 => prices negativos
        - Subtotal >= 0 => prices positivos
        Suave: si price inválido, usa default y avisa (se verá en warnings).
        """
        if not self._is_claro_sbc_context():
            return

        subtotal_raw = (self.subtotal_var.get() or "").strip()
        if not subtotal_raw:
            return
        try:
            subtotal = parse_decimal_user_input(subtotal_raw, field_name="Subtotal")
        except Exception:
            return

        sign = Decimal("-1") if subtotal < 0 else Decimal("1")

        sip_default = Decimal(str(CLARO_SBC_DEFAULT_SIPTRUNK_PRICE))
        sip_dec = self._soft_decimal(
            self.claro_sbc_siptrunk_price_var.get(), field="Siptrunk price", default=sip_default
        )
        sip_dec = abs(sip_dec) * sign
        self._apply_decimal_to_var(self.claro_sbc_siptrunk_price_var, sip_dec)

        lic_default = Decimal(str(CLARO_SBC_DEFAULT_LIC_PRICE))
        lic_dec = self._soft_decimal(
            self.claro_sbc_lic_price_var.get(), field="Licences price", default=lic_default
        )
        lic_dec = abs(lic_dec) * sign
        self._apply_decimal_to_var(self.claro_sbc_lic_price_var, lic_dec)

    # ---------------- Claro handlers ----------------
    def _on_claro_service_changed(self) -> None:
        self._clear_preview_state(clear_warnings=True)
        self._reset_split()
        self._refresh_claro_concepts()
        self._apply_vendor_defaults_and_visibility()
        # normaliza si cae en SBC
        self._normalize_prices_to_subtotal_sign_soft()
        self._show_warnings()

    def _on_claro_concept_changed(self, _event=None) -> None:
        self._clear_preview_state(clear_warnings=True)
        self._reset_split()
        self._apply_vendor_defaults_and_visibility()

    def _refresh_claro_concepts(self) -> None:
        st = self.claro_service_type_var.get()
        values = CLARO_CONCEPTS_BY_TYPE.get(st, [OTRO])
        self.claro_concept_combo.configure(values=values)
        cur = self.claro_concept_var.get().strip()
        if cur not in values:
            self.claro_concept_var.set(values[0])
            self.claro_concept_combo.set(values[0])
        else:
            self.claro_concept_combo.set(cur)
        self.claro_concept_combo.update_idletasks()

    # ---------------- Heading helper (with sorting) ----------------
    def _set_tree_heading(self, col: str, text: str) -> None:
        self.tree.heading(
            col,
            text=text,
            anchor="e",
            command=lambda c=col: self._sort_tree_by_column(c),
        )

    # ---------------- Disable manual resize only ----------------
    def _block_manual_tree_resize(self, event):
        try:
            region = self.tree.identify_region(event.x, event.y)
            if region == "separator":
                return "break"
        except Exception:
            pass
        return None

    # ---------------- Tree autosize (window resize only) ----------------
    def _on_tree_configure(self, _event=None) -> None:
        if self._tree_resizing:
            return
        w = self.tree.winfo_width()
        if w <= 50:
            return
        try:
            self._tree_resizing = True
            self._autosize_tree_columns(w)
        finally:
            self._tree_resizing = False

    def _autosize_tree_columns(self, total_width: int) -> None:
        padding = 18
        available = max(50, total_width - padding)
        weights_sum = sum(self._tree_col_weights.values())
        if weights_sum <= 0:
            return

        widths: Dict[str, int] = {}
        for col in self._tree_cols:
            w = int(available * (self._tree_col_weights[col] / weights_sum))
            widths[col] = max(self._tree_col_minwidth, w)

        used = sum(widths.values())
        diff = available - used
        if diff != 0:
            last = self._tree_cols[-1]
            widths[last] = max(self._tree_col_minwidth, widths[last] + diff)

        for col in self._tree_cols:
            self.tree.column(col, width=widths[col])

    # ---------------- Sorting + indicator ▲/▼ ----------------
    def _sort_tree_by_column(self, col: str) -> None:
        ascending = self._sort_state.get(col, True)
        self._sort_state[col] = not ascending

        items = list(self.tree.get_children(""))
        if items:

            def sort_key(item_id: str):
                values = self.tree.item(item_id, "values")
                idx = self._tree_cols.index(col)
                raw = values[idx] if idx < len(values) else ""
                return self._coerce_sort_value(raw, col)

            items.sort(key=sort_key, reverse=not ascending)
            for i, item in enumerate(items):
                self.tree.move(item, "", i)

        self._update_sort_indicator(col, ascending)

    def _update_sort_indicator(self, col: str, ascending: bool) -> None:
        arrow = " ▲" if ascending else " ▼"
        if self._sorted_col and self._sorted_col != col:
            base_prev = self._base_headings.get(self._sorted_col, self._sorted_col)
            self._set_tree_heading(self._sorted_col, base_prev)
        base_current = self._base_headings.get(col, col)
        self._set_tree_heading(col, base_current + arrow)
        self._sorted_col = col

    def _coerce_sort_value(self, raw: Any, col: str):
        s = "" if raw is None else str(raw).strip()
        if s == "":
            return (1, "")
        numeric_cols = {"cc", "gl", "sub", "iva", "iva_amt", "total"}
        if col in numeric_cols or col == "bill":
            return (0, self._to_float_safe(s))
        if col == "date":
            dt = self._to_date_safe(s)
            return (0, dt) if dt else (0, s)
        return (0, s.lower())

    def _to_float_safe(self, s: str) -> float:
        s2 = s.replace("%", "").replace(" ", "")
        if "," in s2 and "." in s2:
            if s2.rfind(",") > s2.rfind("."):
                s2 = s2.replace(".", "").replace(",", ".")
            else:
                s2 = s2.replace(",", "")
        else:
            s2 = s2.replace(",", ".")
        try:
            return float(s2)
        except ValueError:
            return 0.0

    def _to_date_safe(self, s: str):
        try:
            return datetime.fromisoformat(s)
        except ValueError:
            return None

    # ---------------- Vendor helpers ----------------
    def _selected_vendor(self) -> Optional[Vendor]:
        return self.vendor_by_name.get(self.vendor_var.get().strip())

    def _reset_split(self) -> None:
        self.custom_alloc_mode = None
        self.custom_allocations = []
        self.split_status.configure(text="", bootstyle="secondary")

    def _update_split_status(self) -> None:
        if self.custom_alloc_mode and self.custom_allocations:
            self.split_status.configure(
                text=f"Split configurado: modo={self.custom_alloc_mode} | líneas={len(self.custom_allocations)}",
                bootstyle="success",
            )
        else:
            self.split_status.configure(text="", bootstyle="secondary")

    def _first_alloc_concept(self) -> str:
        for a in self.custom_allocations:
            if a.concept and str(a.concept).strip():
                return str(a.concept).strip()
        return ""

    def _current_general_concept(self) -> str:
        vendor = self._selected_vendor()
        if not vendor:
            return self._first_alloc_concept() or "Concepto personalizado"

        if vendor.vendor_id == EIKON_ID:
            if self.eikon_concept_var.get() == OTRO:
                txt = self.eikon_custom_concept_var.get().strip()
                return txt or self._first_alloc_concept() or "Concepto personalizado"
            return self.eikon_concept_var.get().strip()

        if vendor.vendor_id == CLARO_ID:
            if self.claro_concept_var.get() == OTRO:
                txt = self.claro_custom_concept_var.get().strip()
                return txt or self._first_alloc_concept() or "Concepto personalizado"
            return self.claro_concept_var.get().strip()

        if self.generic_concept_list_var.get() == OTRO:
            txt = self.generic_custom_concept_var.get().strip()
            return txt or self._first_alloc_concept() or "Concepto personalizado"
        return self.generic_concept_list_var.get().strip() or "Concepto personalizado"

    def _is_custom_context(self) -> bool:
        vendor = self._selected_vendor()
        if not vendor:
            return False
        if vendor.vendor_id == EIKON_ID:
            return self.eikon_concept_var.get() == OTRO
        if vendor.vendor_id == CLARO_ID:
            return self.claro_concept_var.get() == OTRO
        return self.generic_concept_list_var.get() == OTRO

    # ---------------- Vendor events ----------------
    def _on_vendor_changed(self, _event=None) -> None:
        # 1) limpiar preview/resumen del vendor anterior
        self._clear_preview_state(clear_warnings=True)

        # 2) reset split y refrescar UI del nuevo vendor
        self._reset_split()
        self._apply_vendor_defaults_and_visibility()

        # 3) normalizaciones propias (si aplica)
        self._normalize_prices_to_subtotal_sign_soft()

    def _on_eikon_concept_changed(self, _event=None) -> None:
        self._clear_preview_state(clear_warnings=True)
        self._reset_split()
        self._apply_vendor_defaults_and_visibility()

    def _on_generic_concept_changed(self, _event=None) -> None:
        self._clear_preview_state(clear_warnings=True)
        self._reset_split()
        self._apply_vendor_defaults_and_visibility()

    def _set_generic_values_and_keep_selection(self, values: List[str], default_value: str) -> None:
        current = self.generic_concept_list_var.get().strip()
        self.generic_combo.configure(values=values)
        if not current or current not in values:
            self.generic_concept_list_var.set(default_value)
            self.generic_combo.set(default_value)
        else:
            self.generic_combo.set(current)
        self.generic_combo.update_idletasks()

    def _apply_vendor_defaults_and_visibility(self) -> None:
        # Hide all blocks
        self.eikon_block.pack_forget()
        self.eikon_custom_block.pack_forget()

        self.generic_block.pack_forget()
        self.generic_custom_block.pack_forget()
        self.bandwidth_block.pack_forget()
        self.phone_lines_block.pack_forget()

        self.claro_block.pack_forget()
        self.claro_custom_block.pack_forget()
        self.claro_siptrunk_extras.pack_forget()
        self.claro_sbc_extras.pack_forget()
        self.claro_mobile_extras.pack_forget()

        self.split_btn.pack_forget()
        self.split_status.pack_forget()

        vendor = self._selected_vendor()
        if not vendor:
            return

        if vendor.vendor_id == EIKON_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (EIKON)")
            self.eikon_block.pack(fill=X)
            if self.eikon_concept_var.get() == OTRO:
                self.eikon_custom_block.pack(fill=X, pady=(8, 0))
                self.split_btn.pack(fill=X, pady=(10, 0))
                self.split_status.pack(fill=X, pady=(5, 0))
                self._update_split_status()
            return

        if vendor.vendor_id == CLARO_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (CLARO)")
            self.claro_block.pack(fill=X)
            self._refresh_claro_concepts()

            st = self.claro_service_type_var.get()
            if st == "siptrunk":
                self.claro_siptrunk_extras.pack(fill=X, pady=(8, 0))
            elif st == "sbc":
                self.claro_sbc_extras.pack(fill=X, pady=(8, 0))
            else:
                self.claro_mobile_extras.pack(fill=X, pady=(8, 0))

            if self.claro_concept_var.get() == OTRO:
                self.claro_custom_block.pack(fill=X, pady=(8, 0))
                self.split_btn.pack(fill=X, pady=(10, 0))
                self.split_status.pack(fill=X, pady=(5, 0))
                self._update_split_status()
            return

        # Generic vendors
        if vendor.vendor_id == AKROS_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (AKROS)")
            self._set_generic_values_and_keep_selection(
                [AKROS_DEFAULT_CONCEPT, OTRO], AKROS_DEFAULT_CONCEPT
            )
            self.generic_block.pack(fill=X)

        elif vendor.vendor_id == SIPBOX_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (SIPBOX)")
            self._set_generic_values_and_keep_selection(
                [SIPBOX_DEFAULT_CONCEPT, OTRO], SIPBOX_DEFAULT_CONCEPT
            )
            self.generic_block.pack(fill=X)

        elif vendor.vendor_id == PUNTONET_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (PUNTONET)")
            self._set_generic_values_and_keep_selection(
                [PUNTONET_DEFAULT_CONCEPT, OTRO], PUNTONET_DEFAULT_CONCEPT
            )
            self.generic_block.pack(fill=X)

        elif vendor.vendor_id == CIRION_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (CIRION)")
            self._set_generic_values_and_keep_selection(
                [CIRION_DEFAULT_CONCEPT, OTRO], CIRION_DEFAULT_CONCEPT
            )
            self.generic_block.pack(fill=X)
            self.bandwidth_block.pack(fill=X, pady=(8, 0))

        elif vendor.vendor_id == MOVISTAR_ID:
            self.vendor_specific_frame.configure(text="Service/ concept (MOVISTAR/OTECEL)")
            self._set_generic_values_and_keep_selection(
                [MOVISTAR_DEFAULT_CONCEPT, OTRO], MOVISTAR_DEFAULT_CONCEPT
            )
            self.generic_block.pack(fill=X)
            self.phone_lines_block.pack(fill=X, pady=(8, 0))

        else:
            self.vendor_specific_frame.configure(text="Service/ concept")
            return

        if self.generic_concept_list_var.get() == OTRO:
            self.generic_custom_block.pack(fill=X, pady=(8, 0))
            self.split_btn.pack(fill=X, pady=(10, 0))
            self.split_status.pack(fill=X, pady=(5, 0))
            self._update_split_status()

    # ---------------- Split dialog ----------------
    def on_open_split_dialog(self) -> None:
        try:
            if not self._is_custom_context():
                messagebox.showinfo(
                    "Split", "El split personalizado solo aplica en 'Otro (personalizado)'."
                )
                return

            subtotal = parse_decimal_user_input(self.subtotal_var.get(), field_name="Subtotal")
            default_concept = self._current_general_concept()

            dialog = SplitEditorDialog(
                parent=self,
                subtotal=subtotal,
                default_concept=default_concept,
                initial_mode=self.custom_alloc_mode or "percent",
                initial_allocations=self.custom_allocations if self.custom_allocations else None,
            )
            res = dialog.show()
            if not res:
                return

            self.custom_alloc_mode = res.mode
            self.custom_allocations = res.allocations
            self._update_split_status()

            try:
                self.on_preview()
            except Exception:
                messagebox.showinfo(
                    "Split cargado", "Split configurado. Completa campos y presiona Previsualizar."
                )

        except Exception as e:
            self._reset_preview_summary()
            messagebox.showerror("Split", str(e))

    # ---------------- Actions ----------------
    def on_clear(self) -> None:
        self.bill_var.set("")
        self.subtotal_var.set("")
        self.iva_var.set(str(self.settings.default_iva))

        self.eikon_concept_var.set(EIKON_CONCEPTS[0])
        self.eikon_custom_concept_var.set("")
        self.eikon_custom_cc_var.set("")
        self.eikon_custom_gl_var.set("")

        self.generic_concept_list_var.set("")
        self.generic_custom_concept_var.set("")
        self.generic_cc_var.set("")
        self.generic_gl_var.set("")

        self.bandwidth_var.set(DEFAULT_BANDWIDTH)
        self.phone_lines_var.set(DEFAULT_PHONE_LINES)

        # Claro reset
        self.claro_service_type_var.set("siptrunk")
        self.claro_concept_var.set(CLARO_SIPTRUNK_DEFAULT_CONCEPT)
        self.claro_custom_concept_var.set("")
        self.claro_cc_var.set("")
        self.claro_gl_var.set("")
        self.claro_siptrunk_bw_var.set(CLARO_SIPTRUNK_DEFAULT_BW)
        self.claro_siptrunk_channels_var.set(CLARO_SIPTRUNK_DEFAULT_CHANNELS)
        self.claro_sbc_siptrunk_mbps_var.set(CLARO_SBC_DEFAULT_SIPTRUNK_MBPS)
        self.claro_sbc_lic_qty_var.set(CLARO_SBC_DEFAULT_LIC_QTY)
        self.claro_sbc_siptrunk_price_var.set(CLARO_SBC_DEFAULT_SIPTRUNK_PRICE)
        self.claro_sbc_lic_price_var.set(CLARO_SBC_DEFAULT_LIC_PRICE)
        self.claro_mobile_lines_qty_var.set(CLARO_MOBILE_DEFAULT_LINES_QTY)

        self._reset_split()
        self.preview_lines = []
        self.save_btn.configure(state="disabled")
        self._clear_tree()
        self._apply_vendor_defaults_and_visibility()

        self._clear_warnings()

        # reset sort indicators
        self._sorted_col = None
        self._sort_state = {}
        for c in self._tree_cols:
            self._set_tree_heading(c, self._base_headings[c])

        self._reset_preview_summary()

    def _on_claro_sbc_price_focus_out(self, _event=None) -> None:
        """
        Al salir de un campo de precio en CLARO->SBC:
        - normaliza los precios según el signo del subtotal
        - invalida preview para forzar nueva previsualización
        """
        self._normalize_prices_to_subtotal_sign_soft()
        self._clear_preview_state(clear_warnings=True)

    def _on_subtotal_focus_out(self, _event=None) -> None:
        """
        Al salir del subtotal:
        - normaliza prices en CLARO->SBC (si aplica)
        - invalida preview para forzar nueva previsualización
        """
        self._normalize_prices_to_subtotal_sign_soft()
        self._clear_preview_state(clear_warnings=True)

    def _on_iva_focus_out(self, _event=None) -> None:
        """
        Al salir del IVA:
        - invalida preview para forzar nueva previsualización
        """
        self._clear_preview_state(clear_warnings=True)

    def on_preview(self) -> None:
        try:
            self._clear_warnings()
            self._normalize_prices_to_subtotal_sign_soft()

            vendor = self._selected_vendor()
            if not vendor:
                raise ValueError("Debes seleccionar un proveedor válido.")

            bill = normalize_bill_number(self.bill_var.get(), field_name="Número de factura")
            subtotal = parse_decimal_user_input(self.subtotal_var.get(), field_name="Subtotal")
            iva_rate = parse_iva(self.iva_var.get())
            invoice_date = parse_ui_date(self.date_entry.entry.get().strip())

            invoice = InvoiceInput(
                invoice_date=invoice_date,
                vendor_id=vendor.vendor_id,
                vendor_name=vendor.vendor_name,
                bill_number=bill,
                subtotal=subtotal,
                iva_rate=iva_rate,
            )

            has_split = bool(self.custom_alloc_mode and self.custom_allocations)
            if has_split:
                invoice.alloc_mode = self.custom_alloc_mode
                invoice.allocations = self.custom_allocations

            # ---- CLARO ----
            if vendor.vendor_id == CLARO_ID:
                st = self.claro_service_type_var.get().strip()
                invoice.service_type = st

                concept_sel = self.claro_concept_var.get().strip()
                invoice.service_concept = concept_sel

                if st == "siptrunk":
                    invoice.extras["bandwidth_mbps"] = self._soft_int_set(
                        self.claro_siptrunk_bw_var,
                        field="CLARO Siptrunk - Bandwidth (MBPS)",
                        default=int(CLARO_SIPTRUNK_DEFAULT_BW),
                        min_value=1,
                        allow_zero=False,
                    )
                    invoice.extras["sip_channels"] = self._soft_int_set(
                        self.claro_siptrunk_channels_var,
                        field="CLARO Siptrunk - Troncal SIP (channels)",
                        default=int(CLARO_SIPTRUNK_DEFAULT_CHANNELS),
                        min_value=1,
                        allow_zero=False,
                    )

                elif st == "sbc":
                    invoice.extras["sbc_siptrunk_mbps"] = self._soft_int_set(
                        self.claro_sbc_siptrunk_mbps_var,
                        field="CLARO SBC - Siptrunk (MBPS)",
                        default=int(CLARO_SBC_DEFAULT_SIPTRUNK_MBPS),
                        min_value=1,
                        allow_zero=False,
                    )
                    invoice.extras["sbc_lic_qty"] = self._soft_int_set(
                        self.claro_sbc_lic_qty_var,
                        field="CLARO SBC - Licences (Qty)",
                        default=int(CLARO_SBC_DEFAULT_LIC_QTY),
                        min_value=1,
                        allow_zero=False,
                    )

                    # Prices ya normalizados por signo: los pasamos tal cual (lo que se ve es lo que se guarda)
                    invoice.extras["sbc_siptrunk_price"] = str(
                        parse_decimal_user_input(
                            self.claro_sbc_siptrunk_price_var.get(), field_name="Siptrunk price"
                        )
                    )
                    invoice.extras["sbc_lic_price"] = str(
                        parse_decimal_user_input(
                            self.claro_sbc_lic_price_var.get(), field_name="Licences price"
                        )
                    )

                else:  # mobile
                    invoice.extras["mobile_phone_lines_qty"] = self._soft_int_set(
                        self.claro_mobile_lines_qty_var,
                        field="CLARO Mobile - Phone lines quantity",
                        default=int(CLARO_MOBILE_DEFAULT_LINES_QTY),
                        min_value=1,
                        allow_zero=False,
                    )

                if concept_sel == OTRO:
                    invoice.extras["custom_concept"] = self.claro_custom_concept_var.get().strip()
                    if not has_split:
                        cc_raw = self.claro_cc_var.get().strip()
                        gl_raw = self.claro_gl_var.get().strip()
                        if not cc_raw or not gl_raw:
                            raise ValueError(
                                "Si no usas split, debes ingresar CC y GL para 1 línea (CLARO)."
                            )
                        invoice.extras["cc"] = int(cc_raw)
                        invoice.extras["gl_account"] = int(gl_raw)

                lines = build_lines(invoice)
                self.preview_lines = lines
                self._render_preview(lines)
                self.save_btn.configure(state="normal")
                self._update_preview_summary(invoice.subtotal)
                self._show_warnings()
                return

            # ---- EIKON ----
            if vendor.vendor_id == EIKON_ID:
                sel = self.eikon_concept_var.get()
                if sel == OTRO:
                    custom_concept = self.eikon_custom_concept_var.get().strip()
                    if not custom_concept and not has_split:
                        raise ValueError("Debes escribir el concepto personalizado.")
                    invoice.service_concept = (
                        custom_concept or self._first_alloc_concept() or "Concepto personalizado"
                    )
                    if not has_split:
                        cc_raw = self.eikon_custom_cc_var.get().strip()
                        gl_raw = self.eikon_custom_gl_var.get().strip()
                        if not cc_raw or not gl_raw:
                            raise ValueError(
                                "Si no usas split, debes ingresar CC y GL para 1 línea."
                            )
                        invoice.extras["cc"] = int(cc_raw)
                        invoice.extras["gl_account"] = int(gl_raw)
                else:
                    invoice.service_concept = sel
                    invoice.alloc_mode = None
                    invoice.allocations = []

            # ---- Generic ----
            else:
                sel = self.generic_concept_list_var.get().strip()
                if not sel:
                    raise ValueError("Debes seleccionar un Service/ concept.")
                invoice.service_concept = sel

                if vendor.vendor_id == CIRION_ID:
                    invoice.extras["bandwidth_mbps"] = self._soft_int_set(
                        self.bandwidth_var,
                        field="Bandwidth (MBPS)",
                        default=int(DEFAULT_BANDWIDTH),
                        min_value=1,
                        allow_zero=False,
                    )

                if vendor.vendor_id == MOVISTAR_ID:
                    invoice.extras["phone_lines_qty"] = self._soft_int_set(
                        self.phone_lines_var,
                        field="Phone lines quantity",
                        default=int(DEFAULT_PHONE_LINES),
                        min_value=1,
                        allow_zero=False,
                    )

                if sel == OTRO:
                    custom_concept = self.generic_custom_concept_var.get().strip()
                    if not custom_concept and not has_split:
                        raise ValueError("Debes escribir el concepto personalizado.")
                    invoice.service_concept = (
                        custom_concept or self._first_alloc_concept() or "Concepto personalizado"
                    )
                    if not has_split:
                        cc_raw = self.generic_cc_var.get().strip()
                        gl_raw = self.generic_gl_var.get().strip()
                        if not cc_raw or not gl_raw:
                            raise ValueError(
                                "Si no usas split, debes ingresar CC y GL para 1 línea."
                            )
                        invoice.extras["cc"] = int(cc_raw)
                        invoice.extras["gl_account"] = int(gl_raw)
                else:
                    invoice.alloc_mode = None
                    invoice.allocations = []

            lines = build_lines(invoice)
            self.preview_lines = lines
            self._render_preview(lines)
            self.save_btn.configure(state="normal")

            # ✅ actualizar resumen para TODOS (EIKON + genéricos)
            self._update_preview_summary(invoice.subtotal)

            self._show_warnings()

        except Exception as e:
            messagebox.showerror("Error de validación", str(e))

    def on_save(self) -> None:
        if not self.preview_lines:
            messagebox.showwarning(
                "Guardar", "No hay previsualización. Primero presiona 'Previsualizar'."
            )
            return

        # ✅ Confirmación adicional si hubo warnings (no bloqueante)
        if self._warnings:
            cont = messagebox.askyesno(
                "Advertencias detectadas",
                "Hay advertencias (correcciones automáticas) en los campos.\n"
                "¿Deseas continuar con el guardado de todas formas?\n\n"
                + "\n".join([f"- {w}" for w in self._warnings]),
            )
            if not cont:
                return

        try:
            vendor = self._selected_vendor()
            if not vendor:
                raise ValueError("Proveedor inválido.")
            bill = normalize_bill_number(self.bill_var.get(), field_name="Número de factura")

            resp = messagebox.askyesno(
                "Confirmar guardado",
                "Se guardarán las líneas en Excel.\n"
                "Si ya existe una factura con el mismo Vendor ID y Bill number en la tabla destino, "
                "se eliminarán esas líneas y se insertará el nuevo split.\n\n"
                "¿Deseas continuar?",
            )
            if not resp:
                return

            table_to_rows: Dict[str, List[Dict[str, object]]] = {}
            for li in self.preview_lines:
                table_to_rows.setdefault(li.table_name, []).append(li.values)

            backup_path, deleted_by_table, backup_created = apply_transaction(
                excel_path=self.excel_path,
                backup_dir=self.backup_dir,
                vendor_id=vendor.vendor_id,
                bill_number=bill,
                table_to_rows=table_to_rows,
                backup_path=self.session_backup_path,
                retention_keep_last_n=30,
                retention_keep_days=30,
            )
            self.session_backup_path = backup_path

            deleted_summary = "\n".join(
                [f"- {t}: {n} filas borradas" for t, n in deleted_by_table.items()]
            )
            backup_msg = (
                "Backup creado (sesión)" if backup_created else "Backup reutilizado (sesión)"
            )

            messagebox.showinfo(
                "Guardado exitoso",
                "✅ Guardado completado.\n\n"
                f"{backup_msg}\n"
                f"Backup:\n{backup_path}\n\n"
                "Sobrescritura (si aplicó):\n"
                f"{deleted_summary if deleted_summary else '- (sin borrados)'}",
            )

            self.on_clear()

        except ExcelWriteError as e:
            messagebox.showerror("Error al escribir Excel", str(e))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- Preview helpers ----------------

    def _clear_preview_state(self, clear_warnings: bool = True) -> None:
        """
        Limpia la previsualización y el resumen para forzar a que el usuario
        vuelva a presionar 'Previsualizar' cuando cambie el vendor.
        """
        self.preview_lines = []
        self._clear_tree()
        self._reset_preview_summary()
        self.save_btn.configure(state="disabled")

        if clear_warnings:
            self._clear_warnings()

    def _clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _render_preview(self, lines: List[LineItem]) -> None:
        self._clear_tree()
        for li in lines:
            v = li.values
            self.tree.insert(
                "",
                END,
                values=(
                    str(li.table_name),
                    str(v.get("Date", "")),
                    str(v.get("Bill number", "")),
                    str(v.get("Vendor", "")),
                    str(v.get("Service/ concept", "")),
                    str(v.get("CC", "")),
                    str(v.get("GL account", "")),
                    str(v.get("Subtotal assigned by CC", "")),
                    str(v.get("% IVA", "")),
                    str(v.get("IVA assigned by CC", "")),
                    str(v.get("Total assigned by CC", "")),
                ),
            )

    def _reset_preview_summary(self) -> None:
        self.preview_lines_count_var.set("0")
        self.preview_tables_var.set("0")
        self.preview_invoice_subtotal_var.set("0.00")
        self.preview_sum_subtotal_var.set("0.00")
        self.preview_diff_subtotal_var.set("0.00")
        self.preview_sum_iva_var.set("0.00")
        self.preview_sum_total_var.set("0.00")
        if hasattr(self, "preview_diff_label"):
            self.preview_diff_label.configure(bootstyle="secondary")

    def _as_decimal_safe(self, value: Any) -> Decimal:
        """
        Convierte value (Decimal/int/float/str) a Decimal de forma segura.
        Vacíos -> 0.
        """
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))

        s = str(value).strip()
        if s == "":
            return Decimal("0")

        # intenta directo
        try:
            return Decimal(s)
        except Exception:
            pass

        # intenta sin comas
        try:
            return Decimal(s.replace(",", ""))
        except Exception:
            return Decimal("0")

    def _update_preview_summary(self, invoice_subtotal: Decimal) -> None:
        """
        Calcula el resumen usando self.preview_lines y el subtotal de la factura.
        """
        lines = self.preview_lines or []
        self.preview_lines_count_var.set(str(len(lines)))

        tables = {li.table_name for li in lines}
        self.preview_tables_var.set(str(len(tables)))

        sum_sub = Decimal("0")
        sum_iva = Decimal("0")
        sum_total = Decimal("0")

        for li in lines:
            v = li.values
            sum_sub += self._as_decimal_safe(v.get("Subtotal assigned by CC"))
            sum_iva += self._as_decimal_safe(v.get("IVA assigned by CC"))
            sum_total += self._as_decimal_safe(v.get("Total assigned by CC"))

        diff = invoice_subtotal - sum_sub

        self.preview_invoice_subtotal_var.set(f"{invoice_subtotal:.2f}")
        self.preview_sum_subtotal_var.set(f"{sum_sub:.2f}")
        self.preview_diff_subtotal_var.set(f"{diff:.2f}")
        self.preview_sum_iva_var.set(f"{sum_iva:.2f}")
        self.preview_sum_total_var.set(f"{sum_total:.2f}")

        if hasattr(self, "preview_diff_label"):
            if diff.copy_abs() > Decimal("0.01"):
                self.preview_diff_label.configure(bootstyle="danger")
            elif diff != 0:
                self.preview_diff_label.configure(bootstyle="warning")
            else:
                self.preview_diff_label.configure(bootstyle="success")

        # Soft warning (como pediste)
        if diff != 0:
            self._add_warning(f"La suma de Subtotal assigned difiere del subtotal por {diff:.2f}.")
