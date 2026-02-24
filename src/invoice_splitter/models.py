from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Any, Dict, List, Optional, Literal

from pydantic import BaseModel, Field


class Allocation(BaseModel):
    """
    Representa una línea de split custom.
    - concept: si None o vacío -> se usa invoice.service_concept (concepto general)
    - cc / gl_account: numéricos
    - percent: porcentaje como número 0..100 (ej: 40.91)  [modo 'percent']
    - amount: valor directo (Decimal)                      [modo 'amount']
    """

    concept: Optional[str] = None
    cc: int
    gl_account: int
    percent: Optional[Decimal] = None
    amount: Optional[Decimal] = None


class InvoiceInput(BaseModel):
    invoice_date: date
    vendor_id: int
    vendor_name: str
    bill_number: str  # 9 dígitos, texto
    subtotal: Decimal
    iva_rate: Decimal = Field(default=Decimal("0.15"))

    # Concepto general (por defecto)
    service_concept: Optional[str] = None

    # Para vendors con subtipo (Claro: siptrunk/sbc/mobile)
    service_type: Optional[str] = None

    # Extras flexibles (bandwidth, channels, etc.)
    extras: Dict[str, Any] = Field(default_factory=dict)

    # Split personalizado
    alloc_mode: Optional[Literal["percent", "amount"]] = None
    allocations: List[Allocation] = Field(default_factory=list)


class LineItem(BaseModel):
    """
    Una fila que será insertada en una Excel Table específica.
    `table_name` es el nombre exacto del ListObject (ej: 'Eikon_table').
    `values` es un dict: 'Nombre columna Excel' -> valor
    """

    table_name: str
    values: Dict[str, Any]
