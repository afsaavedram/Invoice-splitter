from __future__ import annotations

import importlib
import pkgutil
from typing import Callable, Dict, List, Optional

from invoice_splitter.models import InvoiceInput, LineItem

RuleFn = Callable[[InvoiceInput], List[LineItem]]

# Cache global (se inicializa una sola vez)
_RULES_CACHE: Optional[Dict[int, RuleFn]] = None


def _discover_vendor_modules() -> List[str]:
    """
    Retorna la lista de módulos dentro de invoice_splitter.rules cuyo nombre empieza por 'vendor_'.
    Ej: invoice_splitter.rules.vendor_1255097_eikon
    """
    import invoice_splitter.rules as rules_pkg

    module_names: List[str] = []
    for m in pkgutil.iter_modules(rules_pkg.__path__):
        if m.name.startswith("vendor_"):
            module_names.append(f"{rules_pkg.__name__}.{m.name}")
    return module_names


def _load_rules() -> Dict[int, RuleFn]:
    """
    Carga dinámicamente módulos vendor_*.py.
    Cada módulo debe exponer:
      - VENDOR_ID (int)
      - build_lines_for_vendor(invoice) -> list[LineItem]
    """
    rules: Dict[int, RuleFn] = {}

    for module_name in _discover_vendor_modules():
        mod = importlib.import_module(module_name)

        vendor_id = getattr(mod, "VENDOR_ID", None)
        fn = getattr(mod, "build_lines_for_vendor", None)

        # Validaciones defensivas
        if not isinstance(vendor_id, int):
            # No hacemos crash total: solo ignoramos módulo inválido
            continue
        if not callable(fn):
            continue

        # Si hay duplicados de vendor_id, fallamos temprano (mejor que comportamiento ambiguo)
        if vendor_id in rules:
            raise ValueError(
                f"Vendor ID duplicado detectado ({vendor_id}). "
                f"Revisa los módulos vendor_*.py. Duplicado al cargar: {module_name}"
            )

        rules[vendor_id] = fn

    return rules


def _get_rules() -> Dict[int, RuleFn]:
    global _RULES_CACHE
    if _RULES_CACHE is None:
        _RULES_CACHE = _load_rules()
    return _RULES_CACHE


def build_lines(invoice: InvoiceInput) -> List[LineItem]:
    rules = _get_rules()
    fn = rules.get(invoice.vendor_id)
    if not fn:
        raise ValueError(f"No hay regla implementada aún para vendor_id={invoice.vendor_id}")
    return fn(invoice)


def reload_rules() -> None:
    """
    Útil en desarrollo si agregas módulos sin reiniciar el proceso.
    En producción normalmente no lo necesitas.
    """
    global _RULES_CACHE
    _RULES_CACHE = None
