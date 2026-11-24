"""Margem canônica da frota: MARGEM SOBRE A VENDA; ROI é coluna à parte.

margem_venda = (ref - compra) / ref      -> filtra/classifica/ordena (corte 30 = 30% da venda)
roi          = (ref - compra) / compra   -> SÓ informativo (nunca filtra nem ordena)
compra ou ref ausente/<= 0 -> None.
"""
from __future__ import annotations

from typing import Optional

COL_MARGEM = "Margem venda %"
COL_ROI = "ROI aquisição %"
_OLD_COL = "Margin %"


def _ok(ref, compra) -> bool:
    return bool(ref and compra and ref > 0 and compra > 0)


def margem_venda(ref, compra) -> Optional[float]:
    return (ref - compra) / ref if _ok(ref, compra) else None


def roi(ref, compra) -> Optional[float]:
    return (ref - compra) / compra if _ok(ref, compra) else None


def passa_corte(ref, compra, corte: float) -> bool:
    m = margem_venda(ref, compra)
    return m is not None and m >= corte - 1e-9


def require_new_margin(headers) -> None:
    """XLSX antigo ("Margin %" = sobre a COMPRA) não pode ser lido como margem sobre a venda."""
    hs = set(headers)
    if _OLD_COL in hs and COL_MARGEM not in hs:
        raise ValueError(
            f'XLSX no formato antigo (coluna "{_OLD_COL}" = margem sobre a COMPRA). '
            f'Refaça o scan: o formato atual usa "{COL_MARGEM}" (sobre a venda).')
