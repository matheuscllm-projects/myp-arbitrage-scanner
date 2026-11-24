"""Casos de ouro da margem sobre a VENDA (idênticos em toda a frota).

margem_venda = (ref - compra)/ref ; roi = (ref - compra)/compra.
Corte = mesmo número (30), agora sobre a venda. ROI é informativo, nunca filtra.
"""
import pytest

from margin import margem_venda, roi, passa_corte


@pytest.mark.parametrize("compra,ref,margem,r,passa", [
    (75, 100, 0.25, 1 / 3, False),
    (70, 100, 0.30, 3 / 7, True),     # fronteira inclusiva
    (100, 80, -0.25, -0.20, False),
])
def test_casos_de_ouro(compra, ref, margem, r, passa):
    assert margem_venda(ref, compra) == pytest.approx(margem)
    assert roi(ref, compra) == pytest.approx(r)
    assert passa_corte(ref, compra, 0.30) is passa


@pytest.mark.parametrize("compra,ref", [(0, 100), (100, 0), (-1, 100), (100, -5), (None, 100), (100, None)])
def test_invalidos_none(compra, ref):
    assert margem_venda(ref, compra) is None
    assert roi(ref, compra) is None
    assert passa_corte(ref, compra, 0.30) is False


def _xlsx_antigo(path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "All EN Cards"
    ws.append(["Card Name", "Edition", "MYP EN NM (R$)", "TCG Player (R$)", "Margin %"])
    ws.append(["X", "Y", 70.0, 100.0, 0.4286])
    wb.save(path)


def test_xlsx_formato_antigo_falha_explicito(tmp_path):
    from myp_aggregate import load_chunk_cards
    from myp_summary import build_markdown
    p = tmp_path / "old.xlsx"
    _xlsx_antigo(p)
    with pytest.raises(ValueError, match="formato antigo"):
        load_chunk_cards(p)
    with pytest.raises(ValueError, match="formato antigo"):
        build_markdown(str(p), str(tmp_path / "o.md"), scan_type="daily")


def test_scraper_usa_venda_e_roi_informativo():
    from myp_arbitrage_scanner import generate_xlsx, CardData
    from openpyxl import load_workbook
    c = CardData(name="X (001/100)", edition="E", myp_lowest_en_nm=70.0,
                 tcg_player_price=100.0, margin_brl=30.0, margin_pct=margem_venda(100.0, 70.0),
                 roi_pct=roi(100.0, 70.0))
    p = "_tmp_margem.xlsx"
    try:
        generate_xlsx([c], p, 0.30)
        ws = load_workbook(p)["All EN Cards"]
        hdr = [x.value for x in ws[1]]
        row = dict(zip(hdr, [x.value for x in ws[2]]))
        assert "Margin %" not in hdr
        assert row["Margem venda %"] == pytest.approx(0.30)
        assert row["ROI aquisição %"] == pytest.approx(3 / 7)
    finally:
        import os
        os.path.exists(p) and os.remove(p)
