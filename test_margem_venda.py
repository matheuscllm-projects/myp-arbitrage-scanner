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


# ── M1: fronteira exata e ROI-não-filtra, pelos filtros REAIS ─────────────
def _pk(name, compra, ref):
    from myp_arbitrage_scanner import CardData
    return CardData(name=f"{name} (001/100)", edition="E", rarity="R",
                    product_url=f"https://x/{name}", myp_lowest_en_nm=compra,
                    tcg_player_price=ref, margin_brl=ref - compra,
                    margin_pct=margem_venda(ref, compra), roi_pct=roi(ref, compra),
                    en_nm_sellers=3, tcg_source="tcgcsv", last_updated="x")


def _dbz(cls, name, compra, ref):
    return cls(name=name, en_name=name, code="C-1", code_norm="C-1", edition="E",
               section="dbsmasters" if cls.__name__.startswith("Dbz") else "onepiece",
               product_url=f"https://x/{name}", myp_lowest_en_nm=compra,
               en_nm_sellers=3, tcg_usd=1.0, tcg_brl=ref, tcg_url="https://t/1",
               rarity="R", subtype="Normal", join_via="j", margin_brl=ref - compra,
               margin_pct=margem_venda(ref, compra), roi_pct=roi(ref, compra),
               last_updated="x")


def test_pokemon_filtro_e_summary_fronteira_e_roi(tmp_path):
    from myp_arbitrage_scanner import generate_xlsx
    from myp_summary import build_markdown
    from openpyxl import load_workbook
    cards = [_pk("Borda", 50.47, 72.10), _pk("SoRoi", 75.0, 100.0)]
    p = tmp_path / "p.xlsx"
    generate_xlsx(cards, str(p), 0.30)
    names = [r[0] for r in load_workbook(p).worksheets[0].iter_rows(min_row=2, values_only=True) if r[0]]
    assert names == ["Borda (001/100)"], names
    md = tmp_path / "p.md"
    assert build_markdown(str(p), str(md), scan_type="daily") == 0
    txt = md.read_text(encoding="utf-8")
    limpos = txt.split("## ")[1:]
    assert any("Borda" in s and "limpos" in s.lower().split("\n")[0] for s in limpos), txt
    assert "SoRoi" not in txt.split("limpos")[1].split("\n## ")[0]


@pytest.mark.parametrize("mod,gen,cls", [
    ("myp_dbz_scanner", "generate_dbz_xlsx", "DbzCardData"),
    ("myp_op_scanner", "generate_op_xlsx", "OpCardData"),
])
def test_dbz_op_filtro_e_summary_fronteira_e_roi(tmp_path, mod, gen, cls):
    import importlib
    from openpyxl import load_workbook
    m = importlib.import_module(mod)
    summ = importlib.import_module(mod.replace("_scanner", "_summary"))
    C = getattr(m, cls)
    cards = [_dbz(C, "Borda", 50.47, 72.10), _dbz(C, "SoRoi", 75.0, 100.0)]
    p = tmp_path / "x.xlsx"
    getattr(m, gen)(cards, [], str(p), threshold=0.30, fx=5.0, stats={},
                    editions_scanned=1, min_price=50.0)
    wb = load_workbook(p)
    summary = {str(r[0]): r[1] for r in wb["Summary"].iter_rows(values_only=True)}
    assert summary["Deals Found"] == 1
    roi_col = [c.value for c in wb["All EN Cards"][1]].index(COL_ROI_) + 1
    assert wb["All EN Cards"].cell(row=2, column=roi_col).number_format == "0.00%"   # B2
    md = tmp_path / "x.md"
    assert summ.build_markdown(str(p), str(md)) == 0
    txt = md.read_text(encoding="utf-8")
    limpos = txt.split("Top 50 deals limpos")[1].split("\n## ")[0]
    assert "Borda" in limpos and "SoRoi" not in txt


from margin import COL_ROI as COL_ROI_  # noqa: E402


def test_aggregate_chunk_antigo_nomeia_arquivo_e_sai_limpo(tmp_path, monkeypatch, caplog):
    import myp_aggregate
    old = tmp_path / "chunk_velho.xlsx"
    _xlsx_antigo(old)
    monkeypatch.setattr("sys.argv", ["myp_aggregate.py", str(old), "-o", str(tmp_path / "o.xlsx")])
    assert myp_aggregate.main() == 1          # sem traceback
    assert "chunk_velho.xlsx" in caplog.text


def test_roi_col_derivada_do_cabecalho():
    import inspect, myp_arbitrage_scanner as m
    assert "ROI_COL = 25" not in inspect.getsource(m.generate_xlsx)
