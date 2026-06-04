"""
Classificatie Scoring Script
=============================
Voegt een gewogen score toe aan de spare parts classificatie:
- ABC_categorie        (A=200, B=50, C=0)   — 25% gewicht  (niet-lineair; A telt zwaar mee)
- Standaard verkoopprijs                    — 25% gewicht  (hoog = beter; penalty onder €1.000)
- Aantal_klantlocaties_met_orders_5jr       — 25% gewicht  (hoog = beter)
- Gem_orders_per_klantlocatie_5jr           — 25% gewicht  (laag = beter, inverse, gekwadreerd)

Max. gewogen score: 0.25×200 + 0.25×100 + 0.25×100 + 0.25×100 = 125
Drempel: gewogen score >= 60  →  "Opnemen in lijst"
"""

import json
import os
import pandas as pd
import numpy as np

# ── Paden (repo-relatief, override-baar via env-vars) ─────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.environ.get(
    "CLS_INPUT_FILE",
    os.path.join(SCRIPT_DIR, "annual_use_abc_met_artikeldata_complete_europa.xlsx"),
)
OUTPUT_FILE = os.environ.get(
    "CLS_OUTPUT_FILE",
    os.path.join(SCRIPT_DIR, "annual_use_abc_met_artikeldata_complete_europa_scored.xlsx"),
)
# Doelbestand voor de BPA-beheertool (whitelist + LT-metadata).
# Standaard: naast bpa_beheer.py (= deze repo-map).
SELECTIE_PATH = os.environ.get(
    "BPA_SELECTIE_PATH",
    os.path.join(SCRIPT_DIR, "bpa_selectie.json"),
)

THRESHOLD   = 55

COL_ABC          = "ABC_categorie"
COL_PRICE        = "Standaard verkoopprijs"
COL_LOCATIONS    = "Aantal_klantlocaties_met_orders_5jr"
COL_ORDERS       = "Gem_orders_per_klantlocatie_5jr"
COL_ARTICLE_TYPE = "ArticleType"

# Kandidaat-kolomnamen voor de artikelcode en levertijd (eerste gevonden wint).
CODE_COL_CANDIDATES = [
    "Verkooporderregel artikel.Artikel.Artikelcode",
    "Artikelcode",
    "Code",
]
LT_COL_CANDIDATES = [
    "Hoofdleverancier.Levertijd",
    "Levertijd",
    "LT_dagen",
]
# Levertijd-waarden die als 'niet-bevestigde ERP-default' worden gezien
LT_DEFAULT_WAARDEN = {"30", "30 dagen", "30,0", "30.0"}

# Harde filters
ARTICLE_TYPE_FILTER = {"critical", "onbekend"}   # case-insensitief
MIN_KLANTLOCATIES   = 5

WEIGHTS = {
    # COL_ABC:       0.25,   # ── ABC tijdelijk uitgeschakeld ──
    COL_PRICE:     1/3,
    COL_LOCATIONS: 1/3,
    COL_ORDERS:    1/3,
}

# Prijs-penalty: componenten onder deze drempel krijgen een gereduceerde prijsscore
PRICE_PENALTY_THRESHOLD = 1_000
PRICE_PENALTY_FACTOR    = 0.4   # score wordt vermenigvuldigd met dit getal (<1 = penalty)

# Machtsverheffing voor slow-moving score (>1 = sterker niet-lineair, echte slow movers scoren veel hoger)
ORDERS_POWER = 2.0
# ─────────────────────────────────────────────────────────────

print("Laden van Excel-bestand...")
df = pd.read_excel(INPUT_FILE)
print(f"  {len(df)} rijen geladen, {len(df.columns)} kolommen")

# ── Controleer of kolommen aanwezig zijn ──────────────────────
missing = [c for c in [COL_ABC, COL_PRICE, COL_LOCATIONS, COL_ORDERS, COL_ARTICLE_TYPE] if c not in df.columns]
if missing:
    raise ValueError(f"Ontbrekende kolommen in Excel: {missing}\nBeschikbare kolommen: {list(df.columns)}")

# ── 1. ABC_categorie score ────────────────────────────────────
# ── TIJDELIJK UITGESCHAKELD — verwijder het commentaar om weer te activeren ──
# A = hoge annual use   → hoogste score (200)
# B = gemiddeld                     → middenscore  (50)
# C = lage annual use  → laagste score (0)
# Bewust niet-lineair: de stap A→B is veel groter dan B→C
# abc_map = {"A": 100.0, "B": 75.0, "C": 0.0}
# df["Score_ABC"] = df[COL_ABC].astype(str).str.upper().str.strip().map(abc_map)
#
# unmapped = df["Score_ABC"].isna().sum()
# if unmapped > 0:
#     print(f"  Waarschuwing: {unmapped} rijen hebben een onbekende ABC-waarde → score = 0")
#     df["Score_ABC"] = df["Score_ABC"].fillna(0.0)

# ── 2. Standaard verkoopprijs score ──────────────────────────
# Hoge prijs = significant onderdeel → hogere score
# Penalty: componenten onder PRICE_PENALTY_THRESHOLD krijgen score × PRICE_PENALTY_FACTOR
df["Score_Prijs"] = df[COL_PRICE].rank(pct=True, na_option="bottom") * 100
below_threshold = df[COL_PRICE].fillna(0) < PRICE_PENALTY_THRESHOLD
df.loc[below_threshold, "Score_Prijs"] *= PRICE_PENALTY_FACTOR
print(f"  Prijs-penalty toegepast op {below_threshold.sum()} componenten (prijs < €{PRICE_PENALTY_THRESHOLD:,})")

# ── 3. Aantal klantlocaties score (commonality) ───────────────
# Meer locaties = betere commonality → hogere score
df["Score_Locaties"] = df[COL_LOCATIONS].rank(pct=True, na_option="bottom") * 100

# ── 4. Orders per klantlocatie score (INVERSE — slow-moving, niet-lineair) ──
# Minder orders = slow-moving = spare part gedrag → hogere score
# Machtsverheffing (ORDERS_POWER): echte slow movers worden extra beloond,
# fast movers zakken sterk weg (rank_pct=0.5 → score 25 i.p.v. 50 bij power=2)
rank_pct_orders = df[COL_ORDERS].rank(pct=True, ascending=False, na_option="bottom")
df["Score_Orders"] = (rank_pct_orders ** ORDERS_POWER) * 100

# ── 5. Gewogen eindscore (0–100) ──────────────────────────────
df["Gewogen_Score"] = (
    # df["Score_ABC"]      * WEIGHTS[COL_ABC]   # ── ABC uitgeschakeld ──
    + df["Score_Prijs"]    * WEIGHTS[COL_PRICE]
    + df["Score_Locaties"] * WEIGHTS[COL_LOCATIONS]
    + df["Score_Orders"]   * WEIGHTS[COL_ORDERS]
).round(1)

for col in ["Score_Prijs", "Score_Locaties", "Score_Orders"]:  # "Score_ABC" eruit
    df[col] = df[col].round(1)

# ── 6. Classificatiebeslissing ────────────────────────────────
df["Classificatie_Beslissing"] = np.where(
    df["Gewogen_Score"] >= THRESHOLD,
    "Opnemen in lijst",
    "Niet opnemen"
)

# ── Harde filters (ná scoring, zodat percentielrangs op volledige dataset blijven) ──
before = len(df)
df = df[df[COL_ARTICLE_TYPE].astype(str).str.strip().str.lower().isin(ARTICLE_TYPE_FILTER)]
df = df[df[COL_LOCATIONS].fillna(0) >= MIN_KLANTLOCATIES]
print(f"  Harde filters (ArticleType + locaties>={MIN_KLANTLOCATIES}): {before} → {len(df)} rijen ({before - len(df)} uitgesloten)")

# ── Samenvatting ──────────────────────────────────────────────
total      = len(df)
opnemen    = (df["Gewogen_Score"] >= THRESHOLD).sum()
niet       = total - opnemen

print(f"\n{'─'*50}")
print(f"  Totaal componenten      : {total:>6}")
print(f"  Opnemen in lijst (>={THRESHOLD}): {opnemen:>6}  ({opnemen/total*100:.1f}%)")
print(f"  Niet opnemen   (<{THRESHOLD}) : {niet:>6}  ({niet/total*100:.1f}%)")
print(f"{'─'*50}")

# Scorestatistieken per groep (ABC alleen ter info, telt niet meer in score)
print("\nGemiddelde scores per ABC-categorie (ABC telt NIET mee in Gewogen_Score):")
print(
    df.groupby(COL_ABC)[["Score_Prijs","Score_Locaties","Score_Orders","Gewogen_Score"]]
    .mean()
    .round(1)
    .to_string()
)

# ── Opslaan ───────────────────────────────────────────────────
print(f"\nOpslaan als: {OUTPUT_FILE}")
df.to_excel(OUTPUT_FILE, index=False)
print("Klaar!")

# ══════════════════════════════════════════════════════════════
#  EXPORT NAAR BPA-BEHEERTOOL (whitelist + LT-metadata)
# ══════════════════════════════════════════════════════════════

def _find_col(df_, candidates):
    for c in candidates:
        if c in df_.columns:
            return c
    return None

def _parse_lt_dagen(v):
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return int(v) if v > 0 else None
    s = str(v).strip()
    if not s:
        return None
    # Pak het eerste getal (bv. "30 dagen" → 30)
    head = s.split()[0].replace(",", ".")
    try:
        n = int(float(head))
        return n if n > 0 else None
    except ValueError:
        return None

def _lt_bron(v):
    """'geupdate' | 'default' | 'ontbreekt' — heuristiek o.b.v. ERP-default."""
    if pd.isna(v) or str(v).strip() == "":
        return "ontbreekt"
    if str(v).strip() in LT_DEFAULT_WAARDEN:
        return "default"
    return "geupdate"

_code_col = _find_col(df, CODE_COL_CANDIDATES)
_lt_col   = _find_col(df, LT_COL_CANDIDATES)

if _code_col is None:
    print(f"\n⚠  Geen artikelcode-kolom gevonden ({CODE_COL_CANDIDATES}). "
          f"bpa_selectie.json wordt NIET geschreven.")
else:
    selectie_df = df[df["Classificatie_Beslissing"] == "Opnemen in lijst"].copy()

    items = []
    n_geupdate = n_default = n_ontbreekt = 0
    for _, row in selectie_df.iterrows():
        lt_raw = row.get(_lt_col) if _lt_col else None
        bron   = _lt_bron(lt_raw) if _lt_col else "onbekend"
        if   bron == "geupdate":  n_geupdate  += 1
        elif bron == "default":   n_default   += 1
        elif bron == "ontbreekt": n_ontbreekt += 1
        items.append({
            "code":      str(row[_code_col]),
            "score":     float(row["Gewogen_Score"]),
            "lt_dagen":  _parse_lt_dagen(lt_raw) if _lt_col else None,
            "lt_bron":   bron,
            "abc":       str(row.get(COL_ABC, "")),
        })

    payload = {
        "gegenereerd":  str(pd.Timestamp.today()),
        "bron_excel":   INPUT_FILE,
        "threshold":    THRESHOLD,
        "n_items":      len(items),
        "lt_overzicht": {
            "geupdate":  n_geupdate,
            "default":   n_default,
            "ontbreekt": n_ontbreekt,
        },
        "items": items,
    }

    import os
    os.makedirs(os.path.dirname(SELECTIE_PATH) or ".", exist_ok=True)
    with open(SELECTIE_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"\n  ✓ BPA-selectie weggeschreven: {SELECTIE_PATH}")
    print(f"    {len(items)} componenten  |  "
          f"LT geupdate={n_geupdate}  default={n_default}  ontbreekt={n_ontbreekt}")
