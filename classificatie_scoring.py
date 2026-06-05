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
# Extra metadata-kolommen voor codes die NIET in de BPA-Excel (`Filtered`-sheet)
# staan: bpa_beheer.py gebruikt deze om een synthetische rij te bouwen met
# de juiste MTBF (→ λ = N / MTBF), VP, IP, omschrijving en historische orders.
MTBF_COL_CANDIDATES = [
    "MTBF(years)",
    "MTBF (years)",
    "MTBF_years",
    "MTBF(jaren)",
    "MTBF (jaren)",
    "MTBF (dagen)",
    "MTBF(dagen)",
    "MTBF (days)",
    "MTBF(days)",
    "MTBF_days",
    "MTBF",
]
IP_COL_CANDIDATES = [
    "Inkoopprijs (standaard)",
    "Inkoopprijs",
    "IP",
]
DESCR_COL_CANDIDATES = [
    "Omschrijving_standaard_artikelen",
    "Omschrijving",
    "Descr",
]
TOTAAL_ORDERS_COL_CANDIDATES = [
    "Totaal_orders_5jr",
    "Totaal orders 5jr",
]
N_CUST_COL_CANDIDATES = [
    "Aantal_klantlocaties_5jr",
    "Aantal_klantlocaties_met_orders_5jr",
    "n_cust",
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
# Min-max schaling: duurste onderdeel krijgt altijd 100, goedkoopste altijd 0,
# ongeacht hoeveel items dezelfde prijs hebben (geen rank-ties probleem).
# Penalty: componenten onder PRICE_PENALTY_THRESHOLD krijgen score × PRICE_PENALTY_FACTOR
_price_min = df[COL_PRICE].min()
_price_max = df[COL_PRICE].max()
if _price_max > _price_min:
    df["Score_Prijs"] = ((df[COL_PRICE].fillna(_price_min) - _price_min) / (_price_max - _price_min)) * 100
else:
    df["Score_Prijs"] = 100.0
below_threshold = df[COL_PRICE].fillna(0) < PRICE_PENALTY_THRESHOLD
df.loc[below_threshold, "Score_Prijs"] *= PRICE_PENALTY_FACTOR
print(f"  Prijs-penalty toegepast op {below_threshold.sum()} componenten (prijs < €{PRICE_PENALTY_THRESHOLD:,})")

# ── 3. Aantal klantlocaties score (commonality) ───────────────
# Meer locaties = betere commonality → hogere score
# Min-max schaling: meeste locaties krijgt 100, minste locaties 0.
_loc_min = df[COL_LOCATIONS].min()
_loc_max = df[COL_LOCATIONS].max()
if _loc_max > _loc_min:
    df["Score_Locaties"] = ((df[COL_LOCATIONS] - _loc_min) / (_loc_max - _loc_min)) * 100
else:
    df["Score_Locaties"] = 100.0

# ── 4. Orders per klantlocatie score (INVERSE — slow-moving, niet-lineair) ──
# Minder orders = slow-moving = spare part gedrag → hogere score
# Min-max schaling met VASTE floor op 1.0: alle onderdelen met
# gem_orders_per_klantlocatie ≤ 1 krijgen altijd score 100 (echte slow movers).
# Dit voorkomt dat een enkel outlier-item met orders < 1 ervoor zorgt dat
# items met orders = 1 onder de 100 zakken.
# Machtsverheffing (ORDERS_POWER): fast movers zakken extra weg (niet-lineair).
ORDERS_FLOOR = 1.0
_orders_clipped = df[COL_ORDERS].clip(lower=ORDERS_FLOOR)
_orders_max = _orders_clipped.max()
if _orders_max > ORDERS_FLOOR:
    _scaled_orders = ((_orders_max - _orders_clipped) / (_orders_max - ORDERS_FLOOR)).fillna(0.0)
else:
    _scaled_orders = pd.Series(1.0, index=df.index)
df["Score_Orders"] = (_scaled_orders ** ORDERS_POWER) * 100

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

# ── 7. λ_i = 1 / MTBF(jaren) ───────────────────────────────
# Failure rate per individueel component per jaar.
# Robuuste parser (_mtbf_naar_jaren_inline) accepteert tekst/int/float
# en converteert 'dagen'/'days' automatisch naar jaren.
def _find_col_inline(df_, candidates):
    for c in candidates:
        if c in df_.columns:
            return c
    return None

def _mtbf_naar_jaren_inline(raw, col_name):
    """Inline kopie van _mtbf_naar_jaren — robuust voor str/int/float."""
    if col_name is None:
        return None
    try:
        if raw is None or pd.isna(raw):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        unit_van_value = None
    else:
        s = str(raw).strip().lstrip('\x80€ ').strip()
        if not s:
            return None
        s_low = s.lower()
        unit_van_value = None
        if "dag" in s_low or "day" in s_low:
            unit_van_value = "dagen"
        elif "jaar" in s_low or "jaren" in s_low or "year" in s_low:
            unit_van_value = "jaren"
        cleaned = "".join(ch for ch in s if ch.isdigit() or ch in ",.-")
        if not cleaned or cleaned in ("-", ".", ","):
            return None
        if "," in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        try:
            val = float(cleaned)
        except ValueError:
            return None
    if val <= 0:
        return None
    # MTBF is altijd in jaren; dag-indicaties in waarde of kolomnaam zijn
    # brondata-fouten — de waarde wordt ongewijzigd als jaren teruggegeven.
    return val

_mtbf_col_pre = _find_col_inline(df, MTBF_COL_CANDIDATES)
if _mtbf_col_pre is not None:
    _mtbf_jr_series = df[_mtbf_col_pre].apply(
        lambda v: _mtbf_naar_jaren_inline(v, _mtbf_col_pre)
    )
    df["MTBF_jaren"] = _mtbf_jr_series.round(4) 
    df["Lambda_jr"]  = (1.0 / _mtbf_jr_series).round(4)
else:
    df["MTBF_jaren"] = np.nan
    df["Lambda_jr"]  = np.nan

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

def _num(v):
    """Parse numeriek (ook NL-format '€ 1.234,56'). Geeft None bij leeg/onparseerbaar."""
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lstrip('\x80€ ').strip()
    if not s:
        return None
    if ',' in s:
        # Dutch: punt = duizendteken, komma = decimaalteken
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None

def _lt_bron(v):
    """'geupdate' | 'default' | 'ontbreekt' — heuristiek o.b.v. ERP-default."""
    if pd.isna(v) or str(v).strip() == "":
        return "ontbreekt"
    if str(v).strip() in LT_DEFAULT_WAARDEN:
        return "default"
    return "geupdate"

def _mtbf_naar_jaren(raw, col_name):
    """Parseer MTBF-waarde als jaren, robuust voor diverse input-types.

    MTBF is altijd in jaren. Een kolomnaam of waarde die 'dagen'/'days'
    bevat wijst op een fout in de brondata — de numerieke waarde wordt
    ongewijzigd als jaren gebruikt. Accepteert int, float en strings
    (incl. NL-format ``"10,5"``).
    """
    if col_name is None:
        return None
    try:
        if raw is None or pd.isna(raw):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        unit_van_value = None
    else:
        s = str(raw).strip().lstrip('\x80€ ').strip()
        if not s:
            return None
        s_low = s.lower()
        unit_van_value = None
        if "dag" in s_low or "day" in s_low:
            unit_van_value = "dagen"
        elif "jaar" in s_low or "jaren" in s_low or "year" in s_low:
            unit_van_value = "jaren"
        cleaned = "".join(ch for ch in s if ch.isdigit() or ch in ",.-")
        if not cleaned or cleaned in ("-", ".", ","):
            return None
        if "," in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        try:
            val = float(cleaned)
        except ValueError:
            return None

    if val <= 0:
        return None

    # MTBF is altijd in jaren; dag-indicaties in waarde of kolomnaam zijn
    # brondata-fouten — de waarde wordt ongewijzigd als jaren teruggegeven.
    return val

_code_col  = _find_col(df, CODE_COL_CANDIDATES)
_lt_col    = _find_col(df, LT_COL_CANDIDATES)
_mtbf_col  = _find_col(df, MTBF_COL_CANDIDATES)
_ip_col    = _find_col(df, IP_COL_CANDIDATES)
_vp_col    = COL_PRICE if COL_PRICE in df.columns else None
_descr_col = _find_col(df, DESCR_COL_CANDIDATES)
_orders_col = _find_col(df, TOTAAL_ORDERS_COL_CANDIDATES)
_ncust_col = _find_col(df, N_CUST_COL_CANDIDATES)

if _code_col is None:
    print(f"\n⚠  Geen artikelcode-kolom gevonden ({CODE_COL_CANDIDATES}). "
          f"bpa_selectie.json wordt NIET geschreven.")
else:
    selectie_df = df[df["Classificatie_Beslissing"] == "Opnemen in lijst"].copy()

    print("  Metadata-kolommen in payload:")
    print(f"    code   = {_code_col}")
    print(f"    LT     = {_lt_col!r}")
    print(f"    MTBF   = {_mtbf_col!r}")
    print(f"    VP     = {_vp_col!r}")
    print(f"    IP     = {_ip_col!r}")
    print(f"    descr  = {_descr_col!r}")
    print(f"    orders = {_orders_col!r}")
    print(f"    n_cust = {_ncust_col!r}")

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
            # Metadata waarmee bpa_beheer een synthetische rij kan bouwen
            # voor codes die NIET in de BPA-Excel (`Filtered`-sheet) staan.
            # Zo krijgen die ook λ = N / MTBF in plaats van de fallback N/10.
            "descr":             (str(row[_descr_col])[:80]
                                   if _descr_col and pd.notna(row.get(_descr_col)) else ""),
            "ip":                _num(row.get(_ip_col)) if _ip_col else None,
            "vp":                _num(row.get(_vp_col)) if _vp_col else None,
            # MTBF altijd in JAREN opgeslagen; eventuele 'dagen'-aanduiding
            # in de brondata is een fout en wordt genegeerd.
            "mtbf":              _mtbf_naar_jaren(row.get(_mtbf_col), _mtbf_col),
            "totaal_orders_5jr": _num(row.get(_orders_col)) if _orders_col else None,
            "n_cust":            _num(row.get(_ncust_col)) if _ncust_col else None,
            # λ_jr = N_locaties / MTBF(jaren); float of None
            "lambda_jr":         (float(row["Lambda_jr"])
                                   if "Lambda_jr" in row.index
                                      and pd.notna(row["Lambda_jr"])
                                   else None),
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
