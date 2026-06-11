# -*- coding: utf-8 -*-
"""
Classificatie-module
====================
Pure scoring-logica voor spare parts classificatie. Wordt zowel door de
CLI (`classificatie_scoring.py`) als door de Streamlit-tab gebruikt.

Belangrijkste functies:
    laad_ruwe_dataset(path)              → DataFrame
    bereken_scores(df, params)           → DataFrame met Score_*/Gewogen_Score/Beslissing
    pas_harde_filters_toe(df, params)    → DataFrame
    bouw_selectie_payload(df, params)    → dict (klaar voor json.dump)
    schrijf_selectie_json(payload, path) → None
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict
from typing import Iterable

import numpy as np
import pandas as pd

# ── Kolomnamen in de bron-Excel ───────────────────────────────────────────────
COL_ABC          = "ABC_categorie"
COL_PRICE        = "Standaard verkoopprijs"
COL_LOCATIONS    = "Aantal_klantlocaties_met_orders_5jr"
COL_ORDERS       = "Gem_orders_per_klantlocatie_5jr"
COL_ARTICLE_TYPE = "ArticleType"

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
DESCR_COL_CANDIDATES = [
    "Omschrijving_standaard_artikelen",
    "Omschrijving",
    "Descr",
]
IP_COL_CANDIDATES = [
    "Inkoopprijs (standaard)",
    "Inkoopprijs",
]
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
TOTAAL_ORDERS_COL_CANDIDATES = [
    "Totaal_orders_5jr",
]
N_CUST_COL_CANDIDATES = [
    "Aantal_klantlocaties_5jr",
    "Aantal_klantlocaties_met_orders_5jr",
]

VERPLICHTE_KOLOMMEN = [COL_PRICE, COL_LOCATIONS, COL_ORDERS, COL_ARTICLE_TYPE]


# ── Parameters ────────────────────────────────────────────────────────────────

@dataclass
class ClassificatieParams:
    threshold:               float = 55.0
    weight_prijs:            float = 1 / 3
    weight_locaties:         float = 1 / 3
    weight_orders:           float = 1 / 3
    price_penalty_threshold: float = 1_000.0
    price_penalty_factor:    float = 0.4
    orders_power:            float = 2.0
    min_klantlocaties:       int = 5
    article_type_filter:     tuple = ("critical", "onbekend")  # case-insensitief
    lt_default_waarden:      tuple = ("30", "30 dagen", "30,0", "30.0")

    def normaliseer_weights(self) -> "ClassificatieParams":
        """Zorg dat de gewichten optellen tot 1."""
        s = self.weight_prijs + self.weight_locaties + self.weight_orders
        if s <= 0:
            return self
        return ClassificatieParams(
            **{**asdict(self),
               "weight_prijs":    self.weight_prijs    / s,
               "weight_locaties": self.weight_locaties / s,
               "weight_orders":   self.weight_orders   / s}
        )


# ── Laden ─────────────────────────────────────────────────────────────────────

def laad_ruwe_dataset(bron, sheet_name="Filtered ") -> pd.DataFrame:
    """Lees de volledige (ongefilterde) dataset.

    `bron` mag een pad of file-like object zijn. Standaard wordt de
    'Filtered '-sheet gelezen (gelijk aan bpa_beheer.SHEET_NAME); die bevat de
    correcte levertijden en MTBF(years). De eerste sheet 'Final_data' bevat
    placeholder-waarden (levertijd '0 dagen', MTBF in dagen) en mag hier niet
    gebruikt worden. Geef sheet_name=None om expliciet de eerste sheet te lezen.
    """
    df = pd.read_excel(bron, sheet_name=sheet_name) if sheet_name else pd.read_excel(bron)
    return df


def controleer_kolommen(df: pd.DataFrame) -> list[str]:
    """Geef lijst terug van ontbrekende verplichte kolommen."""
    return [c for c in VERPLICHTE_KOLOMMEN if c not in df.columns]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _find_col(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
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
    head = s.split()[0].replace(",", ".")
    try:
        n = int(float(head))
        return n if n > 0 else None
    except ValueError:
        return None


def _lt_bron(v, defaults: Iterable[str]) -> str:
    if pd.isna(v) or str(v).strip() == "":
        return "ontbreekt"
    # Een waarde die naar 0/None parseert (bv. "0 dagen") is geen bevestigde
    # levertijd maar een leeg/placeholder-veld → behandel als ontbrekend,
    # consistent met _parse_lt_dagen (die hier None teruggeeft).
    if _parse_lt_dagen(v) is None:
        return "ontbreekt"
    if str(v).strip() in set(defaults):
        return "default"
    return "geupdate"


def _mtbf_naar_jaren(raw, col_name: str | None) -> float | None:
    """Parseer MTBF-waarde als jaren, robuust voor diverse input-types.

    MTBF is altijd in jaren. Een kolomnaam of waarde die 'dagen'/'days'
    bevat wijst op een fout in de brondata — de numerieke waarde wordt
    ongewijzigd als jaren teruggegeven.

    Accepteert int, float, en strings (incl. NL-format ``"10,5"`` of
    ``"€ 1.234,56"``-achtige notatie). Geeft ``None`` bij leeg/onparseerbaar.
    """
    if col_name is None:
        return None
    # 1) None / NaN check (pd.isna faalt op str-arrays bij list-achtige input,
    #    daarom in try/except)
    try:
        if raw is None or pd.isna(raw):
            return None
    except (TypeError, ValueError):
        pass

    # 2) Numeriek (int/float/bool/Decimal) — direct casten
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        unit_van_value = None
    else:
        # 3) String-pad: strip whitespace, currency-prefix, parse NL-format,
        #    detecteer unit in de string zelf.
        s = str(raw).strip().lstrip('\x80€ ').strip()
        if not s:
            return None
        s_low = s.lower()
        unit_van_value = None
        if "dag" in s_low or "day" in s_low:
            unit_van_value = "dagen"
        elif "jaar" in s_low or "jaren" in s_low or "year" in s_low:
            unit_van_value = "jaren"
        # Houd alleen cijfers + decimaaltekens over
        cleaned = "".join(ch for ch in s if ch.isdigit() or ch in ",.-")
        if not cleaned or cleaned in ("-", ".", ","):
            return None
        if "," in cleaned:
            # NL: punt = duizendteken, komma = decimaal
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


# ── Scoring ───────────────────────────────────────────────────────────────────

def bereken_scores(df: pd.DataFrame, params: ClassificatieParams) -> pd.DataFrame:
    """
    Voegt de kolommen Score_Prijs, Score_Locaties, Score_Orders,
    Gewogen_Score en Classificatie_Beslissing toe.

    Berekent op de volledige dataset (percentielrangs!). Filteren gebeurt
    daarna via pas_harde_filters_toe.
    """
    p = params.normaliseer_weights()
    df = df.copy()

    # Score_Prijs: log + min-max (outlier-robust) + penalty onder threshold
    _price_log = np.log1p(df[COL_PRICE].clip(lower=0).fillna(0))
    _p_min, _p_max = _price_log.min(), _price_log.max()
    if _p_max > _p_min:
        df["Score_Prijs"] = ((_price_log - _p_min) / (_p_max - _p_min)) * 100
    else:
        df["Score_Prijs"] = 100.0
    below = df[COL_PRICE].fillna(0) < p.price_penalty_threshold
    df.loc[below, "Score_Prijs"] *= p.price_penalty_factor

    # Score_Locaties: log + min-max (outlier-robust), meer = beter
    _loc_log = np.log1p(df[COL_LOCATIONS].clip(lower=0).fillna(0))
    _l_min, _l_max = _loc_log.min(), _loc_log.max()
    if _l_max > _l_min:
        df["Score_Locaties"] = ((_loc_log - _l_min) / (_l_max - _l_min)) * 100
    else:
        df["Score_Locaties"] = 100.0

    # Score_Orders: log + inverse min-max met vaste floor=1.0, niet-lineair
    # Alle items met orders <= 1 krijgen score 100 (echte slow movers).
    _orders_floor = 1.0
    _orders_clipped = df[COL_ORDERS].clip(lower=_orders_floor)
    _orders_log = np.log1p(_orders_clipped)
    _of_log = np.log1p(_orders_floor)
    _om_log = _orders_log.max()
    if _om_log > _of_log:
        _scaled_orders = ((_om_log - _orders_log) / (_om_log - _of_log)).fillna(0.0)
    else:
        _scaled_orders = pd.Series(1.0, index=df.index)
    df["Score_Orders"] = (_scaled_orders ** p.orders_power) * 100

    df["Gewogen_Score"] = (
        df["Score_Prijs"]    * p.weight_prijs
        + df["Score_Locaties"] * p.weight_locaties
        + df["Score_Orders"]   * p.weight_orders
    ).round(1)

    for col in ["Score_Prijs", "Score_Locaties", "Score_Orders"]:
        df[col] = df[col].round(1)

    df["Classificatie_Beslissing"] = np.where(
        df["Gewogen_Score"] >= p.threshold,
        "Opnemen in lijst",
        "Niet opnemen",
    )

    # λ_i = 1 / MTBF(jaren) — failure rate per individueel component per jaar.
    # MTBF wordt robuust geconverteerd via _mtbf_naar_jaren (tekst/int/float
    # én eenheid-detectie 'dagen'/'days' worden afgevangen).
    _mtbf_col = _find_col(df, MTBF_COL_CANDIDATES)
    if _mtbf_col is not None:
        _mtbf_jr = df[_mtbf_col].apply(lambda v: _mtbf_naar_jaren(v, _mtbf_col))
        df["MTBF_jaren"] = _mtbf_jr.round(4)
        df["Lambda_jr"]  = (1.0 / _mtbf_jr).round(4)
    else:
        df["MTBF_jaren"] = np.nan
        df["Lambda_jr"]  = np.nan

    return df


def pas_harde_filters_toe(df: pd.DataFrame, params: ClassificatieParams) -> pd.DataFrame:
    """ArticleType-filter (case-insensitief) + minimum klantlocaties."""
    if COL_ARTICLE_TYPE in df.columns:
        df = df[df[COL_ARTICLE_TYPE].astype(str).str.strip().str.lower()
                .isin(set(s.lower() for s in params.article_type_filter))]
    if COL_LOCATIONS in df.columns:
        df = df[df[COL_LOCATIONS].fillna(0) >= params.min_klantlocaties]
    return df


# ── Selectie-payload ──────────────────────────────────────────────────────────

def bouw_selectie_payload(
    df_scored_filtered: pd.DataFrame,
    params: ClassificatieParams,
    bron_excel: str | None = None,
) -> dict:
    """Bouw het JSON-payload dat naar bpa_selectie.json wordt geschreven."""
    code_col  = _find_col(df_scored_filtered, CODE_COL_CANDIDATES)
    lt_col    = _find_col(df_scored_filtered, LT_COL_CANDIDATES)
    descr_col = _find_col(df_scored_filtered, DESCR_COL_CANDIDATES)
    ip_col    = _find_col(df_scored_filtered, IP_COL_CANDIDATES)
    vp_col    = COL_PRICE if COL_PRICE in df_scored_filtered.columns else None
    mtbf_col  = _find_col(df_scored_filtered, MTBF_COL_CANDIDATES)
    orders_col = _find_col(df_scored_filtered, TOTAAL_ORDERS_COL_CANDIDATES)
    ncust_col = _find_col(df_scored_filtered, N_CUST_COL_CANDIDATES)
    if code_col is None:
        raise ValueError(
            f"Geen artikelcode-kolom gevonden ({CODE_COL_CANDIDATES})."
        )

    sel = df_scored_filtered[df_scored_filtered["Classificatie_Beslissing"] == "Opnemen in lijst"].copy()

    def _num(v):
        """Parse numeric value, ook NL-format ('€ 1.234,56') of '1.234,56'."""
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

    items = []
    n_g = n_d = n_o = 0
    for _, row in sel.iterrows():
        lt_raw = row.get(lt_col) if lt_col else None
        bron   = _lt_bron(lt_raw, params.lt_default_waarden) if lt_col else "onbekend"
        if   bron == "geupdate":  n_g += 1
        elif bron == "default":   n_d += 1
        elif bron == "ontbreekt": n_o += 1
        items.append({
            "code":     str(row[code_col]),
            "score":    float(row["Gewogen_Score"]),
            "lt_dagen": _parse_lt_dagen(lt_raw) if lt_col else None,
            "lt_bron":  bron,
            "abc":      str(row.get(COL_ABC, "")),
            # Metadata waarmee bpa_beheer een rij kan bouwen voor codes
            # die NIET in de BPA-Excel (`Filtered `-sheet) staan.
            "descr":             (str(row[descr_col])[:80] if descr_col and pd.notna(row.get(descr_col)) else ""),
            "ip":                _num(row.get(ip_col)) if ip_col else None,
            "vp":                _num(row.get(vp_col)) if vp_col else None,
            # MTBF altijd in JAREN opgeslagen. Bij bron-kolom in dagen
            # ("MTBF (dagen)") wordt automatisch door 365 gedeeld.
            "mtbf":              _mtbf_naar_jaren(row.get(mtbf_col), mtbf_col),
            "totaal_orders_5jr": _num(row.get(orders_col)) if orders_col else None,
            "n_cust":            _num(row.get(ncust_col)) if ncust_col else None,
            # λ_jr = N_locaties / MTBF(jaren); float of None
            "lambda_jr":         (float(row["Lambda_jr"])
                                   if "Lambda_jr" in row.index
                                      and pd.notna(row["Lambda_jr"])
                                   else None),
        })

    return {
        "gegenereerd":  str(pd.Timestamp.today()),
        "bron_excel":   bron_excel,
        "threshold":    params.threshold,
        "parameters":   asdict(params),
        "n_items":      len(items),
        "lt_overzicht": {"geupdate": n_g, "default": n_d, "ontbreekt": n_o},
        "items":        items,
    }


def schrijf_selectie_json(payload: dict, path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ── Convenience: end-to-end run ───────────────────────────────────────────────

def voer_classificatie_uit(
    bron, params: ClassificatieParams, sheet_name="Filtered "
) -> tuple[pd.DataFrame, dict]:
    """End-to-end: laad → score → filter → bouw payload. Schrijft NIETS."""
    df_raw = laad_ruwe_dataset(bron, sheet_name=sheet_name)
    miss = controleer_kolommen(df_raw)
    if miss:
        raise ValueError(f"Ontbrekende kolommen: {miss}")
    df_scored   = bereken_scores(df_raw, params)
    df_filtered = pas_harde_filters_toe(df_scored, params)
    payload     = bouw_selectie_payload(df_filtered, params,
                                        bron_excel=str(bron) if isinstance(bron, str) else None)
    return df_filtered, payload
