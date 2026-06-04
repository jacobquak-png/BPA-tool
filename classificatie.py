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

def laad_ruwe_dataset(bron, sheet_name=None) -> pd.DataFrame:
    """Lees de volledige (ongefilterde) dataset.

    `bron` mag een pad of file-like object zijn. Sheet=None → eerste sheet.
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
    if str(v).strip() in set(defaults):
        return "default"
    return "geupdate"


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

    # Score_Prijs: percentielrang + penalty onder threshold
    df["Score_Prijs"] = df[COL_PRICE].rank(pct=True, na_option="bottom") * 100
    below = df[COL_PRICE].fillna(0) < p.price_penalty_threshold
    df.loc[below, "Score_Prijs"] *= p.price_penalty_factor

    # Score_Locaties: meer = beter
    df["Score_Locaties"] = df[COL_LOCATIONS].rank(pct=True, na_option="bottom") * 100

    # Score_Orders: minder = beter, niet-lineair
    rank_pct_orders = df[COL_ORDERS].rank(pct=True, ascending=False, na_option="bottom")
    df["Score_Orders"] = (rank_pct_orders ** p.orders_power) * 100

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
    code_col = _find_col(df_scored_filtered, CODE_COL_CANDIDATES)
    lt_col   = _find_col(df_scored_filtered, LT_COL_CANDIDATES)
    if code_col is None:
        raise ValueError(
            f"Geen artikelcode-kolom gevonden ({CODE_COL_CANDIDATES})."
        )

    sel = df_scored_filtered[df_scored_filtered["Classificatie_Beslissing"] == "Opnemen in lijst"].copy()

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
    bron, params: ClassificatieParams, sheet_name=None
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
