# -*- coding: utf-8 -*-
"""
Base Stock Level Overzicht
==========================
Berekent de minimale basisvoorraad per component voor BPA bij verschillende
serviceniveaus. Gebruik dit script om een snel overzicht te krijgen voordat
je de volledige feasibility analyse draait.

λ_BPA = n_cust / MTBF(years)   (jaarlijkse vraag op basis van MTBF per klant)
μ     = λ_BPA × L              (verwachte vraag over leverancierslevertijd)
"""

import os
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from model import BPAOptimizationModel

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATIE – PAS HIER DE DRAAIKNOPPEN AAN
# ══════════════════════════════════════════════════════════════════════════════

EXCEL_PATH  = os.path.join(os.path.dirname(__file__),
                           'annual_use_abc_met_artikeldata_complete_europa.xlsx')
SHEET_NAME  = 'Filtered '

# ── Filters (zelfde als test_model.py) ────────────────────────────────────
FILTER_ARTICLE_TYPES = ['Critical', 'Onbekend']
FILTER_ABC           = ['A']
FILTER_INCOURANT     = ['Uitgeschakeld']
FILTER_MIN_VP        = 1000
FILTER_MIN_KLANTEN   = 5

# ── Parameters ─────────────────────────────────────────────────────────────
# Levertijd leverancier→BPA wordt per onderdeel geladen uit Excel (LT_days)

# Aantal subscripties (klanten) waarvoor BPA voorraadhoud – draaiknop
N_KLANTEN = 20

# ── Serviceniveaus om te tonen (kolommen in de tabel) ─────────────────────
SERVICE_LEVELS  = [0.980, 0.985, 0.99, 0.995, 0.998, 0.999]

# Fijnere reeks voor vloeiende grafiek – inclusief exacte SERVICE_LEVELS punten
SERVICE_LEVELS_PLOT = sorted(set(np.linspace(0.98, 0.999, 60).tolist() + SERVICE_LEVELS))


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LADEN
# ══════════════════════════════════════════════════════════════════════════════

def _parse_dutch_price(val):
    if pd.isna(val):
        return 0.0
    # Pandas leest numerieke cellen al als float – direct teruggeven
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().lstrip('\x80€ ').strip()
    # Dutch format: punt = duizendteken, komma = decimaalteken
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_parts() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df = df.rename(columns={
        'Verkooporderregel artikel.Artikel.Artikelcode': 'Code',
        'Omschrijving_standaard_artikelen':             'Descr',
        'Standaard verkoopprijs':                       'VP',
        'Inkoopprijs (standaard)':                      '_IP_raw',
        'Totaal_orders_5jr':                            'Totaal_orders_5jr',
        'Aantal_klantlocaties_5jr':                     'n_cust',
        'Hoofdleverancier.Levertijd':                   '_LT_raw',
        'MTBF(years)':                                  'MTBF',
    })
    df['IP']      = df['_IP_raw'].apply(_parse_dutch_price)
    df['LT_days'] = df['_LT_raw'].apply(
        lambda v: int(str(v).split()[0]) if pd.notna(v) and str(v)[0].isdigit() else 0
    )
    df['MTBF'] = pd.to_numeric(
        df.get('MTBF', pd.Series(np.nan, index=df.index)), errors='coerce'
    )
    mask = (
        df['ArticleType'].isin(FILTER_ARTICLE_TYPES) &
        df['ABC_categorie'].isin(FILTER_ABC) &
        df['Incourant'].isin(FILTER_INCOURANT) &
        (df['VP'] >= FILTER_MIN_VP) &
        (df['n_cust'] >= FILTER_MIN_KLANTEN)
    )
    return df[mask].set_index('Code')[['Descr', 'IP', 'Totaal_orders_5jr', 'n_cust', 'LT_days', 'MTBF']]


def _lambda_bpa(row) -> float:
    """λ_BPA = (1/MTBF) * N_KLANTEN  — vraag per onderdeel per klant × aantal subscripties.
    Fallback op Totaal_orders_5jr/5 als MTBF ontbreekt."""
    mtbf = row['MTBF']
    if pd.notna(mtbf) and mtbf > 0:
        return (1.0 / mtbf) * N_KLANTEN
    return row['Totaal_orders_5jr'] / 5


# ══════════════════════════════════════════════════════════════════════════════
#  HOOFDPROGRAMMA
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    # ── Diagnose: toon welke kolommen in de Excel staan ────────────────────
    _df_raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, nrows=0)
    _mtbf_match = [c for c in _df_raw.columns if 'mtbf' in c.lower()]
    print(f'\nDIAGNOSE – MTBF-gerelateerde kolommen in Excel: {_mtbf_match}')

    parts = load_parts()
    mtbf_ok = parts['MTBF'].notna().sum()
    print(f'MTBF geladen voor {mtbf_ok}/{len(parts)} onderdelen '
          f'({"fallback op historisch voor de rest" if mtbf_ok < len(parts) else "alle gevonden"})\n')

    print(f'\nBase Stock Overzicht  |  Levertijd: leverancier→BPA per onderdeel uit Excel  |  '
          f'Filters: ABC={FILTER_ABC}, VP≥€{FILTER_MIN_VP}')
    print(f'{len(parts)} onderdelen geselecteerd\n')

    # ── Kolombreedtes & header ─────────────────────────────────────────────
    CW = 7   # breedte per serviceniveau-kolom
    PW = 6   # breedte per s/N pooling-kolom
    LINE_W = 22 + 2 + 32 + 2 + 8 + 2 + 10 + 2 + 6 + 2 + 8 + 2 + 8 + 2 + len(SERVICE_LEVELS) * (CW + 2 + PW + 2)

    hdr  = f"{'Code':<22}  {'Omschrijving':<32}  {'MTBF(jr)':>8}  {'λ=n/MTBF':>10}  {'LT(d)':>6}  {'μ = λ·L':>8}  {'HC_BPA':>8}"
    for sl in SERVICE_LEVELS:
        hdr += f"  {'s@'+f'{sl:.1%}':>{CW}}  {'s/N':>{PW}}"
    print(hdr)
    print('─' * LINE_W)

    totals = {sl: 0 for sl in SERVICE_LEVELS}

    for code, row in parts.iterrows():
        lambda_bpa = _lambda_bpa(row)   # n_cust / MTBF(jr); fallback op historisch
        lt_yr = row['LT_days'] / 365
        mu = lambda_bpa * lt_yr
        hc_bpa = 0.20 * row['IP']
        mtbf_val = row['MTBF']
        mtbf_str = f"{mtbf_val:>8.2f}" if pd.notna(mtbf_val) else f"{'(hist)':>8}"

        stocks = {
            sl: BPAOptimizationModel.inverse_service_level(sl, lambda_bpa, lt_yr)
            for sl in SERVICE_LEVELS
        }

        line = (f"{code:<22}  {str(row['Descr'])[:32]:<32}  "
                f"{mtbf_str}  {lambda_bpa:>10.3f}  {int(row['LT_days']):>6}  {mu:>8.4f}  €{hc_bpa:>6.2f}")
        for sl in SERVICE_LEVELS:
            s = stocks[sl]
            line += f"  {s:>{CW}}  {s/N_KLANTEN:>{PW}.2f}"
            totals[sl] += s
        print(line)

    print('─' * LINE_W)
    totline = f"{'TOTAAL BASISVOORRAAD':<22}  {'':32}  {'':>10}  {'':>6}  {'':>8}  {'':>8}"
    for sl in SERVICE_LEVELS:
        totline += f"  {totals[sl]:>{CW}}  {totals[sl]/N_KLANTEN:>{PW}.2f}"
    print(totline)
    print()

    # ── Houdingskosttabel per serviceniveau ───────────────────────────────
    print('Totale houdingskosten BPA per serviceniveau (HC = 25% × inkoopprijs × basisvoorraad):')
    print('─' * 55)
    print(f"  {'Serviceniveau':<20} {'Totale HC_BPA':>18}  {'Δ t.o.v. vorige':>16}")
    print('─' * 55)
    prev_cost = None
    for sl in SERVICE_LEVELS:
        total_hc = sum(
            0.25 * parts.loc[code, 'IP'] *
            BPAOptimizationModel.inverse_service_level(
                sl,
                _lambda_bpa(parts.loc[code]),
                parts.loc[code, 'LT_days'] / 365
            )
            for code in parts.index
        )
        delta = f"+€{total_hc - prev_cost:.2f}" if prev_cost is not None else '—'
        print(f"  {sl:.1%}                    €{total_hc:>14.2f}  {delta:>16}")
        prev_cost = total_hc
    print('─' * 55)

    # ── Grafieken: basisvoorraad en houdingskosten vs. serviceniveau ──────
    print('\nGrafieken genereren...')

    HC_PERCENTAGES = [0.10, 0.15, 0.20, 0.25]
    HC_COLORS      = ['#2ca02c', '#ff7f0e', '#1f77b4', '#d62728']

    sl_pct = [sl * 100 for sl in SERVICE_LEVELS_PLOT]
    xticks      = [98, 98.5, 99, 99.5, 99.8, 99.9]
    xticklabels = ['98', '98.5', '99', '99.5', '99.8', '99.9']
    n = len(parts)

    # Helper: λ voor een specifiek aantal klanten n
    def _lam(row, n_val):
        mtbf = row['MTBF']
        if pd.notna(mtbf) and mtbf > 0:
            return (1.0 / mtbf) * n_val
        return row['Totaal_orders_5jr'] / 5

    # Bereken basisvoorraadcurves voor verschillende N
    N_VARIANTS = [1, 2, 5, 10, 50, 100]
    N_COLORS   = ['#2ca02c', '#ff7f0e', '#1f77b4', '#d62728', '#9467bd', '#8c564b']

    n_bs_curves = {}
    n_bs_marks  = {}
    for n_val in N_VARIANTS:
        curve = []
        for sl in SERVICE_LEVELS_PLOT:
            total = sum(
                BPAOptimizationModel.inverse_service_level(
                    sl, _lam(parts.loc[c], n_val), parts.loc[c, 'LT_days'] / 365
                )
                for c in parts.index
            )
            curve.append(total)
        n_bs_curves[n_val] = curve
        n_bs_marks[n_val] = [
            sum(
                BPAOptimizationModel.inverse_service_level(
                    sl, _lam(parts.loc[c], n_val), parts.loc[c, 'LT_days'] / 365
                )
                for c in parts.index
            )
            for sl in SERVICE_LEVELS
        ]

    # Bereken HC-curves per percentage
    hc_curves = {}
    hc_marks  = {}
    for pct in HC_PERCENTAGES:
        curve = []
        for sl in SERVICE_LEVELS_PLOT:
            total_hc = sum(
                pct * parts.loc[c, 'IP'] *
                BPAOptimizationModel.inverse_service_level(
                    sl, _lambda_bpa(parts.loc[c]), parts.loc[c, 'LT_days'] / 365
                )
                for c in parts.index
            )
            curve.append(total_hc)
        hc_curves[pct] = curve
        hc_marks[pct] = [
            sum(
                pct * parts.loc[c, 'IP'] *
                BPAOptimizationModel.inverse_service_level(
                    sl, _lambda_bpa(parts.loc[c]), parts.loc[c, 'LT_days'] / 365
                )
                for c in parts.index
            )
            for sl in SERVICE_LEVELS
        ]

    # ── Grafiek 1: Totale basisvoorraad (meerdere N) ──────────────────────
    fig1, ax_bs = plt.subplots(figsize=(9, 5))

    for n_val, color in zip(N_VARIANTS, N_COLORS):
        ax_bs.plot(sl_pct, n_bs_curves[n_val], color=color, linewidth=2.5, label=f'N = {n_val}')
        ax_bs.scatter([sl * 100 for sl in SERVICE_LEVELS], n_bs_marks[n_val],
                      color=color, zorder=5, s=50)
        for sl, bs in zip(SERVICE_LEVELS, n_bs_marks[n_val]):
            ax_bs.annotate(f'{bs}', xy=(sl * 100, bs), xytext=(4, 4),
                           textcoords='offset points', fontsize=7, color=color)

    ax_bs.set_xlabel('Service level (%)', fontsize=12)
    ax_bs.set_ylabel('Totale basisvoorraad (stuks)', fontsize=12)
    ax_bs.set_xlim(sl_pct[0], sl_pct[-1])
    ax_bs.set_xticks(xticks)
    ax_bs.set_xticklabels(xticklabels)
    ax_bs.legend(fontsize=10, loc='upper left')
    ax_bs.set_title(
        f'Totale basisvoorraad BPA vs. serviceniveau\n'
        f'({n} onderdelen, levertijd per onderdeel uit Excel)',
        fontsize=12
    )
    ax_bs.grid(True, alpha=0.3)
    fig1.tight_layout()

    plt.show()
