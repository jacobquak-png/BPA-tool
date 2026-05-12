# -*- coding: utf-8 -*-
"""
BPA Feasibility Model - Analyse script
=======================================
Data wordt geladen uit de 'Filtered ' tab van de Excel.
Draaiknoppen staan bovenaan; resultaten zijn geschikt voor sensitivity grafieken.
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from model import BPAOptimizationModel

# Zet stdout op UTF-8 zodat speciale tekens werken in Windows terminal
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATIE – PAS HIER DE DRAAIKNOPPEN AAN
# ══════════════════════════════════════════════════════════════════════════════

EXCEL_PATH  = os.path.join(os.path.dirname(__file__),
                           'annual_use_abc_met_artikeldata_complete_europa.xlsx')
SHEET_NAME  = 'Filtered '          # tabblad in de Excel

# ── Filters (zelfde als de AutoFilters actief in het Excel-tabblad) ─────────
# Pas hier aan om andere onderdelen te selecteren
FILTER_ARTICLE_TYPES  = ['Critical', 'Onbekend']  # ArticleType
FILTER_ABC            = ['A']                      # ABC_categorie
FILTER_INCOURANT      = ['Uitgeschakeld']          # Incourant
FILTER_MIN_VP         = 1000                        # Standaard verkoopprijs >= €1000
FILTER_MIN_KLANTEN    = 5                          # Aantal_klantlocaties_5jr >= 5

# ── DRAAIKNOPPEN ───────────────────────────────────────────────────────────
SERVICE_LEVEL = 0.998   # X   – target serviceniveau BPA  [0.80 – 0.999]
ALPHA         = 0.05   # α   – abonnementsprijs als % van verkoopprijs  [0.01 – 0.50]
KAPPA_BPA     = 0.20   # κ_BPA = c_f^BPA + c_h^BPA + c_o^BPA
KAPPA_C       = 0.25   # κ_c   = c_f^c  + c_h^c  + c_o^c
N_CUSTOMERS   = 25     # n   – aantal subscripties in analyse

# ── Sensitivity sweep ranges (voor grafieken) ──────────────────────────────
SWEEP_SERVICE_LEVELS = [0.99, 0.9925, 0.995, 0.9975, 0.998, 0.999, 0.9995]
SWEEP_ALPHA          = [0.01, 0.02, 0.05, 0.08, 0.10, 0.12, 0.15, 0.20, 0.25]


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LADEN UIT EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def _parse_dutch_price(val):
    """Converteert Nederlandse prijsnotatie (bijv. '€ 1.069,70') naar float."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().lstrip('\x80€ ').strip()
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_parts_from_excel() -> pd.DataFrame:
    """
    Laadt artikeldata uit de Filtered tab en past dezelfde filters toe
    als de actieve AutoFilters in het Excel-tabblad:
        - ArticleType  in FILTER_ARTICLE_TYPES   (Critical, Onbekend)
        - ABC_categorie in FILTER_ABC             (A)
        - Incourant     in FILTER_INCOURANT       (Uitgeschakeld)
        - Standaard verkoopprijs >= FILTER_MIN_VP (750)
        - Aantal_klantlocaties_5jr >= FILTER_MIN_KLANTEN (5)

    Retourneert DataFrame geïndexeerd op artikelcode met kolommen:
        Descr, VP, IP, Totaal_orders_5jr, n_cust, LT_days, Obs
    """
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
    df['Obs'] = pd.to_numeric(
        df.get('Obsolescence cost', pd.Series(0, index=df.index)), errors='coerce'
    ).fillna(0.0)
    df['MTBF'] = pd.to_numeric(
        df['MTBF'] if 'MTBF' in df.columns else np.nan, errors='coerce'
    )

    # Pas Excel-AutoFilters toe
    mask = (
        df['ArticleType'].isin(FILTER_ARTICLE_TYPES) &
        df['ABC_categorie'].isin(FILTER_ABC) &
        df['Incourant'].isin(FILTER_INCOURANT) &
        (df['VP'] >= FILTER_MIN_VP) &
        (df['n_cust'] >= FILTER_MIN_KLANTEN)
    )
    df = df[mask].set_index('Code')
    return df[['Descr', 'VP', 'IP', 'Totaal_orders_5jr', 'n_cust', 'LT_days', 'Obs', 'MTBF']]


# ══════════════════════════════════════════════════════════════════════════════
#  MODEL BOUWEN & UITVOEREN
# ══════════════════════════════════════════════════════════════════════════════

def build_and_run(
    parts_df:      pd.DataFrame,
    service_level: float,
    alpha:         float,
    kappa_bpa:     float = KAPPA_BPA,
    kappa_c:       float = KAPPA_C,
    n_customers:   int   = N_CUSTOMERS,
) -> tuple[BPAOptimizationModel, dict]:
    """
    Bouwt het model en geeft (model, resultaat) terug.

    Demand keys : (part, customer)  → λ_in per jaar
    Levertijd   : leverancier→BPA per onderdeel uit Excel (LT_days / 365)
    κ_BPA       : kappa_bpa  (finance + opslag + obsolescence BPA)
    κ_c         : kappa_c    (finance + opslag + obsolescence klant base)
    """
    customers = {f'Klant_{i+1}': 1 for i in range(n_customers)}

    # Vraagdata: λ_in per klant per jaar
    # Primair: 1/MTBF; fallback: historisch gebruik / n_customers
    demand_data = {}
    for code, row in parts_df.iterrows():
        mtbf = row.get('MTBF') if 'MTBF' in row.index else float('nan')
        if pd.notna(mtbf) and mtbf > 0:
            units_yr_per_cust = 1.0 / mtbf
        else:
            units_yr_per_cust = (row['Totaal_orders_5jr'] / 5) / n_customers
        for i in range(n_customers):
            demand_data[(code, f'Klant_{i+1}')] = units_yr_per_cust

    cost_data = {
        'purchase_price': {code: row['IP'] for code, row in parts_df.iterrows()},
        'sales_price':    {code: row['VP'] for code, row in parts_df.iterrows()},
        # Levertijd leverancier→BPA per onderdeel (dagen → jaren)
        'lead_time': {code: row['LT_days'] / 365 for code, row in parts_df.iterrows()},
    }

    model = BPAOptimizationModel()
    model.add_customers_and_machines(customers)
    model.add_spare_parts(list(parts_df.index))
    model.add_parameters(
        service_level=service_level,
        alpha=alpha,
        kappa_bpa=kappa_bpa,
        kappa_c=kappa_c,
        demand_data=demand_data,
        cost_data=cost_data,
    )
    return model, model.check_feasibility()


# ══════════════════════════════════════════════════════════════════════════════
#  RESULTATEN PRINTEN
# ══════════════════════════════════════════════════════════════════════════════

def print_results(model: BPAOptimizationModel, result: dict, title: str = '') -> None:
    """Print een volledig resultaatoverzicht voor één configuratie."""
    W = 90
    if title:
        print(f"\n{'═' * W}")
        print(f"  {title}")
    print(f"{'═' * W}")
    print(f"  Service level : {result['service_level']:.0%}   "
          f"α: {result['alpha']:.2%}   "
          f"Levertijd: per onderdeel uit Excel (leverancier→BPA)")
    print(f"{'─' * W}")
    print(f"  HAALBAAR          : {'✓  JA' if result['feasible'] else '✗  NEE'}")
    print(f"  BPA winstgevend   : {'✓' if result['bpa_profitable'] else '✗'}  "
          f"(marge €{result['bpa_margin']:+.2f})")
    print(f"  Klanten profiteren: {'✓ allen' if result['all_customers_benefit'] else '✗ niet allen'}")

    # ── α-interval per component ──────────────────────────────────────────
    iv = result['alpha_intervals']
    al = iv['universal_alpha_L']
    au = iv['universal_alpha_U']
    if al is not None:
        print(f"  Universeel α-interval: [{al:.4%}, {au:.4%}]  "
              f"({'HAALBAAR' if iv['universal_feasible'] else 'NIET HAALBAAR'})")

    # ── BPA kosten & omzet per onderdeel ──────────────────────────────────
    bsl = model.calculate_base_stock_levels()
    print(f"\n  {'Onderdeel':<25} {'S*':>4} {'C_BPA':>10} {'Omzet':>10} "
          f"{'α_L,i':>8} {'α_U,i':>8} {'OK':>4}")
    print(f"  {'─' * W}")
    det = model.calculate_detailed_bpa_costs()
    per = iv['per_component']
    for code in model.sets['spare_parts']:
        stock = bsl.get(code, 0)
        cost  = det[code]['total']
        rev   = result['revenue_by_part'].get(code, 0)
        p     = per.get(code, {})
        al_i  = f"{p['alpha_L']:.3%}" if p.get('alpha_L') is not None else '—'
        au_i  = f"{p['alpha_U']:.3%}" if p.get('alpha_U') is not None else '—'
        ok    = '✓' if p.get('feasible') else '✗'
        print(f"  {code:<25} {stock:>4}  €{cost:>8.2f} €{rev:>8.2f} "
              f"{al_i:>8} {au_i:>8} {ok:>4}")

    print(f"  {'─' * W}")
    print(f"  {'TOTAAL':<31} €{result['bpa_costs']:>8.2f} €{result['total_revenue']:>8.2f}")

    # ── Klantkosten ────────────────────────────────────────────────────────
    print(f"\n  {'Klant':<12} {'Eigen voorraadkosten':>22} {'BPA abonnement':>16} "
          f"{'Besparing':>12} {'OK':>4}")
    print(f"  {'─' * 72}")
    for cust, b in result['customer_benefits'].items():
        print(f"  {cust:<12} €{b['self_stocking_cost']:>20.2f} "
              f"€{b['bpa_service_cost']:>14.2f} "
              f"€{b['savings']:>10.2f}  {'✓' if b['benefits'] else '✗'}")
    print()


# ══════════════════════════════════════════════════════════════════════════════
#  SENSITIVITY SWEEPS  – retourneren DataFrames (geschikt voor grafieken)
# ══════════════════════════════════════════════════════════════════════════════

def sweep_alpha(parts_df: pd.DataFrame,
                values: list[float] = SWEEP_ALPHA) -> pd.DataFrame:
    """
    Sweep over α; overige parameters op standaardwaarden.
    Retourneert DataFrame met kolommen:
        alpha, feasible, bpa_profitable, all_cust_benefit,
        bpa_costs, revenue, bpa_margin,
        universal_alpha_L, universal_alpha_U, universal_feasible
    """
    rows = []
    for a in values:
        _, r = build_and_run(parts_df, SERVICE_LEVEL, a)
        iv = r['alpha_intervals']
        rows.append({
            'alpha':               a,
            'feasible':            r['feasible'],
            'bpa_profitable':      r['bpa_profitable'],
            'all_cust_benefit':    r['all_customers_benefit'],
            'bpa_costs':           r['bpa_costs'],
            'revenue':             r['total_revenue'],
            'bpa_margin':          r['bpa_margin'],
            'universal_alpha_L':   iv['universal_alpha_L'],
            'universal_alpha_U':   iv['universal_alpha_U'],
            'universal_feasible':  iv['universal_feasible'],
        })
    return pd.DataFrame(rows)


def sweep_service_level(parts_df: pd.DataFrame,
                        values: list[float] = SWEEP_SERVICE_LEVELS) -> pd.DataFrame:
    """Sweep over service_level; overige parameters op standaardwaarden."""
    rows = []
    for sl in values:
        m, r = build_and_run(parts_df, sl, ALPHA)
        bsl = m.calculate_base_stock_levels()
        rows.append({
            'service_level':   sl,
            'feasible':        r['feasible'],
            'bpa_costs':       r['bpa_costs'],
            'revenue':         r['total_revenue'],
            'bpa_margin':      r['bpa_margin'],
            'total_basestock': sum(bsl.values()),
        })
    return pd.DataFrame(rows)


def plot_service_level_sweep(df: pd.DataFrame) -> None:
    """
    Plot service level sweep: kosten, opbrengsten en marge van BPA
    op de y-as; serviceniveau (%) op de x-as.
    Haalbare punten (feasible=True) worden volgetekend, niet-haalbare open.
    """
    fig, ax1 = plt.subplots(figsize=(10, 6))

    x = df['service_level'] * 100  # als percentage

    # Splits haalbaar / niet haalbaar voor markering
    feas = df['feasible']

    def _plot_line(ax, y_col, color, label, linestyle='-'):
        ax.plot(x, df[y_col], color=color, linestyle=linestyle,
                linewidth=2, label=label, zorder=2)
        # Haalbare punten: gevuld; niet-haalbare: open
        ax.scatter(x[feas],  df[y_col][feas],  color=color, s=60,
                   zorder=3, marker='o')
        ax.scatter(x[~feas], df[y_col][~feas], color=color, s=60,
                   zorder=3, marker='o', facecolors='white', linewidths=1.5)

    _plot_line(ax1, 'revenue',   '#2196F3', 'Omzet (€)')
    _plot_line(ax1, 'bpa_costs', '#F44336', 'BPA kosten (€)', linestyle='--')
    _plot_line(ax1, 'bpa_margin', '#4CAF50', 'Marge (€)',     linestyle=':')

    # Nul-lijn voor marge
    ax1.axhline(0, color='grey', linewidth=0.8, linestyle='-', zorder=1)

    # Verticale lijn op de standaard service level
    ax1.axvline(SERVICE_LEVEL * 100, color='grey', linewidth=1,
                linestyle='--', alpha=0.7, label=f'Huidig SL ({SERVICE_LEVEL:.1%})')

    # ── Rechter y-as: totaal voorraad (S*) ────────────────────────────────
    ax2 = ax1.twinx()
    color_stock = '#FF9800'
    ax2.plot(x, df['total_basestock'], color=color_stock, linestyle='-.',
             linewidth=2, label='Totaal voorraad (S*)', zorder=2)
    ax2.scatter(x[feas],  df['total_basestock'][feas],  color=color_stock, s=60,
                zorder=3, marker='s')
    ax2.scatter(x[~feas], df['total_basestock'][~feas], color=color_stock, s=60,
                zorder=3, marker='s', facecolors='white', linewidths=1.5)
    ax2.set_ylabel('Totaal basisvoorraad (stuks)', fontsize=12, color=color_stock)
    ax2.tick_params(axis='y', labelcolor=color_stock)
    ax2.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    ax1.set_xlabel('Service level (%)', fontsize=12)
    ax1.set_ylabel('Bedrag (€)', fontsize=12)
    ax1.set_title(f'BPA sensitivity – service level\n(α={ALPHA:.2%}, κ_BPA={KAPPA_BPA:.0%}, κ_c={KAPPA_C:.0%})',
                  fontsize=13)

    fmt_eur = mticker.FuncFormatter(lambda v, _: f'€{v:,.0f}')
    ax1.yaxis.set_major_formatter(fmt_eur)

    fmt_pct = mticker.FuncFormatter(lambda v, _: f'{v:.2f}%')
    ax1.xaxis.set_major_formatter(fmt_pct)
    ax1.set_xticks(x)
    plt.xticks(rotation=30, ha='right')

    # Gecombineerde legenda (beide assen)
    from matplotlib.lines import Line2D
    handles1, _ = ax1.get_legend_handles_labels()
    handles2, _ = ax2.get_legend_handles_labels()
    handles1.append(Line2D([0], [0], marker='o', color='grey',
                           label='● haalbaar  ○ niet haalbaar',
                           markerfacecolor='grey', markersize=7, linestyle='None'))
    ax1.legend(handles=handles1 + handles2, loc='best', fontsize=10)

    ax1.grid(True, alpha=0.3)
    fig.tight_layout()
    plt.show()


# ══════════════════════════════════════════════════════════════════════════════
#  HOOFDPROGRAMMA
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':

    # ── 1. Data laden ──────────────────────────────────────────────────────
    print('\nData laden uit Excel...')
    parts = load_parts_from_excel()
    print(f'Filters: ArticleType={FILTER_ARTICLE_TYPES}, ABC={FILTER_ABC}, '
          f'Incourant={FILTER_INCOURANT}, VP>={FILTER_MIN_VP}, '
          f'Klanten>={FILTER_MIN_KLANTEN}')
    print(f'{len(parts)} onderdelen geselecteerd.\n')

    print(f"{'Code':<25} {'Omschrijving':<35} {'VP':>9} {'IP':>9} "
          f"{'Vraag/jr':>10} {'#Klanten':>9} {'LT(d)':>7}")
    print('─' * 100)
    for code, row in parts.iterrows():
        units_yr = row['Totaal_orders_5jr'] / 5
        print(f"{code:<25} {str(row['Descr'])[:33]:<35} "
              f"€{row['VP']:>7.2f} €{row['IP']:>7.2f} "
              f"{units_yr:>9.3f} {int(row['n_cust']):>9} {int(row['LT_days']):>7}")

    # ── 2. Basisanalyse met standaard draaiknoppen ─────────────────────────
    model, result = build_and_run(
        parts, SERVICE_LEVEL, ALPHA
    )
    print_results(model, result,
                  title=f'BASISANALYSE  (SL={SERVICE_LEVEL:.0%}, α={ALPHA:.2%})')

    # ── 3. Sensitivity: alpha ──────────────────────────────────────────────
    print(f"\n{'═' * 80}")
    print(f"  SENSITIVITY: ALPHA  (SL={SERVICE_LEVEL:.0%})")
    print(f"{'═' * 80}")
    df_price = sweep_alpha(parts)
    print(df_price.to_string(index=False, float_format='{:.4f}'.format))

    # ── 4. Sensitivity: service level ─────────────────────────────────────
    print(f"\n{'═' * 80}")
    print(f"  SENSITIVITY: SERVICE LEVEL  (α={ALPHA:.2%})")
    print(f"{'═' * 80}")
    df_sl = sweep_service_level(parts)
    print(df_sl.to_string(index=False, float_format='{:.4f}'.format))

    # ── 5. Grafiek: service level vs kosten / opbrengsten / marge ──────────
    plot_service_level_sweep(df_sl)

    # ── 6. Gedetailleerde kosten & basisvoorraden bij standaard config ─────
    print()
    model.print_detailed_costs()
    model.print_base_stock_levels()
