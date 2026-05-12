# -*- coding: utf-8 -*-
"""
BPA Spare Parts – Monte Carlo Feasibility Simulatie
====================================================
Simuleert per run welke klanten welk onderdeel gebruiken (gerandomiseerd
op basis van de historische klantspreiding uit de Excel) en checkt wanneer
het BPA-model haalbaar is.

Draaiknoppen staan bovenaan in de CONFIG-sectie.
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

EXCEL_PATH = os.path.join(os.path.dirname(__file__),
                          'annual_use_abc_met_artikeldata_complete_europa.xlsx')
SHEET_NAME = 'Filtered '

# ── Filters (zelfde als actieve AutoFilters in Excel) ─────────────────────
FILTER_ARTICLE_TYPES = ['Critical', 'Onbekend']
FILTER_ABC           = ['A']
FILTER_INCOURANT     = ['Uitgeschakeld']
FILTER_MIN_VP        = 1000
FILTER_MIN_KLANTEN   = 5

# ── DRAAIKNOPPEN ──────────────────────────────────────────────────────────
N_CUSTOMERS   = 30     # Aantal BPA-klanten in de simulatie
N_RUNS        = 200    # Aantal Monte Carlo runs per configuratie
SERVICE_LEVEL = 0.99   # Target serviceniveau β
ALPHA         = 0.15   # α – abonnementsprijs als % van verkoopprijs
KAPPA_BPA     = 0.20   # κ_BPA = c_f^BPA + c_h^BPA + c_o^BPA
KAPPA_C       = 0.25   # κ_c   = c_f^c  + c_h^c  + c_o^c
SEED          = 42     # Reproduceerbare randomisatie (None = willekeurig)

# P(klant gebruikt onderdeel) = n_cust / N_REF_CUSTOMERS
N_REF_CUSTOMERS = 10

# ── Sweep-reeksen voor gevoeligheidsanalyse ──────────────────────────────────
SWEEP_N_CUSTOMERS = [5, 10, 15, 20, 25, 30]
SWEEP_ALPHA       = [0.01, 0.02, 0.05, 0.08, 0.10, 0.12, 0.15, 0.20, 0.25, 0.30]


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LADEN
# ══════════════════════════════════════════════════════════════════════════════

def _parse_dutch_price(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip().lstrip('\x80€ ').strip()
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
        'MTBF(years)':                                   'MTBF',
    })
    df['IP'] = df['_IP_raw'].apply(_parse_dutch_price)
    df['LT_days'] = df['_LT_raw'].apply(
        lambda v: int(str(v).split()[0]) if pd.notna(v) and str(v)[0].isdigit() else 0
    )
    df['Obs'] = pd.to_numeric(
        df.get('Obsolescence cost', pd.Series(0, index=df.index)), errors='coerce'
    ).fillna(0.0)
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
    df = df[mask].set_index('Code')
    return df[['Descr', 'VP', 'IP', 'Totaal_orders_5jr', 'n_cust', 'LT_days', 'Obs', 'MTBF']]


# ══════════════════════════════════════════════════════════════════════════════
#  SIMULATIE – ÉÉN RUN
# ══════════════════════════════════════════════════════════════════════════════

def _run_once(parts_df: pd.DataFrame,
              rng: np.random.Generator,
              n_customers: int,
              service_level: float,
              alpha: float,
              kappa_bpa: float,
              kappa_c: float) -> dict:
    """
    Één Monte Carlo run.

    Demand keys : (part, customer)  → λ_in per jaar
    κ_BPA / κ_c : carrying rates (incl. obs)
    α           : subscription percentage
    """
    customers = {f'Klant_{i+1}': 1 for i in range(n_customers)}
    demand_data = {}
    customers_per_part = []

    for code, row in parts_df.iterrows():
        p_use = min(row['n_cust'] / N_REF_CUSTOMERS, 1.0)
        uses = rng.random(n_customers) < p_use
        k = int(uses.sum())
        customers_per_part.append(k)

        mtbf = row['MTBF']
        if pd.notna(mtbf) and mtbf > 0:
            demand_per_cust = 1.0 / mtbf
        else:
            demand_per_cust = (row['Totaal_orders_5jr'] / 5) / row['n_cust']
        for i in range(n_customers):
            demand_data[(code, f'Klant_{i+1}')] = demand_per_cust if uses[i] else 0.0

    cost_data = {
        'purchase_price': {c: r['IP'] for c, r in parts_df.iterrows()},
        'sales_price':    {c: r['VP'] for c, r in parts_df.iterrows()},
        'obsolescence_rate_c': {
            c: (r['Obs'] / r['VP'] if r['VP'] > 0 else 0.0)
            for c, r in parts_df.iterrows()
        },
        'lead_time': {c: r['LT_days'] / 365 for c, r in parts_df.iterrows()},
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

    result = model.check_feasibility()
    result['n_parts_active']         = len(parts_df)
    result['avg_customers_per_part'] = float(np.mean(customers_per_part))
    result['bpa_detail']             = model.calculate_detailed_bpa_costs()

    return result


# ══════════════════════════════════════════════════════════════════════════════
#  SIMULATIE – MEERDERE RUNS
# ══════════════════════════════════════════════════════════════════════════════

def run_simulation(parts_df: pd.DataFrame,
                   n_customers: int   = N_CUSTOMERS,
                   n_runs: int        = N_RUNS,
                   service_level: float = SERVICE_LEVEL,
                   alpha: float        = ALPHA,
                   kappa_bpa: float    = KAPPA_BPA,
                   kappa_c: float      = KAPPA_C,
                   seed: int           = SEED,
                   verbose: bool       = False,
                   print_base_stock: bool = False) -> pd.DataFrame:
    """
    Voert n_runs Monte Carlo simulaties uit en retourneert een DataFrame
    met per run de feasibility-uitkomsten.
    """
    rng = np.random.default_rng(seed)
    records = []
    bpa_details = []

    for run_idx in range(n_runs):
        if verbose and (run_idx + 1) % 50 == 0:
            print(f'  run {run_idx + 1}/{n_runs}...')
        r = _run_once(parts_df, rng, n_customers, service_level,
                      alpha, kappa_bpa, kappa_c)
        bpa_details.append(r['bpa_detail'])
        records.append({
            'run':                   run_idx + 1,
            'feasible':              r['feasible'],
            'bpa_profitable':        r['bpa_profitable'],
            'all_customers_benefit': r['all_customers_benefit'],
            'bpa_margin':            r['bpa_margin'],
            'bpa_costs':             r['bpa_costs'],
            'total_revenue':         r['total_revenue'],
            'avg_cust_per_part':     r['avg_customers_per_part'],
        })

    if print_base_stock and bpa_details:
        print(f'\n  Gemiddelde base stock niveaus over {n_runs} runs:')
        print(f'  {"Code":<30} {"gem. λ_BPA":>10} {"gem. s_min":>10} {"gem. HC-kost":>12} {"actief%":>8}')
        print(f'  {"─"*30} {"─"*10} {"─"*10} {"─"*12} {"─"*8}')
        for part in parts_df.index:
            active = [d[part] for d in bpa_details if d.get(part, {}).get('demand', 0) > 0]
            if not active:
                continue
            avg_demand = float(np.mean([x['demand'] for x in active]))
            avg_s      = float(np.mean([x['base_stock'] for x in active]))
            avg_cost   = float(np.mean([x['total'] for x in active]))
            pct_active = len(active) / n_runs
            print(f'  {str(part):<30} {avg_demand:>10.4f} {avg_s:>10.2f} €{avg_cost:>11.2f} {pct_active:>7.1%}')
        print()

    return pd.DataFrame(records)


def sweep_n_customers(parts_df: pd.DataFrame,
                      values: list = SWEEP_N_CUSTOMERS,
                      n_runs: int  = N_RUNS) -> pd.DataFrame:
    """
    Sweep over het aantal klanten; retourneert feasibility-statistieken per waarde.
    """
    rows = []
    for n in values:
        print(f'  Sweep n_customers={n}...')
        df = run_simulation(parts_df, n_customers=n, n_runs=n_runs)
        rows.append({
            'n_customers':           n,
            'feasibility_rate':      df['feasible'].mean(),
            'bpa_profitable_rate':   df['bpa_profitable'].mean(),
            'all_cust_benefit_rate': df['all_customers_benefit'].mean(),
            'avg_bpa_margin':        df['bpa_margin'].mean(),
            'avg_bpa_costs':         df['bpa_costs'].mean(),
            'avg_revenue':           df['total_revenue'].mean(),
        })
    return pd.DataFrame(rows)


def sweep_alpha(parts_df: pd.DataFrame,
                values: list = SWEEP_ALPHA,
                n_runs: int  = N_RUNS) -> pd.DataFrame:
    """
    Sweep over α: bij welk percentage wordt het model feasible?
    """
    rows = []
    for a in values:
        print(f'  Sweep α={a:.2%}...')
        df = run_simulation(parts_df, alpha=a, n_runs=n_runs)
        rows.append({
            'alpha':                 a,
            'feasibility_rate':      df['feasible'].mean(),
            'bpa_profitable_rate':   df['bpa_profitable'].mean(),
            'all_cust_benefit_rate': df['all_customers_benefit'].mean(),
            'avg_bpa_margin':        df['bpa_margin'].mean(),
            'avg_bpa_costs':         df['bpa_costs'].mean(),
            'avg_revenue':           df['total_revenue'].mean(),
        })
    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════════════
#  RESULTATEN PRINTEN
# ══════════════════════════════════════════════════════════════════════════════

def print_simulation_summary(df: pd.DataFrame, n_customers: int) -> None:
    W = 65
    print(f'\n{"═" * W}')
    print(f'  SIMULATIE-RESULTATEN  ({len(df)} runs, {n_customers} klanten)')
    print(f'{"═" * W}')
    print(f'  Feasibility rate           : {df["feasible"].mean():.1%}')
    print(f'  BPA winstgevend (rate)     : {df["bpa_profitable"].mean():.1%}')
    print(f'  Alle klanten profiteren (r): {df["all_customers_benefit"].mean():.1%}')
    print(f'{"─" * W}')
    print(f'  BPA marge  – gem. : €{df["bpa_margin"].mean():>10.2f}')
    print(f'               min  : €{df["bpa_margin"].min():>10.2f}')
    print(f'               max  : €{df["bpa_margin"].max():>10.2f}')
    print(f'               std  : €{df["bpa_margin"].std():>10.2f}')
    print(f'{"─" * W}')
    print(f'  Gem. kosten BPA  : €{df["bpa_costs"].mean():>10.2f}')
    print(f'  Gem. omzet BPA   : €{df["total_revenue"].mean():>10.2f}')
    print(f'{"═" * W}\n')


# ══════════════════════════════════════════════════════════════════════════════
#  GRAFIEKEN
# ══════════════════════════════════════════════════════════════════════════════

def plot_simulation(df_runs: pd.DataFrame, n_customers: int) -> None:
    """Grafiek 1 + 2: verdeling BPA-marge en feasibility-breakdown per run."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5))

    # ── Histogram BPA marge ────────────────────────────────────────────────
    feasible_margins     = df_runs.loc[df_runs['feasible'],     'bpa_margin']
    not_feasible_margins = df_runs.loc[~df_runs['feasible'],    'bpa_margin']

    bins = np.linspace(df_runs['bpa_margin'].min(), df_runs['bpa_margin'].max(), 30)
    ax1.hist(not_feasible_margins, bins=bins, color='#d62728', alpha=0.7, label='Niet haalbaar')
    ax1.hist(feasible_margins,     bins=bins, color='#2ca02c', alpha=0.7, label='Haalbaar')
    ax1.axvline(0, color='black', linewidth=1.2, linestyle='--')
    ax1.set_xlabel('BPA marge (€)', fontsize=11)
    ax1.set_ylabel('Aantal runs', fontsize=11)
    ax1.set_title(f'Verdeling BPA-marge\n({len(df_runs)} runs, {n_customers} klanten)', fontsize=11)
    ax1.legend(fontsize=9)
    ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'€{x:,.0f}'))
    ax1.grid(True, alpha=0.3)

    # ── Gestapeld staafdiagram: feasibility per 20 runs ───────────────────
    group_size = max(1, len(df_runs) // 10)
    groups = []
    labels_g = []
    for start in range(0, len(df_runs), group_size):
        chunk = df_runs.iloc[start:start + group_size]
        groups.append([
            chunk['feasible'].sum(),
            chunk['bpa_profitable'].sum() - chunk['feasible'].sum(),
            (~chunk['bpa_profitable']).sum(),
        ])
        labels_g.append(f'{start+1}–{min(start+group_size, len(df_runs))}')

    groups = np.array(groups).T
    x = np.arange(len(labels_g))
    bar_w = 0.6
    ax2.bar(x, groups[0], bar_w, label='Haalbaar',              color='#2ca02c')
    ax2.bar(x, groups[1], bar_w, bottom=groups[0],
            label='BPA winstgevend, klant niet',                 color='#ff7f0e')
    ax2.bar(x, groups[2], bar_w, bottom=groups[0] + groups[1],
            label='BPA niet winstgevend',                        color='#d62728')
    ax2.set_xticks(x)
    ax2.set_xticklabels(labels_g, rotation=45, ha='right', fontsize=8)
    ax2.set_ylabel('Aantal runs', fontsize=11)
    ax2.set_title(f'Feasibility per run-groep\n({n_customers} klanten)', fontsize=11)
    ax2.legend(fontsize=8, loc='upper right')
    ax2.grid(True, alpha=0.3, axis='y')

    fig.tight_layout()


def plot_sweep_price(df_sweep: pd.DataFrame) -> None:
    """Grafiek 3: feasibility rate en gemiddelde BPA-marge vs. α."""
    fig, ax1 = plt.subplots(figsize=(10, 5))
    color_feas  = '#1f77b4'
    color_marge = '#d62728'
    color_cust  = '#ff7f0e'

    x_pct = df_sweep['alpha'] * 100

    ax1.plot(x_pct, df_sweep['feasibility_rate'] * 100,
             color=color_feas, linewidth=2.5, marker='o', label='Feasibility rate (%)')
    ax1.plot(x_pct, df_sweep['bpa_profitable_rate'] * 100,
             color='#2ca02c', linewidth=1.8, marker='^', linestyle='--',
             label='BPA winstgevend rate (%)')
    ax1.plot(x_pct, df_sweep['all_cust_benefit_rate'] * 100,
             color=color_cust, linewidth=1.8, marker='s', linestyle=':',
             label='Alle klanten profiteren rate (%)')
    ax1.set_xlabel('α (abonnementsprijs als % van verkoopprijs)', fontsize=12)
    ax1.set_ylabel('Rate (%)', fontsize=12)
    ax1.set_ylim(0, 105)
    ax1.set_xticks(x_pct.tolist())
    ax1.set_xticklabels([f'{p:.0f}%' for p in x_pct], rotation=45)

    ax2 = ax1.twinx()
    ax2.plot(x_pct, df_sweep['avg_bpa_margin'],
             color=color_marge, linewidth=2, linestyle='--', marker='D',
             label='Gem. BPA-marge (€)')
    ax2.axhline(0, color='black', linewidth=0.8, linestyle=':')
    ax2.set_ylabel('Gemiddelde BPA-marge (€)', color=color_marge, fontsize=12)
    ax2.tick_params(axis='y', labelcolor=color_marge)
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'€{x:,.0f}'))

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='center right', fontsize=9)
    ax1.set_title(
        f'Feasibility vs. abonnementsprijs\n'
        f'(SL={SERVICE_LEVEL:.0%}, n={N_CUSTOMERS} klanten, '
        f'{N_RUNS} runs per punt, \u03ba_BPA={KAPPA_BPA:.0%})',
        fontsize=12
    )
    ax1.grid(True, alpha=0.3)
    fig.tight_layout()


def plot_sweep_n_customers(df_sweep: pd.DataFrame) -> None:
    """Grafiek 4: feasibility rate en gemiddelde BPA-marge vs. aantal klanten."""
    fig, ax1 = plt.subplots(figsize=(10, 5))
    color_feas  = '#1f77b4'
    color_marge = '#d62728'
    color_cust  = '#ff7f0e'

    x = df_sweep['n_customers']

    ax1.plot(x, df_sweep['feasibility_rate'] * 100,
             color=color_feas, linewidth=2.5, marker='o', label='Feasibility rate (%)')
    ax1.plot(x, df_sweep['bpa_profitable_rate'] * 100,
             color='#2ca02c', linewidth=1.8, marker='^', linestyle='--',
             label='BPA winstgevend rate (%)')
    ax1.plot(x, df_sweep['all_cust_benefit_rate'] * 100,
             color=color_cust, linewidth=1.8, marker='s', linestyle=':',
             label='Alle klanten profiteren rate (%)')
    ax1.set_xlabel('Aantal klanten (subscriptions)', fontsize=12)
    ax1.set_ylabel('Rate (%)', fontsize=12)
    ax1.set_ylim(0, 105)
    ax1.set_xticks(x.tolist())

    ax2 = ax1.twinx()
    ax2.plot(x, df_sweep['avg_bpa_margin'],
             color=color_marge, linewidth=2, linestyle='--', marker='D',
             label='Gem. BPA-marge (€)')
    ax2.axhline(0, color='black', linewidth=0.8, linestyle=':')
    ax2.set_ylabel('Gemiddelde BPA-marge (€)', color=color_marge, fontsize=12)
    ax2.tick_params(axis='y', labelcolor=color_marge)
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'€{x:,.0f}'))

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='center right', fontsize=9)
    ax1.set_title(
        f'Feasibility vs. aantal klanten\n'
        f'(SL={SERVICE_LEVEL:.0%}, \u03b1={ALPHA:.2%}, '
        f'{N_RUNS} runs per punt, \u03ba_BPA={KAPPA_BPA:.0%})',
        fontsize=12
    )
    ax1.grid(True, alpha=0.3)
    fig.tight_layout()




if __name__ == '__main__':

    # ── 1. Data laden ──────────────────────────────────────────────────────
    print('\nData laden...')
    parts = load_parts()
    print(f'{len(parts)} onderdelen geladen (filters: ABC={FILTER_ABC}, '
          f'VP≥€{FILTER_MIN_VP}, n_cust≥{FILTER_MIN_KLANTEN})\n')

    # ── 2. Simulatie bij standaard N_CUSTOMERS ─────────────────────────────
    print(f'Simulatie starten: {N_RUNS} runs × {N_CUSTOMERS} klanten '
          f'(SL={SERVICE_LEVEL:.0%}, α={ALPHA:.2%})...')
    df_runs = run_simulation(parts, verbose=True, print_base_stock=True)
    print_simulation_summary(df_runs, N_CUSTOMERS)

    # ── 3. Sweep: feasibility vs. α ────────────────────────────────────────
    print(f'Sweep over α: {SWEEP_ALPHA}...')
    df_sweep_price = sweep_alpha(parts)

    print(f'\n{"═" * 70}')
    print('  SWEEP: FEASIBILITY VS. ABONNEMENTSPRIJS (α)')
    print(f'{"═" * 70}')
    print(df_sweep_price.to_string(index=False,
          formatters={
              'alpha':                 '{:.2%}'.format,
              'feasibility_rate':      '{:.1%}'.format,
              'bpa_profitable_rate':   '{:.1%}'.format,
              'all_cust_benefit_rate': '{:.1%}'.format,
              'avg_bpa_margin':        '€{:.2f}'.format,
              'avg_bpa_costs':         '€{:.2f}'.format,
              'avg_revenue':           '€{:.2f}'.format,
          }))

    # ── 4. Sweep: feasibility vs. n_customers ──────────────────────────────
    print(f'\nSweep over n_customers: {SWEEP_N_CUSTOMERS}...')
    df_sweep_n = sweep_n_customers(parts)

    print(f'\n{"═" * 70}')
    print('  SWEEP: FEASIBILITY VS. AANTAL KLANTEN')
    print(f'{"═" * 70}')
    print(df_sweep_n.to_string(index=False,
          formatters={
              'n_customers':           '{:.0f}'.format,
              'feasibility_rate':      '{:.1%}'.format,
              'bpa_profitable_rate':   '{:.1%}'.format,
              'all_cust_benefit_rate': '{:.1%}'.format,
              'avg_bpa_margin':        '€{:.2f}'.format,
              'avg_bpa_costs':         '€{:.2f}'.format,
              'avg_revenue':           '€{:.2f}'.format,
          }))

    # ── 5. Grafieken ───────────────────────────────────────────────────────
    plot_simulation(df_runs, N_CUSTOMERS)
    plot_sweep_price(df_sweep_price)
    plot_sweep_n_customers(df_sweep_n)
    plt.show()
