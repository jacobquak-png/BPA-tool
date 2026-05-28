# -*- coding: utf-8 -*-
"""
BPA Jaarlijks Beheer Tool
=========================
Eenmaal per jaar kan BPA via dit script:
  1. Het aantal subscripties (klanten) per component aanpassen.
  2. Nieuwe componenten toevoegen met bijbehorende lambda en levertijd.
  3. Een overzicht opslaan van de bijgewerkte basisvoorraden.

Instellingen worden opgeslagen in bpa_config.json naast dit script.
Bij de volgende run worden die automatisch geladen.

Gebruik:
    python bpa_beheer.py
"""

import json
import os
import sys
from datetime import date

import numpy as np
import pandas as pd

# ── pad zoeken naar model.py ──────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from model import BPAOptimizationModel

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ══════════════════════════════════════════════════════════════════════════════
#  PADEN
# ══════════════════════════════════════════════════════════════════════════════

SCRIPT_DIR  = os.path.dirname(__file__)
CONFIG_PATH   = os.path.join(SCRIPT_DIR, 'bpa_config.json')
HISTORY_PATH  = os.path.join(SCRIPT_DIR, 'bpa_history.json')
EXCEL_PATH  = os.path.join(SCRIPT_DIR,
                           'annual_use_abc_met_artikeldata_complete_europa.xlsx')
SHEET_NAME  = 'Filtered '

# ── Filters (zelfde als base_stock_overview.py) ───────────────────────────────
FILTER_ARTICLE_TYPES = ['Critical', 'Onbekend']
FILTER_ABC           = ['A']
FILTER_INCOURANT     = ['Uitgeschakeld']
FILTER_MIN_VP        = 1000
FILTER_MIN_KLANTEN   = 5

# Standaard serviceniveaus voor het overzicht
SERVICE_LEVELS = [0.98, 0.990, 0.995, 0.999]

# Standaard aantal subscripties (globale fallback)
DEFAULT_N_KLANTEN = 20


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG – laden / opslaan
# ══════════════════════════════════════════════════════════════════════════════

def _leeg_config() -> dict:
    return {
        "aangemaakt":          str(date.today()),
        "aangepast":           str(date.today()),
        "standaard_n_klanten": DEFAULT_N_KLANTEN,
        # code → aantal subscripties (overschrijft het standaard getal)
        "n_klanten_overrides": {},
        # code → inkoopprijs override (€)
        "ip_overrides": {},
        # code → levertijd override (dagen)
        "lt_overrides": {},
        # handmatig toegevoegde componenten
        # code → {"descr": str, "lambda_per_jaar": float, "lt_dagen": int,
        #          "n_klanten": int, "ip": float}
        "handmatige_componenten": {},
        # codes die volledig uit het model worden uitgesloten (Excel én handmatig)
        "uitgesloten_componenten": [],
    }


def laad_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding='utf-8') as f:
            cfg = json.load(f)
        # voeg ontbrekende sleutels toe bij oudere bestanden
        basis = _leeg_config()
        for k, v in basis.items():
            cfg.setdefault(k, v)
        return cfg
    return _leeg_config()


def sla_config_op(cfg: dict) -> None:
    cfg['aangepast'] = str(date.today())
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    print(f"  ✓ Configuratie opgeslagen in {CONFIG_PATH}")
    _sla_history_snapshot(cfg)


def _sla_history_snapshot(cfg: dict) -> None:
    """Voeg een snapshot toe aan bpa_history.json na elke config-update."""
    try:
        df = bereken_overzicht(cfg)
    except Exception:
        return
    if df.empty:
        return

    sl_cols = [c for c in df.columns if c.startswith('s@')]
    snapshot = {
        'datum':      str(date.today()),
        'n_klanten':  cfg['standaard_n_klanten'],
        'n_actief':   len(df),
        'totalen':    {c: int(df[c].sum()) for c in sl_cols},
        'componenten': {
            str(code): {c: int(df.at[code, c]) for c in sl_cols}
            for code in df.index
        },
    }

    history = []
    if os.path.exists(HISTORY_PATH):
        with open(HISTORY_PATH, encoding='utf-8') as f:
            try:
                history = json.load(f)
            except json.JSONDecodeError:
                history = []

    # Vervang bestaande snapshot van vandaag (meest recente wint)
    history = [h for h in history if h.get('datum') != snapshot['datum']]
    history.append(snapshot)
    history.sort(key=lambda h: h['datum'])

    with open(HISTORY_PATH, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
#  DATA LADEN UIT EXCEL
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


def laad_excel_onderdelen(excel_file=None) -> pd.DataFrame:
    """Laad gefilterde onderdelen uit de Excel-spreadsheet.
    excel_file: bestandspad (str) of file-like object (BytesIO / UploadedFile).
    Valt terug op EXCEL_PATH als niet opgegeven."""
    bron = excel_file if excel_file is not None else EXCEL_PATH
    df = pd.read_excel(bron, sheet_name=SHEET_NAME)
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
    return df[mask].set_index('Code')[['Descr', 'VP', 'IP', 'Totaal_orders_5jr', 'n_cust', 'LT_days', 'MTBF']]


def _lambda_voor_rij(row, n_klanten: int) -> float:
    """λ = n_klanten / MTBF(jaar); fallback op historisch gebruik."""
    mtbf = row['MTBF']
    if pd.notna(mtbf) and mtbf > 0:
        return n_klanten / mtbf
    return row['Totaal_orders_5jr'] / 5


# ══════════════════════════════════════════════════════════════════════════════
#  BEREKEN OVERZICHT
# ══════════════════════════════════════════════════════════════════════════════

def bereken_overzicht(cfg: dict, excel_file=None) -> pd.DataFrame:
    """
    Geeft een DataFrame terug met per component:
      Code, Descr, n_klanten, lambda_per_jaar, LT_dagen, mu, en s per serviceniveau.
    Combineert Excel-onderdelen met handmatige toevoegingen.
    excel_file: optioneel bestandspad of file-like object; valt terug op EXCEL_PATH.
    """
    rijen = []

    # 1. Excel-onderdelen
    try:
        excel_parts = laad_excel_onderdelen(excel_file)
    except Exception as e:
        print(f"  ⚠  Excel niet geladen: {e}")
        excel_parts = pd.DataFrame()

    standaard_n = cfg['standaard_n_klanten']

    uitgesloten = set(cfg.get('uitgesloten_componenten', []))

    ip_ov = cfg.get('ip_overrides', {})
    lt_ov = cfg.get('lt_overrides', {})

    for code, row in excel_parts.iterrows():
        if str(code) in uitgesloten:
            continue
        n     = cfg['n_klanten_overrides'].get(str(code), standaard_n)
        ip    = ip_ov.get(str(code), row['IP'])
        lt_d  = lt_ov.get(str(code), int(row['LT_days']))
        lam   = _lambda_voor_rij(row, n)
        lt_jr = lt_d / 365
        mu    = lam * lt_jr
        stocks = {
            sl: BPAOptimizationModel.inverse_service_level(sl, lam, lt_jr)
            for sl in SERVICE_LEVELS
        }
        rij = {
            'Code':      code,
            'Descr':     str(row['Descr'])[:40],
            'n_klanten': n,
            'lambda_jr': round(lam, 4),
            'LT_dagen':  lt_d,
            'IP':        round(ip, 2),
            'VP':        round(ip * 2, 2),
            'mu':        round(mu, 4),
            'bron':      'excel',
        }
        for sl in SERVICE_LEVELS:
            rij[f's@{sl:.1%}'] = stocks[sl]
        rijen.append(rij)

    # 2. Handmatige componenten
    for code, hcomp in cfg['handmatige_componenten'].items():
        if code in uitgesloten:
            continue
        lam   = hcomp['lambda_per_jaar']
        lt_d  = hcomp['lt_dagen']
        lt_jr = lt_d / 365
        mu    = lam * lt_jr
        n     = hcomp.get('n_klanten', standaard_n)
        ip    = hcomp.get('ip', 0.0)
        stocks = {
            sl: BPAOptimizationModel.inverse_service_level(sl, lam, lt_jr)
            for sl in SERVICE_LEVELS
        }
        rij = {
            'Code':      code,
            'Descr':     hcomp.get('descr', '')[:40],
            'n_klanten': n,
            'lambda_jr': round(lam, 4),
            'LT_dagen':  lt_d,
            'IP':        round(ip, 2),
            'VP':        round(ip * 2, 2),
            'mu':        round(mu, 4),
            'bron':      'handmatig',
        }
        for sl in SERVICE_LEVELS:
            rij[f's@{sl:.1%}'] = stocks[sl]
        rijen.append(rij)

    return pd.DataFrame(rijen).set_index('Code') if rijen else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
#  KOSTENMODEL
# ══════════════════════════════════════════════════════════════════════════════

def bouw_model_kosten(
    overzicht_df: pd.DataFrame,
    alpha:         float,
    kappa_bpa:     float,
    kappa_c:       float,
    service_level: float,
    n_klanten_override=None,   # int | None – overschrijft alle n_klanten en schaalt lambda_jr
) -> tuple:
    """
    Bouwt een BPAOptimizationModel vanuit de overzicht DataFrame en geeft
    (model, result) terug.  Demand-keys: (part, customer).
    Elk component krijgt zijn eigen n_klanten; de klantenpool is de unie
    van alle klanten (grootte = max(n_klanten)).
    n_klanten_override: overschrijft n_klanten voor alle componenten en
    schaalt lambda_jr proportioneel (Λ_new = N_new · λ_per_cust).
    """
    if overzicht_df.empty:
        raise ValueError("overzicht_df is leeg – laad eerst het overzicht.")

    df = overzicht_df.copy()
    if n_klanten_override is not None:
        _n_new = int(n_klanten_override)
        df['lambda_jr'] = df.apply(
            lambda r: _n_new * (r['lambda_jr'] / r['n_klanten'])
            if r['n_klanten'] > 0 else r['lambda_jr'],
            axis=1,
        )
        df['n_klanten'] = _n_new

    n_max = int(df['n_klanten'].max())
    customers = {f'Klant_{i+1}': 1 for i in range(n_max)}

    demand_data: dict = {}
    for code, row in df.iterrows():
        n_i = int(row['n_klanten'])
        lam_per_cust = row['lambda_jr'] / n_i if n_i > 0 else row['lambda_jr']
        for i in range(n_max):
            demand_data[(code, f'Klant_{i+1}')] = lam_per_cust if i < n_i else 0.0

    vp_col = 'VP' if 'VP' in df.columns else 'IP'
    cost_data = {
        'purchase_price': {code: row['IP']      for code, row in df.iterrows()},
        'sales_price':    {code: row[vp_col]    for code, row in df.iterrows()},
        'lead_time':      {code: row['LT_dagen'] / 365 for code, row in df.iterrows()},
    }

    model = BPAOptimizationModel()
    model.add_customers_and_machines(customers)
    model.add_spare_parts(list(df.index))
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
#  MENU-FUNCTIES
# ══════════════════════════════════════════════════════════════════════════════

def menu_toon_overzicht(cfg: dict) -> None:
    print("\nOverzicht berekenen…")
    df = bereken_overzicht(cfg)
    if df.empty:
        print("  Geen onderdelen gevonden.")
        return

    sl_cols = [c for c in df.columns if c.startswith('s@')]
    hdr = f"{'Code':<22}  {'Omschrijving':<40}  {'N':>4}  {'λ/jr':>8}  {'LT(d)':>6}  {'μ':>7}"
    for c in sl_cols:
        hdr += f"  {c:>8}"
    hdr += f"  {'bron':<10}"
    print(hdr)
    print('─' * len(hdr))

    totals = {c: 0 for c in sl_cols}
    for code, row in df.iterrows():
        line = (f"{code:<22}  {row['Descr']:<40}  {int(row['n_klanten']):>4}  "
                f"{row['lambda_jr']:>8.4f}  {int(row['LT_dagen']):>6}  {row['mu']:>7.4f}")
        for c in sl_cols:
            val = int(row[c])
            line += f"  {val:>8}"
            totals[c] += val
        line += f"  {row['bron']:<10}"
        print(line)

    print('─' * len(hdr))
    totline = f"{'TOTAAL':<22}  {'':40}  {'':>4}  {'':>8}  {'':>6}  {'':>7}"
    for c in sl_cols:
        totline += f"  {totals[c]:>8}"
    print(totline)

    # Optioneel opslaan als CSV
    antw = input("\nCSV opslaan? (j/n) [n]: ").strip().lower()
    if antw == 'j':
        pad = os.path.join(SCRIPT_DIR, f"bpa_overzicht_{date.today()}.csv")
        df.to_csv(pad, sep=';', decimal=',')
        print(f"  ✓ Opgeslagen als {pad}")


def menu_pas_subscripties_aan(cfg: dict) -> None:
    print(f"\nHuidige standaard: {cfg['standaard_n_klanten']} subscripties")
    antw = input("Nieuw standaard aantal subscripties (Enter = niet wijzigen): ").strip()
    if antw:
        try:
            cfg['standaard_n_klanten'] = int(antw)
            print(f"  ✓ Standaard aangepast naar {cfg['standaard_n_klanten']}")
        except ValueError:
            print("  ✗ Ongeldige invoer.")

    cfg.setdefault('ip_overrides', {})
    cfg.setdefault('lt_overrides', {})

    print("\nOverride per component – N, inkoopprijs en levertijd (laat code leeg om te stoppen):")
    print("  Laat een veld leeg om het ongewijzigd te laten. Typ 'x' om een override te verwijderen.")
    while True:
        code = input("  Artikelcode (of Enter om te stoppen): ").strip()
        if not code:
            break

        # N
        huidig_n = cfg['n_klanten_overrides'].get(code, cfg['standaard_n_klanten'])
        antw = input(f"    N subscripties (huidig: {huidig_n}, 'x'=verwijder): ").strip()
        if antw.lower() == 'x':
            cfg['n_klanten_overrides'].pop(code, None)
            print(f"    ✓ N-override verwijderd.")
        elif antw:
            try:
                cfg['n_klanten_overrides'][code] = int(antw)
                print(f"    ✓ N → {int(antw)}")
            except ValueError:
                print("    ✗ Ongeldige invoer, overgeslagen.")

        # Inkoopprijs
        huidig_ip = cfg['ip_overrides'].get(code, '(uit Excel)')
        antw = input(f"    Inkoopprijs € (huidig: {huidig_ip}, 'x'=verwijder): ").strip().replace(',', '.')
        if antw.lower() == 'x':
            cfg['ip_overrides'].pop(code, None)
            print(f"    ✓ IP-override verwijderd.")
        elif antw:
            try:
                cfg['ip_overrides'][code] = float(antw)
                print(f"    ✓ IP → €{float(antw):.2f}")
            except ValueError:
                print("    ✗ Ongeldige invoer, overgeslagen.")

        # Levertijd
        huidig_lt = cfg['lt_overrides'].get(code, '(uit Excel)')
        antw = input(f"    Levertijd dagen (huidig: {huidig_lt}, 'x'=verwijder): ").strip()
        if antw.lower() == 'x':
            cfg['lt_overrides'].pop(code, None)
            print(f"    ✓ LT-override verwijderd.")
        elif antw:
            try:
                cfg['lt_overrides'][code] = int(antw)
                print(f"    ✓ LT → {int(antw)} dagen")
            except ValueError:
                print("    ✗ Ongeldige invoer, overgeslagen.")

    sla_config_op(cfg)


def menu_voeg_component_toe(cfg: dict) -> None:
    print("\n── Nieuw component toevoegen ──")
    code = input("  Artikelcode: ").strip()
    if not code:
        print("  Geannuleerd.")
        return
    if code in cfg['handmatige_componenten']:
        print(f"  ⚠  {code} bestaat al. Gebruik 'Subscripties aanpassen' of verwijder het eerst.")
        return

    descr = input("  Omschrijving: ").strip()
    try:
        lam   = float(input("  Lambda (vraag per jaar): ").replace(',', '.'))
        lt    = int(input("  Levertijd leverancier→BPA (dagen): "))
        n     = int(input(f"  Aantal subscripties [{cfg['standaard_n_klanten']}]: ").strip()
                   or cfg['standaard_n_klanten'])
        ip_raw = input("  Inkoopprijs (optioneel, bijv. 1234.56): ").strip().replace(',', '.')
        ip    = float(ip_raw) if ip_raw else 0.0
    except ValueError:
        print("  ✗ Ongeldige invoer, component niet toegevoegd.")
        return

    cfg['handmatige_componenten'][code] = {
        'descr':           descr,
        'lambda_per_jaar': lam,
        'lt_dagen':        lt,
        'n_klanten':       n,
        'ip':              ip,
    }
    print(f"  ✓ Component '{code}' toegevoegd.")
    sla_config_op(cfg)


def menu_verwijder_component(cfg: dict) -> None:
    """Verwijder een component uit het model (handmatig of Excel)."""
    handmatig  = cfg['handmatige_componenten']
    uitgesloten = cfg.setdefault('uitgesloten_componenten', [])

    # Laad Excel-codes voor weergave
    try:
        excel_codes = list(laad_excel_onderdelen().index)
    except FileNotFoundError:
        excel_codes = []

    alle_codes = list(handmatig.keys()) + [c for c in excel_codes if c not in handmatig]
    actief     = [c for c in alle_codes if c not in uitgesloten]

    if not actief:
        print("\n  Geen actieve componenten om te verwijderen.")
        return

    print("\nActieve componenten:")
    for c in actief:
        bron = 'handmatig' if c in handmatig else 'excel'
        descr = handmatig[c].get('descr', '') if c in handmatig else ''
        print(f"  {c:<22}  [{bron}]  {descr}")

    code = input("Artikelcode om te verwijderen (Enter = annuleren): ").strip()
    if not code:
        return
    if code in handmatig:
        del handmatig[code]
        print(f"  ✓ '{code}' (handmatig) verwijderd.")
        sla_config_op(cfg)
    elif code in excel_codes:
        if code not in uitgesloten:
            uitgesloten.append(code)
        print(f"  ✓ '{code}' (Excel) uitgesloten van het model.")
        sla_config_op(cfg)
    else:
        print(f"  ✗ '{code}' niet gevonden.")


def menu_toon_config(cfg: dict) -> None:
    print(f"\nConfiguratiebestand: {CONFIG_PATH}")
    print(f"  Aangemaakt   : {cfg['aangemaakt']}")
    print(f"  Aangepast    : {cfg['aangepast']}")
    print(f"  Standaard N  : {cfg['standaard_n_klanten']}")
    overrides = cfg['n_klanten_overrides']
    if overrides:
        print(f"\n  Overrides ({len(overrides)}):")
        for code, n in overrides.items():
            print(f"    {code:<22}  →  {n} subscripties")
    else:
        print("  Geen per-component overrides.")
    handmatig = cfg['handmatige_componenten']
    if handmatig:
        print(f"\n  Handmatige componenten ({len(handmatig)}):")
        for code, v in handmatig.items():
            print(f"    {code:<22}  λ={v['lambda_per_jaar']:.4f}/jr  "
                  f"LT={v['lt_dagen']}d  N={v.get('n_klanten','std')}  "
                  f"IP=€{v.get('ip',0):.2f}  – {v.get('descr','')}")
    else:
        print("  Geen handmatige componenten.")


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT BIJ AFSLUITING
# ══════════════════════════════════════════════════════════════════════════════

def exporteer_afsluiting(cfg: dict) -> None:
    """Sla bij afsluiting automatisch een CSV op met alle componenten en basisvoorraden."""
    print("\nOverzicht opslaan…")
    df = bereken_overzicht(cfg)
    if df.empty:
        print("  Geen onderdelen om op te slaan.")
        return

    pad = os.path.join(SCRIPT_DIR, f"bpa_base_stock_{date.today()}.csv")
    df.to_csv(pad, sep=';', decimal=',')
    print(f"  ✓ Basisvoorraden opgeslagen in {pad}")
    print(f"     {len(df)} componenten  |  kolommen: {', '.join(df.columns.tolist())}")


# ══════════════════════════════════════════════════════════════════════════════
#  HOOFDMENU
# ══════════════════════════════════════════════════════════════════════════════

MENU = """
╔══════════════════════════════════════════════════╗
║         BPA Jaarlijks Beheer Tool                ║
╠══════════════════════════════════════════════════╣
║  1. Toon overzicht basisvoorraden                ║
║  2. Pas subscripties aan (standaard of per code) ║
║  3. Voeg nieuw component toe                     ║
║  4. Verwijder component uit model                ║
║  5. Toon huidige configuratie                    ║
║  0. Afsluiten                                    ║
╚══════════════════════════════════════════════════╝
Keuze: """


def main() -> None:
    cfg = laad_config()
    print(f"\nBPA Beheer Tool – configuratie geladen (aangepast: {cfg['aangepast']})")

    while True:
        keuze = input(MENU).strip()
        if keuze == '1':
            menu_toon_overzicht(cfg)
        elif keuze == '2':
            menu_pas_subscripties_aan(cfg)
        elif keuze == '3':
            menu_voeg_component_toe(cfg)
        elif keuze == '4':
            menu_verwijder_component(cfg)
        elif keuze == '5':
            menu_toon_config(cfg)
        elif keuze == '0':
            exporteer_afsluiting(cfg)
            print("Tot de volgende jaarlijkse update!")
            break
        else:
            print("  Ongeldige keuze, probeer opnieuw.")


if __name__ == '__main__':
    main()


