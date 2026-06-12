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

# Tabs voor de subscriptie-simulatie (binomiale adoptie per component/klant).
SHEET_ADOPTIE      = 'Adoptie'
SHEET_BESTELLINGEN = 'bestellingen_per_klant'

# Selectiebestand geproduceerd door classificatie_scoring.py.
# Aanwezigheid activeert de classificatie-koppeling (whitelist + LT-bron).
SELECTIE_PATH = os.path.join(SCRIPT_DIR, 'bpa_selectie.json')

# ── Filters (zelfde als base_stock_overview.py) ───────────────────────────────
FILTER_ARTICLE_TYPES = ['Critical', 'Onbekend']
FILTER_ABC           = ['A']
FILTER_INCOURANT     = ['Uitgeschakeld']
FILTER_MIN_VP        = 1000
FILTER_MIN_KLANTEN   = 5

# Standaard serviceniveaus voor het overzicht
SERVICE_LEVELS = [0.980, 0.990, 0.995, 0.999]

# Standaard aantal subscripties (globale fallback)
DEFAULT_N_KLANTEN = 20

# Default MTBF (jaren) wanneer MTBF onbekend is in zowel Excel als classificatie.
# λ valt dan terug op n_klanten / DEFAULT_MTBF_JR i.p.v. historische orders.
DEFAULT_MTBF_JR = 10.0


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
    # Sla eerst een snapshot van de HUIDIGE (old) staat op vóór de nieuwe
    # config wordt weggeschreven, zodat de Δ-kolommen in het overzicht de
    # werkelijke wijziging tonen en niet altijd 0 zijn.
    try:
        oude_cfg = laad_config()
        _sla_history_snapshot(oude_cfg)
    except Exception:
        pass

    cfg['aangepast'] = str(date.today())
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    print(f"  ✓ Configuratie opgeslagen in {CONFIG_PATH}")


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


def laad_excel_onderdelen(excel_file=None, *, skip_hard_filter: bool = True) -> pd.DataFrame:
    """Laad onderdelen uit de Excel-spreadsheet.
    excel_file: bestandspad (str) of file-like object (BytesIO / UploadedFile).
    Valt terug op EXCEL_PATH als niet opgegeven.
    skip_hard_filter: standaard True — alle BPA-hardfilters (ABC/VP/n_cust/...)
    worden overgeslagen. De selectie wordt volledig bepaald door de
    classificatie-tool. Zet expliciet op False om de oude hardfilter toe te passen."""
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
    if skip_hard_filter:
        mask = pd.Series(True, index=df.index)
    else:
        mask = (
            df['ArticleType'].isin(FILTER_ARTICLE_TYPES) &
            df['ABC_categorie'].isin(FILTER_ABC) &
            df['Incourant'].isin(FILTER_INCOURANT) &
            (df['VP'] >= FILTER_MIN_VP) &
            (df['n_cust'] >= FILTER_MIN_KLANTEN)
        )
    return df[mask].set_index('Code')[['Descr', 'VP', 'IP', 'Totaal_orders_5jr', 'n_cust', 'LT_days', 'MTBF']]


def _lambda_voor_rij(row, n_klanten: int) -> float:
    """λ = n_klanten / MTBF(jaar).

    Fallback: als MTBF ontbreekt of ≤ 0, gebruik DEFAULT_MTBF_JR (= 10 jr)
    in plaats van de historische orders-berekening. Dit geeft een
    voorspelbare λ ook voor componenten zonder MTBF-data.
    """
    mtbf = row['MTBF']
    if pd.notna(mtbf) and mtbf > 0:
        return n_klanten / mtbf
    return n_klanten / DEFAULT_MTBF_JR


# ══════════════════════════════════════════════════════════════════════════════
#  CLASSIFICATIE-KOPPELING
# ══════════════════════════════════════════════════════════════════════════════

def laad_classificatie_selectie() -> dict:
    """
    Lees bpa_selectie.json (geproduceerd door classificatie_scoring.py).

    Returns:
        dict met:
          'items'       : {code(str) → {'score','lt_dagen','lt_bron','abc'}}
          'gegenereerd' : timestamp-string of None
          'lt_overzicht': samenvatting of {}
        Lege dict als het bestand niet bestaat of corrupt is.
    """
    if not os.path.exists(SELECTIE_PATH):
        return {}
    try:
        with open(SELECTIE_PATH, encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"  ⚠  bpa_selectie.json niet leesbaar: {e}")
        return {}
    return {
        'items':        {str(it['code']): it for it in data.get('items', [])},
        'gegenereerd':  data.get('gegenereerd'),
        'lt_overzicht': data.get('lt_overzicht', {}),
        'threshold':    data.get('threshold'),
    }


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

    standaard_n = cfg['standaard_n_klanten']
    uitgesloten = set(cfg.get('uitgesloten_componenten', []))
    ip_ov = cfg.get('ip_overrides', {})
    lt_ov = cfg.get('lt_overrides', {})

    # ── Classificatie-koppeling ───────────────────────────────────────────
    # Als bpa_selectie.json bestaat: gebruik als whitelist (alleen die codes
    # uit de Excel) en haal LT-bron + classificatie-score op. Bij actieve
    # whitelist slaan we de BPA-hardfilter over zodat ALLE "Opnemen"-codes
    # zichtbaar worden (ook B/C-categorieën of items met lage VP).
    _cls = laad_classificatie_selectie()
    _cls_items = _cls.get('items', {})
    _gebruik_cls_whitelist = bool(_cls_items)
    if _gebruik_cls_whitelist:
        print(f"  ✓ Classificatie-selectie actief: {len(_cls_items)} codes "
              f"(gegenereerd {_cls.get('gegenereerd')})")

    # 1. Excel-onderdelen — hardfilter is altijd uit; selectie loopt via classificatie.
    try:
        excel_parts = laad_excel_onderdelen(excel_file, skip_hard_filter=True)
    except Exception as e:
        print(f"  ⚠  Excel niet geladen: {e}")
        excel_parts = pd.DataFrame()

    for code, row in excel_parts.iterrows():
        if str(code) in uitgesloten:
            continue
        # Whitelist: sla over als niet in classificatie-selectie
        if _gebruik_cls_whitelist and str(code) not in _cls_items:
            continue
        n     = cfg['n_klanten_overrides'].get(str(code), standaard_n)
        # VP komt uit Excel (standaard verkoopprijs); IP = VP / 2 tenzij override.
        vp    = float(row.get('VP', 0.0) or 0.0)
        ip    = ip_ov.get(str(code), vp / 2)
        # LT: configuratie-override telt als 'bevestigd' (gebruiker heeft 'm
        # zelf gezet); anders: bron volgens classificatie; anders Excel-waarde.
        if str(code) in lt_ov:
            lt_d    = lt_ov[str(code)]
            lt_bron = 'override'
        else:
            lt_d    = int(row['LT_days'])
            lt_bron = _cls_items.get(str(code), {}).get('lt_bron', 'onbekend')
        # LT=0 → vul aan met 30 dagen en markeer met aparte bron (blauwe cel in UI)
        if lt_d == 0 and lt_bron != 'override':
            lt_d = 30
            lt_bron = 'nul→30'
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
            'LT_bron':   lt_bron,
            'Cls_score': _cls_items.get(str(code), {}).get('score'),
            'IP':        round(ip, 2),
            'VP':        round(vp, 2),
            'mu':        round(mu, 4),
            'bron':      'excel',
        }
        for sl in SERVICE_LEVELS:
            rij[f's@{sl:.1%}'] = stocks[sl]
        rijen.append(rij)

    # ── Diagnostiek: classificatie-codes die niet in de BPA-Excel staan ──
    if _gebruik_cls_whitelist:
        _excel_codes = {str(c) for c in excel_parts.index} if not excel_parts.empty else set()
        _missing = [c for c in _cls_items.keys() if c not in _excel_codes and c not in uitgesloten]
        if _missing:
            print(f"  ⚠  {len(_missing)} classificatie-codes ontbreken in de BPA-Excel "
                  f"(eerste 5: {_missing[:5]}) — toegevoegd vanuit classificatie-metadata")

        # Bouw synthetische rijen vanuit de classificatie-payload (descr, ip, vp,
        # mtbf, totaal_orders_5jr, n_cust, lt_dagen). Zo verschijnen ALLE
        # "Opnemen in lijst"-codes in het overzicht.
        for code in _missing:
            cls_meta = _cls_items.get(code, {})
            n   = cfg['n_klanten_overrides'].get(code, standaard_n)
            # VP komt uit classificatie-metadata; IP = VP/2 (of override).
            vp_meta = cls_meta.get('vp')
            vp  = float(vp_meta) if vp_meta is not None else 0.0
            ip  = ip_ov.get(code, vp / 2)

            # LT: override > classificatie > default 30
            if code in lt_ov:
                lt_d    = int(lt_ov[code])
                lt_bron = 'override'
            else:
                lt_raw  = cls_meta.get('lt_dagen')
                if lt_raw is None or int(lt_raw) == 0:
                    lt_d    = 30
                    lt_bron = 'nul→30'
                else:
                    lt_d    = int(lt_raw)
                    lt_bron = cls_meta.get('lt_bron', 'onbekend')

            # Lambda: MTBF heeft voorrang; bij ontbreken valt λ terug op
            # n / DEFAULT_MTBF_JR (= 10 jr). Geen fallback meer op historische
            # orders, zodat het gedrag consistent is met _lambda_voor_rij().
            mtbf = cls_meta.get('mtbf')
            if mtbf is not None and mtbf > 0:
                lam = n / mtbf
            else:
                lam = n / DEFAULT_MTBF_JR

            lt_jr = lt_d / 365
            mu    = lam * lt_jr
            stocks = {
                sl: BPAOptimizationModel.inverse_service_level(sl, lam, lt_jr)
                for sl in SERVICE_LEVELS
            }
            rij = {
                'Code':      code,
                'Descr':     str(cls_meta.get('descr', ''))[:40],
                'n_klanten': n,
                'lambda_jr': round(lam, 4),
                'LT_dagen':  lt_d,
                'LT_bron':   lt_bron,
                'Cls_score': cls_meta.get('score'),
                'IP':        round(ip, 2),
                'VP':        round(vp, 2),
                'mu':        round(mu, 4),
                'bron':      'classificatie',
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
            'LT_bron':   'handmatig',
            'Cls_score': None,
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
#  SUBSCRIPTIE-SIMULATIE  (binomiale adoptie per component)
# ══════════════════════════════════════════════════════════════════════════════

def laad_adoptie_data(excel_file=None) -> pd.DataFrame:
    """
    Laad de 'Adoptie'-tab met per (component, klant)-combinatie het aantal
    orders en de adoption rate.

    excel_file: bestandspad (str) of file-like object; valt terug op EXCEL_PATH.

    Returns een DataFrame met kolommen:
        Code, Klant, Orders_component_klant, Orders_klant_totaal, Adoption_rate
    """
    bron = excel_file if excel_file is not None else EXCEL_PATH
    # Reset de leespositie van een geüploade buffer: deze functie kan meerdere
    # keren op hetzelfde file-like object worden aangeroepen (bv. runs +
    # metadata), waardoor een eerder gelezen buffer aan het einde zou staan.
    if excel_file is not None:
        try:
            bron.seek(0)
        except (AttributeError, ValueError):
            pass
    df = pd.read_excel(bron, sheet_name=SHEET_ADOPTIE)
    df = df.rename(columns={
        'Verkooporderregel artikel.Artikel.Artikelcode': 'Code',
        'Order.Relatie':                                 'Klant',
        'Aantal_orders_5jr':                             'Orders_component_klant',
        'Totaal_bestellingen_klant_5jr':                 'Orders_klant_totaal',
        'Adoption_rate':                                 'Adoption_rate',
    })
    df['Code']  = df['Code'].astype(str).str.strip()
    df['Klant'] = df['Klant'].astype(str).str.strip()
    df['Adoption_rate'] = (
        pd.to_numeric(df['Adoption_rate'], errors='coerce').fillna(0.0).clip(0.0, 1.0)
    )
    df['Orders_component_klant'] = (
        pd.to_numeric(df['Orders_component_klant'], errors='coerce')
        .fillna(0).clip(lower=0).astype(int)
    )
    df['Orders_klant_totaal'] = (
        pd.to_numeric(df['Orders_klant_totaal'], errors='coerce')
        .fillna(0).clip(lower=0).astype(int)
    )
    return df[['Code', 'Klant', 'Orders_component_klant',
               'Orders_klant_totaal', 'Adoption_rate']]


def classificatie_codes() -> set:
    """Geef de set artikelcodes uit de classificatie-selectie (bpa_selectie.json).

    Lege set als er geen selectiebestand is. Wordt gebruikt om de
    subscriptie-simulatie standaard te beperken tot de componenten die uit
    de classificatie komen.
    """
    sel = laad_classificatie_selectie()
    return {str(c).strip() for c in sel.get('items', {}).keys()}


def simuleer_subscripties_per_component(
    n_runs:      int   = 500,
    seed:        int   = 42,
    excel_file         = None,
    codes              = None,
) -> pd.DataFrame:
    """
    Monte Carlo simulatie van het aantal subscripties per component.

    Per run en per (component, klant) wordt een binomiale trekking gedaan over
    de orders van de klant voor dat component; de klant neemt een subscriptie
    als er minstens één "succes" is (binaire uitkomst):

        K_klant   ~ Binomiaal(n = orders van de klant voor het component,
                              p = adoption rate van de klant)
        X_klant   = 1 als K_klant >= 1, anders 0

    Hierdoor neemt de kans dat een klant een subscriptie neemt toe naarmate de
    klant meer orders voor het component heeft:
        P(X_klant = 1) = 1 - (1 - adoption_rate) ** aantal_orders.
    Het aantal subscripties voor een component in een run is de som van deze
    binaire beslissingen over alle klanten (maximaal het aantal klanten). Over
    alle runs wordt de verdeling samengevat.

    Parameters
    ----------
    n_runs : aantal Monte Carlo runs (stochastische trekkingen).
    seed   : seed voor reproduceerbaarheid (None = willekeurig).
    excel_file : bestandspad of file-like object; valt terug op EXCEL_PATH.
    codes : optionele iterable van artikelcodes om de simulatie te beperken.
            Bij None worden standaard alleen de classificatie-componenten
            (bpa_selectie.json) gesimuleerd.

    Returns
    -------
    DataFrame geïndexeerd op Code met kolommen:
        gem_subs, std_subs, p05, p50, p95, min_subs, max_subs,
        n_klanten, n_orders, runs
    """
    runs = simuleer_subscripties_runs(
        n_runs=n_runs, seed=seed, excel_file=excel_file, codes=codes,
    )
    if not runs:
        return pd.DataFrame()

    adoptie = laad_adoptie_data(excel_file)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    n_klanten = adoptie.groupby('Code').size()
    n_orders  = adoptie.groupby('Code')['Orders_component_klant'].sum()

    resultaten: dict = {}
    for code, subs_per_run in runs.items():
        resultaten[code] = {
            'gem_subs':  float(subs_per_run.mean()),
            'std_subs':  float(subs_per_run.std(ddof=0)),
            'p05':       float(np.percentile(subs_per_run, 5)),
            'p50':       float(np.percentile(subs_per_run, 50)),
            'p95':       float(np.percentile(subs_per_run, 95)),
            'min_subs':  int(subs_per_run.min()),
            'max_subs':  int(subs_per_run.max()),
            'n_klanten': int(n_klanten.get(code, 0)),
            'n_orders':  int(n_orders.get(code, 0)),
            'runs':      int(n_runs),
        }
    out = pd.DataFrame.from_dict(resultaten, orient='index')
    out.index.name = 'Code'
    return out.sort_values('gem_subs', ascending=False)


def simuleer_subscripties_runs(
    n_runs:      int   = 500,
    seed:        int   = 42,
    excel_file         = None,
    codes              = None,
) -> dict:
    """
    Voer de subscriptie-simulatie uit en geef de ruwe trekkingen terug.

    Per (component, klant) wordt een binomiale trekking gedaan over de orders
    van de klant voor dat component; de klant neemt een subscriptie als er
    minstens één succes is:
        K_klant ~ Binomiaal(n = orders voor het component, p = adoption rate)
        X_klant = 1 als K_klant >= 1, anders 0.
    Het aantal subscripties per component is de som van deze binaire
    beslissingen over de klanten (maximaal het aantal klanten).

    Bij codes=None wordt standaard beperkt tot de classificatie-componenten.

    Returns een dict {code → np.ndarray van lengte n_runs} met per run het
    gesimuleerde aantal subscripties voor dat component. Handig voor het
    plotten van de verdeling (histogram) per component.
    """
    adoptie = laad_adoptie_data(excel_file)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    if adoptie.empty:
        return {}

    rng = np.random.default_rng(seed)
    runs: dict = {}
    for code, grp in adoptie.groupby('Code', sort=True):
        n_arr = grp['Orders_component_klant'].to_numpy()
        p_arr = grp['Adoption_rate'].to_numpy()
        # Binomiale trekking per klant over zijn orders voor dit component.
        # Elke order is een trial met succeskans = adoption rate. De klant neemt
        # een subscriptie (binair) als er minstens één succes is. Zo stijgt de
        # kans op een subscriptie met het aantal orders van de klant.
        # (n_runs, n_klanten) binomiale trekkingen → binariseren → som per run.
        draws = rng.binomial(n_arr[None, :], p_arr[None, :],
                             size=(n_runs, n_arr.shape[0]))
        runs[code] = (draws >= 1).sum(axis=1)
    return runs


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
