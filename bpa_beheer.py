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
        # code → aantal subscripties (overschrijft het aantal klantlocaties)
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
        'n_klanten':  int(df['n_klanten'].median()) if not df.empty else 0,
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


def _parse_lt_dagen(val) -> int:
    """Parseer levertijd in dagen, robuust voor int/float/strings.

    Spiegelt classificatie._parse_lt_dagen zodat de tool exact dezelfde
    levertijd overneemt als de classificatie. Accepteert numerieke cellen
    (30.0), '30', '30 dagen', NL-decimalen ('30,0') en waarden met een
    leidende spatie. Geeft 0 bij leeg/onparseerbaar; 0 wordt downstream
    aangevuld tot de default van 30 dagen (LT_bron 'nul→30').

    De oude implementatie (``int(str(v).split()[0])`` met
    ``str(v)[0].isdigit()``) faalde op float-cellen ('30.0' → ValueError,
    waardoor de héle Excel-load afbrak) en op NL-decimalen of een leidende
    spatie (→ stil 0), waardoor de juiste levertijd vaak niet meekwam.
    """
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return int(val) if val > 0 else 0
    s = str(val).strip()
    if not s:
        return 0
    head = s.split()[0].replace(',', '.')
    try:
        n = int(float(head))
        return n if n > 0 else 0
    except ValueError:
        return 0


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
    df['LT_days'] = df['_LT_raw'].apply(_parse_lt_dagen)
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


def _n_uit_klantlocaties(raw) -> int:
    """Aantal subscripties uit het werkelijke aantal klantlocaties (n_cust).

    Vervangt de oude globale standaardwaarde: bij ontbrekende of ongeldige
    data valt het aantal terug op 1 (minimaal één potentiële klant).
    """
    try:
        if raw is None or pd.isna(raw):
            return 1
        n = int(float(raw))
        return n if n >= 1 else 1
    except (TypeError, ValueError):
        return 1


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
        n     = cfg['n_klanten_overrides'].get(str(code),
                                               _n_uit_klantlocaties(row.get('n_cust')))
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
            n   = cfg['n_klanten_overrides'].get(
                code, _n_uit_klantlocaties(cls_meta.get('n_cust')))
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
        n     = hcomp.get('n_klanten', 1)
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

def regionale_adoptie_parameter(
    p0: float,
    alpha: float, X: float,
    alpha0: float, X0: float,
    gamma_alpha: float = 1.0, gamma_X: float = 0.5,
    alpha_max: float = None, s_alpha: float = 0.02,
) -> float:
    """Regionale adoptieparameter p_r(α, X) volgens de logit-specificatie.

    Het baseline-niveau p0 = p_{r,0} geldt bij de referentie-instelling
    (α0, X0). Prijs en service level schuiven de adoptie op de logit-schaal:

        p_r(α, X) = σ( logit(p0)
                       − γ_α · (α − α0) / α0
                       + γ_X · [g(X) − g(X0)] )

    met σ(y) = 1/(1+e^{−y}),  logit(p) = ln(p/(1−p))  en  g(X) = −ln(1−X).

    Een hogere prijs (α > α0) verlaagt de adoptie; een hoger service level
    (X > X0) verhoogt de adoptie. Bij (α0, X0) geldt p_r = p0. Het resultaat
    blijft begrensd tussen 0 en 1.

    WTP-plafond op α (= α_U = κ_c)
    ------------------------------
    Een klant abonneert alleen als α·U_i^c ≤ κ_c·U_i^c, dus α ≤ κ_c (U_i^c
    valt weg). Als ``alpha_max`` (= κ_c) is opgegeven, wordt de logit-kans
    vermenigvuldigd met een gladde poort die de adoptie naar 0 brengt naarmate
    α → α_max:

        gate(α) = 1 − e^{−(α_max − α)/s}   voor α < α_max,   anders 0

    Dit is gate(α) = P(κ_{c,n} > α) met klant-specifieke breakeven
    κ_{c,n} = κ_max − E, E ~ Exp(gemiddelde s). De bovengrens α_max = κ_c is
    dus de maximale breakeven: bij α = α_max geldt gate = 0, zodat Z_i exact 0
    wordt zodra de bovengrens α_U bereikt wordt. Kleine s → bijna harde cutoff;
    grotere s → geleidelijker aflopen door spreiding in de klant-κ_c.
    """
    _eps = 1e-9
    p0 = float(min(max(p0, _eps), 1.0 - _eps))
    if alpha0 == 0:
        alpha0 = _eps

    def _g(x: float) -> float:
        x = float(min(max(x, _eps), 1.0 - _eps))
        return -np.log(1.0 - x)

    _logit_p0 = np.log(p0 / (1.0 - p0))
    _y = (
        _logit_p0
        - gamma_alpha * (alpha - alpha0) / alpha0
        + gamma_X * (_g(X) - _g(X0))
    )
    p = float(1.0 / (1.0 + np.exp(-_y)))

    # WTP-plafond: adoptie zakt glad naar 0 naarmate α de bovengrens α_U nadert.
    if alpha_max is not None:
        s = float(s_alpha)
        if alpha >= alpha_max:
            gate = 0.0
        elif s <= 0:
            gate = 1.0
        else:
            gate = 1.0 - np.exp(-(alpha_max - alpha) / s)
        p *= gate
    return float(p)


def verwacht_subscripties_per_component(
    excel_file     = None,
    codes          = None,
    rate_overrides = None,
) -> pd.Series:
    """Analytisch verwacht aantal subscripties E[Z_i] per component.

    Gebruikt de gesloten vorm E[Z_i] = Σ_n q_{in} met
    q_{in} = 1 − (1 − p_r)^{h_{in}}, zonder Monte Carlo. Handig om de
    verwachte Z snel (en deterministisch) door te zetten naar de andere
    tabs bij wijziging van α, X of de adoptie-parameters.

    Returns een Series geïndexeerd op Code met het verwachte aantal subs.
    """
    adoptie = laad_adoptie_data(excel_file, rate_overrides=rate_overrides)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    if adoptie.empty:
        return pd.Series(dtype=float)
    _q = 1.0 - (1.0 - adoptie['Adoption_rate']) ** adoptie['Orders_component_klant']
    _ez = _q.groupby(adoptie['Code']).sum()
    _ez.index.name = 'Code'
    return _ez.sort_values(ascending=False)


def gevoeligheid_verwachte_z(
    waarden,
    parameter:     str,
    p_dichtbij0:   float,
    p_ver0:        float,
    alpha:         float,
    X:             float,
    alpha0:        float,
    X0:            float,
    gamma_alpha:   float = 1.0,
    gamma_X:       float = 0.5,
    excel_file           = None,
    codes                = None,
    kappa_c:       float = None,
    s_alpha:       float = 0.02,
):
    """Totaal verwacht aantal subscripties Σ_i E[Z_i] als functie van α of X.

    Voor elke waarde in ``waarden`` wordt het prijspercentage α
    (``parameter='alpha'``) of het service level X (``parameter='X'``)
    gevarieerd, terwijl de andere op zijn vaste waarde blijft. De Adoptie-data
    wordt één keer geladen; per gridpunt worden de regionale adoptieparameters
    p_r(α,X) opnieuw berekend en het totaal Σ_i Σ_n q_{in} bepaald.

    Returns een lijst totalen, parallel aan ``waarden``.
    """
    if parameter not in ('alpha', 'X'):
        raise ValueError("parameter moet 'alpha' of 'X' zijn")
    # Originele (ongewijzigde) rates laden om de twee niveaus te detecteren.
    adoptie = laad_adoptie_data(excel_file, rate_overrides=None)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    if adoptie.empty:
        return [0.0 for _ in waarden]
    rates = adoptie['Adoption_rate'].to_numpy(dtype=float)
    h     = adoptie['Orders_component_klant'].to_numpy(dtype=float)
    r4    = np.round(rates, 4)
    _pos  = r4[r4 > 0]
    totalen = []
    if _pos.size == 0:
        return [0.0 for _ in waarden]
    _niveaus = sorted(pd.Series(_pos).value_counts().index.tolist()[:2])
    _laag = _niveaus[0]
    _hoog = _niveaus[-1]
    high_mask = r4 == round(_hoog, 4)
    low_mask  = r4 == round(_laag, 4)
    for _v in waarden:
        _a = float(_v) if parameter == 'alpha' else float(alpha)
        _x = float(_v) if parameter == 'X' else float(X)
        _p_d = regionale_adoptie_parameter(
            p_dichtbij0, _a, _x, alpha0, X0, gamma_alpha, gamma_X,
            alpha_max=kappa_c, s_alpha=s_alpha)
        _p_v = regionale_adoptie_parameter(
            p_ver0, _a, _x, alpha0, X0, gamma_alpha, gamma_X,
            alpha_max=kappa_c, s_alpha=s_alpha)
        _rate = rates.copy()
        _rate[high_mask] = _p_d
        if _hoog != _laag:
            _rate[low_mask] = _p_v
        _q = 1.0 - (1.0 - _rate) ** h
        totalen.append(float(_q.sum()))
    return totalen


def pareto_alpha_X(
    overzicht_df,
    alpha_waarden,
    X_waarden,
    p_dichtbij0:   float,
    p_ver0:        float,
    alpha0:        float,
    X0:            float,
    kappa_bpa:     float,
    kappa_c:       float,
    gamma_alpha:   float = 1.0,
    gamma_X:       float = 0.5,
    excel_file           = None,
    codes                = None,
    gebruik_simulatie: bool = False,
    n_runs:        int = 500,
    seed:          int = 42,
    s_alpha:       float = 0.02,
) -> pd.DataFrame:
    """Pareto-analyse over (α, X): BPA-marge vs. totaal klantsurplus.

    Voor elk paar (α, X) uit het cartesisch product van ``alpha_waarden`` en
    ``X_waarden`` wordt de volledige keten doorgerekend:

      1. p_r(α, X) per regio via de logit-specificatie (willingness-to-pay);
      2. Z_i per component: analytisch E[Z_i] = Σ_n q_{in} (default) of —
         als ``gebruik_simulatie=True`` — het Monte-Carlo gemiddelde over
         ``n_runs`` Bernoulli-trekkingen X_{in} ~ Bernoulli(q_{in}) (zelfde
         keuze-model als de simulatietab);
      3. een overzicht met n_klanten = Z_i (λ proportioneel meegeschaald);
      4. het kostenmodel → BPA-marge en het totale klantsurplus
         Σ_klant (zelf-voorraadkosten − abonnementskosten).

    Parameters
    ----------
    gebruik_simulatie : gebruik Monte-Carlo gesimuleerde Z i.p.v. de
                        analytische verwachtingswaarde.
    n_runs, seed      : aantal trekkingen en seed voor de Monte-Carlo modus.
                        Dezelfde seed wordt per (α,X) hergebruikt (common
                        random numbers) zodat de marge-curve glad blijft.

    Returns een DataFrame met kolommen:
        alpha, X, margin, surplus, total_Z, feasible.
    """
    if overzicht_df is None or overzicht_df.empty:
        return pd.DataFrame()

    adoptie = laad_adoptie_data(excel_file, rate_overrides=None)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    if adoptie.empty:
        return pd.DataFrame()

    rates    = adoptie['Adoption_rate'].to_numpy(dtype=float)
    h        = adoptie['Orders_component_klant'].to_numpy(dtype=float)
    code_arr = adoptie['Code'].to_numpy()
    r4       = np.round(rates, 4)
    _pos     = r4[r4 > 0]
    if _pos.size == 0:
        return pd.DataFrame()
    _niveaus = sorted(pd.Series(_pos).value_counts().index.tolist()[:2])
    _laag = _niveaus[0]
    _hoog = _niveaus[-1]
    high_mask = r4 == round(_hoog, 4)
    low_mask  = r4 == round(_laag, 4)

    # Per-component λ per klant uit het basisoverzicht (λ/n is invariant in n).
    base      = overzicht_df
    base_n    = base['n_klanten'].astype(float)
    base_lam  = base['lambda_jr'].astype(float)
    lam_per_cust = (base_lam / base_n.replace(0, np.nan)).fillna(0.0)

    rijen = []
    for _a in alpha_waarden:
        for _x in X_waarden:
            _p_d = regionale_adoptie_parameter(
                p_dichtbij0, _a, _x, alpha0, X0, gamma_alpha, gamma_X,
                alpha_max=kappa_c, s_alpha=s_alpha)
            _p_v = regionale_adoptie_parameter(
                p_ver0, _a, _x, alpha0, X0, gamma_alpha, gamma_X,
                alpha_max=kappa_c, s_alpha=s_alpha)
            _rate = rates.copy()
            _rate[high_mask] = _p_d
            if _hoog != _laag:
                _rate[low_mask] = _p_v
            if gebruik_simulatie:
                # Bernoulli-keuze per (component, klant): X_in ~ Bernoulli(q_in)
                # met q_in = 1 - (1 - p_r)^{h_in} = P(≥ 1 conversie over h_in
                # orders). Z_i = Σ_n X_in (Poisson-binomiaal over de klanten);
                # gemiddeld over de runs. Zelfde seed per (α,X) → gladde curve.
                _q_in  = 1.0 - (1.0 - _rate) ** h
                _rng   = np.random.default_rng(seed)
                _draws = _rng.random((int(n_runs), _q_in.shape[0])) < _q_in[None, :]
                _q = _draws.mean(axis=0)
            else:
                _q = 1.0 - (1.0 - _rate) ** h
            _ez = pd.Series(_q).groupby(code_arr).sum()

            # Overzicht met n_klanten = Z_i (λ meegeschaald); componenten
            # zonder adoptie-data behouden hun basiswaarden.
            _n_new   = _ez.reindex(base.index)
            _present = _n_new.notna()
            # Aanwezige componenten mogen naar 0 abonnees zakken (Z_i = 0 bij
            # α ≥ κ_c); niet-aanwezige componenten behouden hun basiswaarde.
            _n_int   = _n_new.where(_present, base_n).round().clip(lower=0).astype(int)
            mod = base.copy()
            mod['n_klanten'] = _n_int.values
            mod['lambda_jr'] = np.where(
                _present.values,
                _n_int.values.astype(float) * lam_per_cust.values,
                base_lam.values,
            )
            if int(mod['n_klanten'].sum()) <= 0:
                # Geen enkele abonnee → geen omzet en geen kosten: marge = 0.
                _margin  = 0.0
                _surplus = 0.0
                _feas    = False
            else:
                try:
                    _model, _res = bouw_model_kosten(mod, _a, kappa_bpa, kappa_c, _x)
                    _margin  = float(_res['bpa_margin'])
                    _surplus = float(sum(
                        v['savings'] for v in _res['customer_benefits'].values()))
                    _feas    = bool(_res['feasible'])
                except Exception:
                    _margin = _surplus = float('nan')
                    _feas   = False
            rijen.append({
                'alpha':    float(_a),
                'X':        float(_x),
                'margin':   _margin,
                'surplus':  _surplus,
                'total_Z':  float(_ez.sum()),
                'feasible': _feas,
            })
    return pd.DataFrame(rijen)


def metrieken_voor_wtp_grid(
    overzicht_df,
    param_dicts,
    kappa_bpa: float,
    kappa_c:   float,
    excel_file       = None,
    codes            = None,
):
    """Meerdere keten-uitkomsten per WTP-parametercombinatie.

    Generaliseert de keten van ``pareto_alpha_X``: voor elke parameter-dict in
    ``param_dicts`` worden de twee regionale adoptieparameters p_r berekend
    (Benelux/overig), daaruit het verwachte aantal subscripties E[Z_i] per
    component, en vervolgens via het kostenmodel de BPA-marge en het totale
    klantsurplus.

    Elke dict bevat de sleutels: ``p_dichtbij0``, ``p_ver0``, ``alpha``, ``X``,
    ``alpha0``, ``X0``, ``gamma_alpha``, ``gamma_X``, ``s_alpha`` en optioneel
    ``alpha_max`` (= κ_c-plafond op α; None = geen plafond).

    De Adoptie-data en het basisoverzicht worden één keer geladen; per dict
    wordt alleen p_r → E[Z] → kostenmodel doorgerekend.

    Returns een lijst dicts (parallel aan ``param_dicts``), elk met sleutels:
        ``bpa_margin``  – totale BPA-marge (€);
        ``surplus``     – totaal klantsurplus (€);
        ``total_Z``     – verwacht totaal aantal subscripties E[Z];
        ``p_dichtbij``  – adoptieparameter p_r (Benelux);
        ``p_ver``       – adoptieparameter p_r (overig);
        ``feasible``    – of het kostenmodel haalbaar was.
    NaN bij een niet-doorgerekende combinatie.
    """
    def _leeg():
        return {
            'bpa_margin': float('nan'), 'surplus': float('nan'),
            'total_Z': float('nan'), 'p_dichtbij': float('nan'),
            'p_ver': float('nan'), 'feasible': False,
        }

    if overzicht_df is None or overzicht_df.empty or not param_dicts:
        return [_leeg() for _ in (param_dicts or [])]

    adoptie = laad_adoptie_data(excel_file, rate_overrides=None)
    codes_set = {str(c).strip() for c in codes} if codes is not None else classificatie_codes()
    if codes_set:
        adoptie = adoptie[adoptie['Code'].isin(codes_set)]
    if adoptie.empty:
        return [_leeg() for _ in param_dicts]

    rates    = adoptie['Adoption_rate'].to_numpy(dtype=float)
    h        = adoptie['Orders_component_klant'].to_numpy(dtype=float)
    code_arr = adoptie['Code'].to_numpy()
    r4       = np.round(rates, 4)
    _pos     = r4[r4 > 0]
    if _pos.size == 0:
        return [_leeg() for _ in param_dicts]
    _niveaus = sorted(pd.Series(_pos).value_counts().index.tolist()[:2])
    _laag = _niveaus[0]
    _hoog = _niveaus[-1]
    high_mask = r4 == round(_hoog, 4)
    low_mask  = r4 == round(_laag, 4)

    base      = overzicht_df
    base_n    = base['n_klanten'].astype(float)
    base_lam  = base['lambda_jr'].astype(float)
    lam_per_cust = (base_lam / base_n.replace(0, np.nan)).fillna(0.0)

    resultaten = []
    for p in param_dicts:
        _a    = float(p['alpha'])
        _x    = float(p['X'])
        _amax = p.get('alpha_max', None)
        _s    = float(p.get('s_alpha', 0.02))
        _p_d = regionale_adoptie_parameter(
            p['p_dichtbij0'], _a, _x, p['alpha0'], p['X0'],
            p['gamma_alpha'], p['gamma_X'], alpha_max=_amax, s_alpha=_s)
        _p_v = regionale_adoptie_parameter(
            p['p_ver0'], _a, _x, p['alpha0'], p['X0'],
            p['gamma_alpha'], p['gamma_X'], alpha_max=_amax, s_alpha=_s)
        _rate = rates.copy()
        _rate[high_mask] = _p_d
        if _hoog != _laag:
            _rate[low_mask] = _p_v
        _q  = 1.0 - (1.0 - _rate) ** h
        _ez = pd.Series(_q).groupby(code_arr).sum()
        _total_z = float(_ez.sum())
        _n_new   = _ez.reindex(base.index)
        _present = _n_new.notna()
        _n_int   = _n_new.where(_present, base_n).round().clip(lower=0).astype(int)
        mod = base.copy()
        mod['n_klanten'] = _n_int.values
        mod['lambda_jr'] = np.where(
            _present.values,
            _n_int.values.astype(float) * lam_per_cust.values,
            base_lam.values,
        )
        _rec = {
            'total_Z':    _total_z,
            'p_dichtbij': float(_p_d),
            'p_ver':      float(_p_v),
        }
        if int(mod['n_klanten'].sum()) <= 0:
            _rec.update({'bpa_margin': 0.0, 'surplus': 0.0, 'feasible': False})
        else:
            try:
                _model, _res = bouw_model_kosten(mod, _a, kappa_bpa, kappa_c, _x)
                _rec.update({
                    'bpa_margin': float(_res['bpa_margin']),
                    'surplus':    float(sum(
                        v['savings'] for v in _res['customer_benefits'].values())),
                    'feasible':   bool(_res['feasible']),
                })
            except Exception:
                _rec.update({'bpa_margin': float('nan'),
                             'surplus': float('nan'), 'feasible': False})
        resultaten.append(_rec)
    return resultaten


def winst_voor_wtp_grid(
    overzicht_df,
    param_dicts,
    kappa_bpa: float,
    kappa_c:   float,
    excel_file       = None,
    codes            = None,
):
    """Totale BPA-marge (€) voor een lijst WTP-parametercombinaties.

    Dunne wrapper rond :func:`metrieken_voor_wtp_grid` die alleen de
    BPA-marge per parameter-dict teruggeeft (lijst van floats, parallel aan
    ``param_dicts``; NaN bij een niet-doorgerekende combinatie).
    """
    _recs = metrieken_voor_wtp_grid(
        overzicht_df, param_dicts, kappa_bpa, kappa_c,
        excel_file=excel_file, codes=codes,
    )
    return [r.get('bpa_margin', float('nan')) for r in _recs]


def optimale_alpha_bij_X(
    overzicht_df,
    X_fix:         float,
    alpha_grid,
    p_dichtbij0:   float,
    p_ver0:        float,
    alpha0:        float,
    X0:            float,
    kappa_bpa:     float,
    kappa_c:       float,
    gamma_alpha:   float = 1.0,
    gamma_X:       float = 0.5,
    excel_file           = None,
    codes                = None,
    alleen_haalbaar:     bool = False,
    gebruik_simulatie: bool = False,
    n_runs:        int = 500,
    seed:          int = 42,
    s_alpha:       float = 0.02,
):
    """Zoek het prijspercentage α dat de BPA-marge maximaliseert bij vast X.

    Houdt het service level X vast op ``X_fix`` en evalueert de volledige
    keten (α,X) → p_r → E[Z_i] → kostenmodel voor elke α in ``alpha_grid``.
    Omdat een hogere α de omzet per klant verhoogt maar de adoptie verlaagt,
    bestaat doorgaans een inwendig marge-optimum.

    Parameters
    ----------
    alleen_haalbaar : als True wordt het optimum alleen gezocht onder de
                      haalbare combinaties (marge ≥ 0 én alle klanten
                      profiteren); bij geen enkele haalbare α valt de functie
                      terug op de marge-maximalisatie over alle α.

    Returns
    -------
    (curve_df, best_row) waarbij:
        curve_df : DataFrame met kolommen alpha, X, margin, surplus,
                   total_Z, feasible (de volledige α-sweep bij X_fix);
        best_row : de rij (pd.Series) met de gekozen optimale α, of None
                   als er geen geldige marge berekend kon worden.
    """
    curve = pareto_alpha_X(
        overzicht_df, list(alpha_grid), [float(X_fix)],
        p_dichtbij0, p_ver0, alpha0, X0, kappa_bpa, kappa_c,
        gamma_alpha, gamma_X, excel_file=excel_file, codes=codes,
        gebruik_simulatie=gebruik_simulatie, n_runs=n_runs, seed=seed,
        s_alpha=s_alpha,
    )
    if curve.empty:
        return curve, None
    _geldig = curve.dropna(subset=['margin'])
    if _geldig.empty:
        return curve, None
    _kandidaten = _geldig
    if alleen_haalbaar:
        _haalbaar = _geldig[_geldig['feasible']]
        if not _haalbaar.empty:
            _kandidaten = _haalbaar
    best = _kandidaten.loc[_kandidaten['margin'].idxmax()]
    return curve, best


def _hermap_adoption_rates(rates: pd.Series, rate_overrides) -> pd.Series:
    """Hermap de twee adoption-rate niveaus (Benelux hoog / overig laag).

    De data bevat per (component, klant) een vooraf bepaalde adoption rate met
    twee niveaus: een hoog niveau voor Benelux-klanten en een laag niveau voor
    de overige klanten. Met deze functie kunnen die twee niveaus in de tool
    worden overschreven zonder de data opnieuw te genereren.

    rate_overrides : dict met optionele sleutels 'benelux' en/of 'overig'
                     (waarden tussen 0 en 1). De twee oorspronkelijke niveaus
                     worden gedetecteerd als de twee meest voorkomende positieve
                     waarden; rijen op het hoge niveau krijgen de Benelux-rate,
                     rijen op het lage niveau de overig-rate. Nul-waarden (geen
                     data) blijven ongemoeid.
    """
    if not rate_overrides:
        return rates
    p_ben = rate_overrides.get('benelux')
    p_ov  = rate_overrides.get('overig')
    if p_ben is None and p_ov is None:
        return rates
    _pos = rates[rates > 0]
    if _pos.empty:
        return rates
    # Twee meest voorkomende positieve niveaus = Benelux (hoog) en overig (laag).
    _niveaus = sorted(_pos.round(4).value_counts().index.tolist()[:2])
    if not _niveaus:
        return rates
    _laag = _niveaus[0]
    _hoog = _niveaus[-1]
    _r = rates.round(4)
    out = rates.copy()
    if p_ben is not None:
        out = out.mask(_r == round(_hoog, 4), float(p_ben))
    if p_ov is not None and _hoog != _laag:
        out = out.mask(_r == round(_laag, 4), float(p_ov))
    return out.clip(0.0, 1.0)


def laad_adoptie_data(excel_file=None, rate_overrides=None) -> pd.DataFrame:
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
    # Optioneel: de twee adoption-rate niveaus (Benelux/overig) overschrijven.
    df['Adoption_rate'] = _hermap_adoption_rates(df['Adoption_rate'], rate_overrides)
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
    rate_overrides     = None,
) -> pd.DataFrame:
    """
    Monte Carlo simulatie van het aantal subscripties per component.

    Per run en per (component, klant) wordt de abonneerkans q_klant bepaald en
    een Bernoulli-trekking gedaan voor de binaire abonneer-beslissing:

        q_klant   = 1 - (1 - adoption_rate) ** aantal_orders
        X_klant   ~ Bernoulli(q_klant)   (1 = abonneert, 0 = niet)

    De kans dat een klant een subscriptie neemt stijgt zo met het aantal orders
    voor het component:
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
        rate_overrides=rate_overrides,
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
    rate_overrides     = None,
) -> dict:
    """
    Voer de subscriptie-simulatie uit en geef de ruwe trekkingen terug.

    Per klant wordt eerst de abonneerkans q_klant = 1 - (1 - adoption_rate)^{orders}
    bepaald (= kans op ≥ 1 conversie over de orders). Daarna volgt een
    Bernoulli-trekking voor de binaire abonneer-beslissing:
        q_klant = 1 - (1 - adoption_rate) ** orders
        X_klant ~ Bernoulli(q_klant)   (1 = abonneert, 0 = niet)
    Het aantal subscripties per component is de som van deze binaire
    beslissingen over de klanten (Poisson-binomiaal, maximaal het aantal
    klanten).

    Bij codes=None wordt standaard beperkt tot de classificatie-componenten.

    Returns een dict {code → np.ndarray van lengte n_runs} met per run het
    gesimuleerde aantal subscripties voor dat component. Handig voor het
    plotten van de verdeling (histogram) per component.
    """
    adoptie = laad_adoptie_data(excel_file, rate_overrides=rate_overrides)
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
        # Bernoulli-keuze per klant: X_klant ~ Bernoulli(q_klant) met
        # q_klant = 1 - (1 - adoption_rate)^{orders} = P(≥ 1 conversie over de
        # orders van de klant). Zo stijgt de abonneerkans met het aantal orders.
        # Z_component = Σ_klant X_klant (Poisson-binomiaal over de klanten).
        # (n_runs, n_klanten) Bernoulli-trekkingen → som per run.
        q_arr = 1.0 - (1.0 - p_arr) ** n_arr
        draws = rng.random((n_runs, q_arr.shape[0])) < q_arr[None, :]
        runs[code] = draws.sum(axis=1)
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
    cfg.setdefault('ip_overrides', {})
    cfg.setdefault('lt_overrides', {})

    print("\nOverride per component – N, inkoopprijs en levertijd (laat code leeg om te stoppen):")
    print("  Laat een veld leeg om het ongewijzigd te laten. Typ 'x' om een override te verwijderen.")
    while True:
        code = input("  Artikelcode (of Enter om te stoppen): ").strip()
        if not code:
            break

        # N
        huidig_n = cfg['n_klanten_overrides'].get(code, '(uit data)')
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
        n     = int(input("  Aantal subscripties [1]: ").strip() or 1)
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
║  2. Pas subscripties aan (per code)               ║
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
