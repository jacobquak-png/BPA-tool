"""
BPA Jaarlijks Beheer Tool – Streamlit UI
=========================================
Start met:
    streamlit run src/bpa_beheer_ui.py

Vereist:
    pip install streamlit
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
from datetime import date
import pandas as pd
import numpy as np
import json

# Hergebruik alle logica uit bpa_beheer.py
from bpa_beheer import (
    laad_config,
    sla_config_op,
    bereken_overzicht,
    bouw_model_kosten,
    laad_excel_onderdelen,
    laad_classificatie_selectie,
    regionale_adoptie_parameter,
    adoptie_kans,
    aantal_klanten_per_component,
    binomiale_verdeling,
    verwacht_subscripties_per_component,
    gevoeligheid_verwachte_z,
    pareto_alpha_X,
    optimale_alpha_bij_X,
    beta_r_winstband,
    winst_voor_wtp_grid,
    metrieken_voor_wtp_grid,
    SERVICE_LEVELS,
    CONFIG_PATH,
    HISTORY_PATH,
    SCRIPT_DIR,
    SELECTIE_PATH,
    EXCEL_PATH,
    SUBSCRIPTIES_PATH,
)
from classificatie import (
    ClassificatieParams,
    voer_classificatie_uit,
    schrijf_selectie_json,
    controleer_kolommen,
    laad_ruwe_dataset,
    bereken_scores,
    pas_basis_filters_toe,
    pas_topn_selectie_toe,
    bouw_selectie_payload,
    weight_sensitivity,
)
from model import BPAOptimizationModel

# ══════════════════════════════════════════════════════════════════════════════
#  CACHE-WRAPPERS  (sterk versnellen Streamlit-reruns)
# ══════════════════════════════════════════════════════════════════════════════
#
# Streamlit voert dit script opnieuw uit bij élke widget-interactie. Zonder
# caching wordt de (grote) Excel telkens opnieuw geparsed en doorloopt
# `bereken_overzicht` weer alle componenten. De wrappers hieronder zorgen dat
# we alleen herrekenen als (a) een bron-bestand op disk gewijzigd is óf
# (b) de gebruiker de config heeft aangepast. Cache wordt automatisch
# ongeldig zodra een van die inputs verandert.

def _file_mtime(path: str) -> float:
    """Return mtime in seconds; 0.0 als bestand ontbreekt."""
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


@st.cache_data(show_spinner=False, max_entries=4)
def _cached_laad_classificatie_selectie(_mtime: float) -> dict:
    """Cached versie van laad_classificatie_selectie — keyed op bestand-mtime."""
    return laad_classificatie_selectie()


@st.cache_data(show_spinner=False, max_entries=4)
def _cached_laad_ruwe_dataset(_excel_mtime: float, sheet_name, upload=None) -> pd.DataFrame:
    """Cache de (trage) Excel-parse voor de classificatie.

    Keyed op bestand-mtime + sheet voor de repo-Excel, of op de geüploade
    file-inhoud (Streamlit hasht een UploadedFile op inhoud, dus de parameter
    krijgt GEEN underscore-prefix — anders zou een tweede upload met dezelfde
    sheet-naam onterecht de vorige cache-hit teruggeven). Hierdoor wordt de
    Excel maar één keer geparsed per uniek bestand; daarna gaan parameter-tweaks
    razendsnel omdat alleen de gevectoriseerde scoring opnieuw draait.
    """
    if upload is not None:
        # Reset de leespositie: een eerder gelezen/gehashte buffer kan aan het
        # einde staan, waardoor pd.read_excel niets zou inlezen.
        try:
            upload.seek(0)
        except (AttributeError, ValueError):
            pass
        bron = upload
    else:
        bron = EXCEL_PATH
    return laad_ruwe_dataset(bron, sheet_name=sheet_name)


@st.cache_data(show_spinner=False, max_entries=4)
def _cached_bereken_overzicht(cfg_json: str, _excel_mtime: float, _selectie_mtime: float) -> pd.DataFrame:
    """Cached versie van bereken_overzicht — keyed op JSON-config + bestand-mtimes."""
    return bereken_overzicht(json.loads(cfg_json))


def get_classificatie_info() -> dict:
    """Lees bpa_selectie.json (cached). Auto-invalideert bij file-update."""
    return _cached_laad_classificatie_selectie(_file_mtime(SELECTIE_PATH))


def get_overzicht_df(cfg: dict) -> pd.DataFrame:
    """Bereken het overzicht (cached). Auto-invalideert bij config- of bestand-wijziging."""
    cfg_json = json.dumps(cfg, sort_keys=True, default=str)
    return _cached_bereken_overzicht(
        cfg_json,
        _file_mtime(EXCEL_PATH),
        _file_mtime(SELECTIE_PATH),
    )


def invalidate_caches() -> None:
    """Forceer een verse Excel/JSON-read bij volgende aanroep."""
    _cached_bereken_overzicht.clear()
    _cached_laad_classificatie_selectie.clear()
    _cached_laad_ruwe_dataset.clear()


@st.cache_data(show_spinner="Gewichten-sweep berekenen…", max_entries=8)
def _cached_weight_sweep(_df_scored: pd.DataFrame, params_json: str, step: float,
                         versie: int = 2):
    """Cached gewicht-sweep. `_df_scored` (leidende underscore) wordt NIET
    gehasht; de cache-sleutel is `params_json` + `step` + `versie`. De
    versie-token wordt opgehoogd wanneer de output-vorm wijzigt (bv. nieuwe
    rangorde-kolommen), zodat oude cache-resultaten automatisch verlopen.
    """
    p = json.loads(params_json)
    params = ClassificatieParams(
        threshold=p["threshold"],
        selectie_modus=p["selectie_modus"],
        top_n=p["top_n"],
        weight_prijs=p["weight_prijs"],
        weight_locaties=p["weight_locaties"],
        weight_orders=p["weight_orders"],
        orders_power=p["orders_power"],
        min_prijs=p.get("min_prijs", 0.0),
        min_orders=p.get("min_orders", 0.0),
        min_klantlocaties=p["min_klantlocaties"],
        article_type_filter=tuple(p["article_type_filter"]),
        score_methode=p.get("score_methode", "arithmetisch"),
        epsilon=p.get("epsilon", 1.0),
    )
    return weight_sensitivity(_df_scored, params, step=step, return_combos=True)


def representatieve_z(default: int = 1) -> int:
    """Representatief aantal subscripties (Z) uit het huidige overzicht.

    Vervangt de oude globale standaardwaarde: neemt de mediaan van het
    werkelijke aantal klantlocaties (n_klanten) over alle componenten. Wordt
    gebruikt als referentie-/startwaarde in de gevoeligheidsgrafieken.
    """
    _df = st.session_state.get("overzicht_df")
    if _df is not None and not _df.empty and "n_klanten" in _df.columns:
        _med = pd.to_numeric(_df["n_klanten"], errors="coerce").median()
        if pd.notna(_med) and _med >= 1:
            return int(round(_med))
    return default


# ══════════════════════════════════════════════════════════════════════════════
#  PAGINA-INSTELLINGEN
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="BPA Beheer Tool",
    page_icon="⚙️",
    layout="wide",
)

_logo_path = os.path.join(os.path.dirname(__file__), "BPA.png")
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if os.path.exists(_logo_path):
        st.image(_logo_path, width=120)
with col_title:
    st.title("BPA Jaarlijks Beheer Tool")
    st.caption(f"Configuratiebestand: `{CONFIG_PATH}`")

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG IN SESSION STATE LADEN
# ══════════════════════════════════════════════════════════════════════════════

if "cfg" not in st.session_state:
    st.session_state.cfg = laad_config()

cfg = st.session_state.cfg

# ── Excel altijd uit de repository ────────────────────────────────────────
_excel_file = None  # gebruik altijd EXCEL_PATH uit de repo

# Overzicht altijd vers berekenen bij opstarten van de sessie
if "overzicht_df" not in st.session_state:
    with st.spinner("Excel laden en basisvoorraden berekenen…"):
        _df = get_overzicht_df(cfg)
    if not _df.empty:
        st.session_state.overzicht_df = _df

# ══════════════════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════════════════

tab_overzicht, tab_subscripties, tab_toevoegen, tab_verwijderen, tab_config, tab_historie, tab_kosten, tab_drempel, tab_classificatie, tab_budget, tab_subsim, tab_sensitivity = st.tabs([
    "📊 Overzicht",
    "✏️ Subscripties aanpassen",
    "➕ Component toevoegen",
    "🗑️ Component verwijderen",
    "⚙️ Configuratie",
    "📈 Historiek",
    "💰 Kostenanalyse",
    "🔢 Subscriptiedrempel",
    "🏷️ Classificatie",
    "💼 Budget-scenario",
    "📈 Verwachte subscripties",
    "📐 Sensitivity (WTP)",
])

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 – OVERZICHT
# ─────────────────────────────────────────────────────────────────────────────

with tab_overzicht:
    st.subheader("Basisvoorraden per component")

    # Excel-bestandsdatum ophalen
    try:
        from bpa_beheer import EXCEL_PATH as _EXCEL_PATH
        _excel_mtime = date.fromtimestamp(os.path.getmtime(_EXCEL_PATH)).isoformat()
    except Exception:
        _excel_mtime = "onbekend"

    st.write(
        f"Configuratie bijgewerkt: **{cfg['aangepast']}** · "
        f"Excel gewijzigd: **{_excel_mtime}**"
    )

    # ── Classificatie-koppeling status ────────────────────────────────────
    _cls_info = get_classificatie_info()
    if _cls_info:
        _lt_ov = _cls_info.get('lt_overzicht', {})
        _n_cls = len(_cls_info.get('items', {}))
        st.success(
            f"🔗 Classificatie-koppeling actief — **{_n_cls}** componenten geselecteerd "
            f"(gegenereerd {_cls_info.get('gegenereerd', '?')}). "
            f"LT-bron: ✅ geupdate **{_lt_ov.get('geupdate', 0)}**  ·  "
            f"⚠️ ERP-default **{_lt_ov.get('default', 0)}**  ·  "
            f"❌ ontbreekt **{_lt_ov.get('ontbreekt', 0)}**"
        )
    else:
        st.info(
            f"ℹ️  Geen classificatie-selectie gevonden ({SELECTIE_PATH}). "
            f"Draai `classificatie_scoring.py` om de koppeling te activeren."
        )

    if st.button("🔄 Herbereken (laadt Excel opnieuw)"):
        invalidate_caches()
        with st.spinner("Berekenen…"):
            df = get_overzicht_df(cfg)
        if df.empty:
            st.warning("Geen onderdelen gevonden.")
        else:
            st.session_state.overzicht_df = df
            st.rerun()

    if "overzicht_df" in st.session_state:
        df = st.session_state.overzicht_df
        sl_cols = [c for c in df.columns if c.startswith("s@")]

        # Samenvattingsregel
        totals = {c: int(df[c].sum()) for c in sl_cols}
        st.write("**Totale basisvoorraad:**  " +
                 "  |  ".join(f"`{c}` = **{totals[c]}**" for c in sl_cols))

        # Aandeel S* > 1 — extra voorraadkosten bovenop S*=1
        if sl_cols and 'IP' in df.columns:
            _parts = []
            for _sc in sl_cols:
                _ip_vals     = df['IP'].fillna(0)
                _extra_units = (df[_sc] - 1).clip(lower=0)          # max(S*-1, 0) per component
                _base_cost   = _ip_vals.sum()                        # Σ 1 × IP (S*=1 scenario)
                _extra_cost  = (_extra_units * _ip_vals).sum()       # Σ (S*-1) × IP
                _total_cost  = (df[_sc] * _ip_vals).sum()
                _pct_extra   = _extra_cost / _total_cost * 100 if _total_cost > 0 else 0.0
                _n_gt1       = int((df[_sc] > 1).sum())
                _parts.append(
                    f"`{_sc}` → **{_n_gt1}** comp. met S\u002a > 1, "
                    f"extra kost boven S\u002a=1: **€ {_extra_cost:,.0f}** (**{_pct_extra:.1f}%** van totale inv.)"
                )
            if _parts:
                st.caption("Extra inv. bovenop S\u002a=1:  \n" + "  \n".join(_parts))

        # Laad vorige snapshot voor Δ-kolommen
        _prev_comp = {}
        _prev_datum = None
        # 1) Voorkeur: vorige overzicht_df uit session_state (vastgelegd bij opslaan)
        _prev_df = st.session_state.get("overzicht_df_prev")
        if _prev_df is not None and not _prev_df.empty:
            _prev_datum = "vorige opgeslagen staat"
            for _code in _prev_df.index:
                _prev_comp[str(_code)] = {
                    _sc: int(_prev_df.at[_code, _sc])
                    for _sc in _prev_df.columns if _sc.startswith("s@")
                }
        # 2) Fallback: oudere snapshot uit history-bestand (legacy)
        elif os.path.exists(HISTORY_PATH):
            try:
                with open(HISTORY_PATH, encoding='utf-8') as _fh:
                    _hist_ov = json.load(_fh)
                for _snap_ov in reversed(_hist_ov):
                    if 'componenten' in _snap_ov:
                        _prev_comp  = _snap_ov['componenten']
                        _prev_datum = _snap_ov['datum']
                        break
            except Exception:
                pass

        # Bouw weergave-df met Δ-kolommen
        _df_disp = df.reset_index().copy()
        _delta_cols = []
        # Vectoriseer: bouw één lookup-DataFrame van vorige S*-waarden per Code,
        # zodat we per SL-kolom alleen een Series-aftrekking nodig hebben
        # (i.p.v. .apply(axis=1) — orde van grootte sneller bij veel rijen).
        if _prev_comp and sl_cols:
            _prev_df_lookup = (
                pd.DataFrame.from_dict(_prev_comp, orient="index")
                  .reindex(columns=sl_cols)
                  .apply(pd.to_numeric, errors="coerce")
            )
            _codes_str = _df_disp["Code"].astype(str)
            for _sc in sl_cols:
                _dc = f"\u0394{_sc}"
                _delta_cols.append(_dc)
                _prev_series = _codes_str.map(_prev_df_lookup[_sc])
                _df_disp[_dc] = (
                    pd.to_numeric(_df_disp[_sc], errors="coerce") - _prev_series
                )
        else:
            # Geen vorige snapshot beschikbaar — vul Δ-kolommen met NaN
            for _sc in sl_cols:
                _dc = f"\u0394{_sc}"
                _delta_cols.append(_dc)
                _df_disp[_dc] = float("nan")

        def _fmt_delta(v):
            if pd.isna(v): return '\u2014'
            iv = int(v)
            return f"+{iv}" if iv > 0 else str(iv)

        def _style_delta(v):
            if pd.isna(v): return ''
            if v > 0: return 'background-color: #f8d7da'   # rood: omhoog
            if v < 0: return 'background-color: #cce5ff'   # blauw: omlaag
            return 'background-color: #d4edda'              # groen: gelijk

        if _prev_datum:
            st.caption(f"\u0394 ten opzichte van snapshot: **{_prev_datum}**")
        else:
            st.caption("\u0394-kolommen beschikbaar na eerste snapshot (tabblad \U0001f4c8 Historiek).")

        # ── Investering per component ──────────────────────────────────────
        _kp_ov = st.session_state.get('kosten_params', {})
        _sl_ov = _kp_ov.get('service_level', 0.990)
        _sl_ov_col = f"s@{_sl_ov:.1%}"
        # Fallback: gebruik de eerste beschikbare SL-kolom als de gewenste er niet in zit
        if _sl_ov_col not in df.columns and sl_cols:
            _sl_ov_col = sl_cols[0]
            _sl_ov = float(_sl_ov_col[2:-1]) / 100
        if _sl_ov_col in df.columns and 'IP' in df.columns:
            _df_disp['Inv. (€)'] = (_df_disp[_sl_ov_col] * df['IP'].values).round(2)
            _inv_totaal = _df_disp['Inv. (€)'].sum()
            _df_disp['Inv. %'] = (
                (_df_disp['Inv. (€)'] / _inv_totaal * 100).round(1)
                if _inv_totaal > 0 else 0.0
            )
            _inv_cols = ['Inv. (€)', 'Inv. %']
            st.caption(
                f"Inv. (€) = S\u002a × IP bij **{_sl_ov_col}** · "
                f"Totale voorraadwaarde: **€ {_inv_totaal:,.0f}** · "
                f"_(pas service level aan via tabblad 💰 Kostenanalyse)_"
            )
        else:
            _inv_cols = []

        # Tabel
        _fmt_inv  = {c: "{:.0f}" for c in _inv_cols if 'Inv. (€)' in c}
        _fmt_inv |= {c: "{:.1f}%" for c in _inv_cols if 'Inv. %' in c}

        def _style_inv_share(v):
            if pd.isna(v) or _inv_totaal == 0:
                return ''
            intensity = min(int(v / 100 * 255), 255)
            return f'background-color: rgba(25, 118, 210, {v/100:.2f}); color: {"white" if v > 50 else "black"}'

        # ── LT-status kolom (vanuit classificatie-koppeling) ──────────────
        _LT_ICOON = {
            'geupdate':  '✅ geupdate',
            'override':  '✏️ override',
            'default':   '⚠️ ERP-default',
            'ontbreekt': '❌ ontbreekt',
            'handmatig': '🛠 handmatig',
            'onbekend':  '❔ onbekend',
            'nul→30':   '🔵 0→30 dagen',
        }
        if 'LT_bron' in _df_disp.columns:
            _df_disp['LT-status'] = (
                _df_disp['LT_bron'].astype(str).map(_LT_ICOON).fillna(_LT_ICOON['onbekend'])
            )

            def _kleur_lt(v):
                s = str(v)
                if '🔵' in s:             return 'background-color: #bbdefb'  # blauw: LT was 0 → 30
                if '✅' in s or '✏️' in s: return 'background-color: #e8f5e9'
                if '⚠️' in s:              return 'background-color: #fff8e1'
                if '❌' in s:              return 'background-color: #ffebee'
                if '🛠' in s:              return 'background-color: #e3f2fd'
                return ''

            _n_bevest = _df_disp['LT_bron'].isin(['geupdate', 'override', 'handmatig', 'nul→30']).sum()
            _n_warn   = len(_df_disp) - _n_bevest
            if _n_warn > 0:
                st.warning(
                    f"⚠️ {_n_warn}/{len(_df_disp)} componenten hebben een niet-bevestigde "
                    f"levertijd (ERP-default of ontbrekend). Corrigeer via tab "
                    f"**✏️ Subscripties aanpassen** — een ingevulde LT-override telt als bevestigd."
                )

        styled = (
            _df_disp.style
                .format({
                    "lambda_jr": "{:.4f}",
                    "mu":        "{:.4f}",
                    **{c: "{:.0f}" for c in sl_cols},
                    **{dc: _fmt_delta for dc in _delta_cols},
                    **({'Inv. (€)': '€ {:,.0f}', 'Inv. %': '{:.1f}%'} if _inv_cols else {}),
                })
                .map(_style_delta, subset=_delta_cols)
        )
        if 'LT-status' in _df_disp.columns:
            styled = styled.map(_kleur_lt, subset=['LT-status'])
        if _inv_cols and 'Inv. %' in _df_disp.columns:
            styled = styled.map(_style_inv_share, subset=['Inv. %'])

        st.dataframe(styled, use_container_width=True, height=500)

        # Download
        csv = df.to_csv(sep=";", decimal=",").encode("utf-8")
        st.download_button(
            label="⬇️ Download als CSV",
            data=csv,
            file_name=f"bpa_base_stock_{date.today()}.csv",
            mime="text/csv",
        )

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 2 – SUBSCRIPTIES / IP / LEVERTIJD AANPASSEN
# ─────────────────────────────────────────────────────────────────────────────

with tab_subscripties:
    st.subheader("Subscripties per component")
    st.info(
        "Het aantal subscripties (Z) per component komt automatisch uit het "
        "werkelijke aantal klantlocaties; varieer prijs α en service level X in "
        "de tabs Verwachte subscripties / Sensitivity om het verwachte aantal abonnees te zien. "
        "Een vaste override per component kun je hieronder bij 'IP / Levertijd / "
        "Z aanpassen' instellen.",
        icon="ℹ️",
    )

    st.divider()
    st.subheader("Snelle actie — alle componenten in één keer")
    _bulk_col1, _bulk_col2 = st.columns([1, 2])
    with _bulk_col1:
        _bulk_z = st.number_input(
            "Z voor alle componenten", min_value=1, value=1, step=1, key="bulk_z_value",
        )
    with _bulk_col2:
        st.caption("Zet het aantal subscripties (Z) in één keer gelijk voor "
                   "alle componenten uit het huidige overzicht.")
    if st.button(f"🔁 Zet Z = {int(st.session_state.get('bulk_z_value', 1))} voor alle componenten"):
        _df_bulk = st.session_state.get("overzicht_df")
        if _df_bulk is None or _df_bulk.empty:
            st.warning("Geen overzicht beschikbaar — bereken eerst het overzicht in de tab 'Overzicht'.")
        else:
            # De artikelcode staat in de index van het overzicht (of in een 'Code'-kolom)
            if "Code" in _df_bulk.columns:
                _code_vals = _df_bulk["Code"]
            else:
                _code_vals = _df_bulk.index.to_series()
            cfg.setdefault("n_klanten_overrides", {})
            _codes = [str(c) for c in _code_vals.dropna().unique()]
            _z = int(_bulk_z)
            cfg["n_klanten_overrides"] = {c: _z for c in _codes}
            # Snapshot vóór recompute (zelfde patroon als bij 'Opslaan overrides')
            if "overzicht_df" in st.session_state:
                st.session_state.overzicht_df_prev = st.session_state.overzicht_df.copy()
            sla_config_op(cfg)
            st.toast(f"Z = {_z} gezet voor {len(_codes)} componenten.", icon="✅")
            st.session_state.pop("overzicht_df", None)
            st.rerun()

    st.divider()
    st.subheader("Overrides per artikelcode")
    st.caption("Z = aantal subscripties, IP = inkoopprijs (€), LT = levertijd (dagen). "
               "Laat een cel leeg om de Excel-waarde te gebruiken.")

    cfg.setdefault("ip_overrides", {})
    cfg.setdefault("lt_overrides", {})

    # Bouw gecombineerde tabel van alle codes met minstens één override
    alle_codes = sorted(
        set(cfg["n_klanten_overrides"]) |
        set(cfg["ip_overrides"]) |
        set(cfg["lt_overrides"])
    )
    override_rows = [
        {
            "Artikelcode": c,
            "N":           cfg["n_klanten_overrides"].get(c),
            "IP (€)":      cfg["ip_overrides"].get(c),
            "LT (dagen)":  cfg["lt_overrides"].get(c),
        }
        for c in alle_codes
    ]

    edited = st.data_editor(
        pd.DataFrame(override_rows) if override_rows else pd.DataFrame(
            columns=["Artikelcode", "N", "IP (€)", "LT (dagen)"]
        ),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Artikelcode": st.column_config.TextColumn("Artikelcode", required=True),
            "N":           st.column_config.NumberColumn("Z (subscripties)", min_value=1, step=1),
            "IP (€)":      st.column_config.NumberColumn("IP (€)", min_value=0.0, format="%.2f"),
            "LT (dagen)":  st.column_config.NumberColumn("LT (dagen)", min_value=1, step=1),
        },
        key="overrides_editor",
    )

    if st.button("💾 Opslaan overrides"):
        n_ov, ip_ov, lt_ov = {}, {}, {}
        for _, row in edited.iterrows():
            code = row.get("Artikelcode")
            if not code or pd.isna(code):
                continue
            code = str(code)
            if pd.notna(row["N"]):
                n_ov[code]  = int(row["N"])
            if pd.notna(row["IP (€)"]):
                ip_ov[code] = float(row["IP (€)"])
            if pd.notna(row["LT (dagen)"]):
                lt_ov[code] = int(row["LT (dagen)"])
        cfg["n_klanten_overrides"] = n_ov
        cfg["ip_overrides"]        = ip_ov
        cfg["lt_overrides"]        = lt_ov
        # Bewaar huidige overzicht_df als vorige snapshot vóór recompute
        if "overzicht_df" in st.session_state:
            st.session_state.overzicht_df_prev = st.session_state.overzicht_df.copy()
        sla_config_op(cfg)
        st.toast(f"Overrides opgeslagen — {len(n_ov)} Z, {len(ip_ov)} IP, {len(lt_ov)} LT.", icon="✅")
        st.session_state.pop("overzicht_df", None)
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 3 – COMPONENT TOEVOEGEN
# ─────────────────────────────────────────────────────────────────────────────

with tab_toevoegen:
    st.subheader("Nieuw component toevoegen")
    st.write("Gebruik dit voor componenten die nog niet in de Excel staan.")

    with st.form("form_toevoegen"):
        col1, col2 = st.columns(2)
        with col1:
            f_code  = st.text_input("Artikelcode *")
            f_descr = st.text_input("Omschrijving")
            f_lam   = st.number_input(
                "Lambda – vraag per jaar *",
                min_value=0.0001, value=1.0, step=0.1, format="%.4f",
            )
        with col2:
            f_lt = st.number_input(
                "Levertijd leverancier → BPA (dagen) *",
                min_value=1, value=30, step=1,
            )
            f_n = st.number_input(
                "Aantal subscripties (Z)",
                min_value=1, value=1, step=1,
            )
            f_ip = st.number_input(
                "Inkoopprijs (€)", min_value=0.0, value=0.0, step=10.0, format="%.2f",
            )
        submitted = st.form_submit_button("➕ Component opslaan")

    if submitted:
        if not f_code:
            st.error("Artikelcode is verplicht.")
        elif f_code in cfg["handmatige_componenten"]:
            st.warning(f"'{f_code}' bestaat al. Verwijder het eerst via het tabblad 'Component verwijderen'.")
        else:
            cfg["handmatige_componenten"][f_code] = {
                "descr":           f_descr,
                "lambda_per_jaar": float(f_lam),
                "lt_dagen":        int(f_lt),
                "n_klanten":       int(f_n),
                "ip":              float(f_ip),
            }
            sla_config_op(cfg)
            st.success(f"Component '{f_code}' toegevoegd.")

            # Preview berekende basisvoorraden
            lt_jr = int(f_lt) / 365
            preview = {
                f"s@{sl:.1%}": BPAOptimizationModel.inverse_service_level(sl, float(f_lam), lt_jr)
                for sl in SERVICE_LEVELS
            }
            st.write("**Berekende basisvoorraden voor dit component:**")
            st.dataframe(pd.DataFrame([preview]), use_container_width=False)

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 4 – COMPONENT VERWIJDEREN
# ─────────────────────────────────────────────────────────────────────────────

with tab_verwijderen:
    st.subheader("Component verwijderen uit model")
    st.write("Handmatig toegevoegde componenten worden permanent verwijderd. "
             "Excel-componenten worden uitgesloten (en kunnen later weer worden teruggezet).")

    handmatig   = cfg["handmatige_componenten"]
    uitgesloten = cfg.setdefault("uitgesloten_componenten", [])

    # Gebruik dezelfde codes als in het Overzicht-tab (= classificatie-whitelist
    # toegepast, inclusief synthetische classificatie-rijen).
    _ov_df = st.session_state.get("overzicht_df")
    if _ov_df is None or _ov_df.empty:
        try:
            _ov_df = get_overzicht_df(cfg)
            st.session_state["overzicht_df"] = _ov_df
        except Exception:
            _ov_df = pd.DataFrame()

    if _ov_df is not None and not _ov_df.empty and "bron" in _ov_df.columns:
        excel_codes = [str(c) for c, b in zip(_ov_df.index, _ov_df["bron"])
                       if b in ("excel", "classificatie")]
    else:
        excel_codes = []

    # Alle actieve codes met bron
    opties = (
        [(c, "handmatig", handmatig[c].get("descr", "")) for c in handmatig if c not in uitgesloten] +
        [(c, "excel",     "") for c in excel_codes if c not in handmatig and c not in uitgesloten]
    )

    if not opties:
        st.info("Geen actieve componenten om te verwijderen.")
    else:
        keuze = st.selectbox(
            "Selecteer component",
            options=[c for c, _, _ in opties],
            format_func=lambda c: next(
                f"{c}  [{bron}]  {descr}" for code, bron, descr in opties if code == c
            ),
        )
        bron_keuze = next(bron for c, bron, _ in opties if c == keuze)
        if bron_keuze == "handmatig":
            v = handmatig[keuze]
            st.write(f"**{keuze}** (handmatig) &nbsp;|&nbsp; λ = {v['lambda_per_jaar']:.4f}/jr "
                     f"&nbsp;|&nbsp; LT = {v['lt_dagen']} d")
            st.warning("Dit component wordt permanent verwijderd.")
            if st.button("🗑️ Verwijder permanent", type="primary"):
                del cfg["handmatige_componenten"][keuze]
                sla_config_op(cfg)
                st.success(f"'{keuze}' verwijderd.")
                st.rerun()
        else:
            st.write(f"**{keuze}** (uit Excel) – wordt uitgesloten van berekeningen.")
            st.info("Het artikel blijft in de Excel staan maar telt niet meer mee in het model.")
            if st.button("🚫 Uitsluiten van model", type="primary"):
                if keuze not in uitgesloten:
                    uitgesloten.append(keuze)
                sla_config_op(cfg)
                st.success(f"'{keuze}' uitgesloten.")
                st.rerun()

    # Uitgesloten Excel-componenten terugzetten
    if uitgesloten:
        st.divider()
        st.subheader("Uitgesloten componenten terugzetten")
        terugzetten = st.selectbox(
            "Selecteer component om terug te zetten",
            options=uitgesloten,
            key="terugzetten_selectbox",
        )
        if st.button("↩️ Zet terug in model"):
            uitgesloten.remove(terugzetten)
            sla_config_op(cfg)
            st.success(f"'{terugzetten}' is weer actief.")
            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 5 – CONFIGURATIE
# ─────────────────────────────────────────────────────────────────────────────

with tab_config:
    st.subheader("Huidige configuratie")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Mediane Z", representatieve_z())
    col2.metric("Z-overrides", len(cfg.get("n_klanten_overrides", {})))
    col3.metric("IP/LT-overrides",
                max(len(cfg.get("ip_overrides", {})), len(cfg.get("lt_overrides", {}))))
    col4.metric("Uitgesloten", len(cfg.get("uitgesloten_componenten", [])))
    st.write(f"Aangemaakt: `{cfg['aangemaakt']}`  |  Aangepast: `{cfg['aangepast']}`")

    if cfg["n_klanten_overrides"]:
        st.write("**Overrides:**")
        st.dataframe(
            pd.DataFrame([{"Code": k, "Z": v}
                          for k, v in cfg["n_klanten_overrides"].items()]),
            use_container_width=False,
        )

    if cfg["handmatige_componenten"]:
        st.write("**Handmatige componenten:**")
        st.dataframe(
            pd.DataFrame([
                {
                    "Code":        k,
                    "Omschrijving": v.get("descr", ""),
                    "λ/jr":        v["lambda_per_jaar"],
                    "LT(d)":       v["lt_dagen"],
                    "Z":           v.get("n_klanten", "std"),
                    "IP(€)":       v.get("ip", 0),
                }
                for k, v in cfg["handmatige_componenten"].items()
            ]),
            use_container_width=True,
        )

    st.divider()
    st.write("**Ruwe JSON:**")
    st.json(cfg)

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 6 – HISTORIEK
# ─────────────────────────────────────────────────────────────────────────────

with tab_historie:
    st.subheader("Historiek basisvoorraden")
    st.caption("Elke keer dat je een wijziging opslaat wordt automatisch een snapshot bewaard.")

    if not os.path.exists(HISTORY_PATH):
        st.info("Nog geen historiek beschikbaar. Sla een wijziging op om de eerste snapshot te maken.")
    else:
        with open(HISTORY_PATH, encoding="utf-8") as _f:
            history = json.load(_f)

        if not history:
            st.info("Nog geen snapshots.")
        else:
            # Bouw DataFrame op
            rows = []
            for h in history:
                row = {"Datum": h["datum"], "Z": h["n_klanten"], "# componenten": h["n_actief"]}
                row.update(h.get("totalen", {}))
                rows.append(row)
            hist_df = pd.DataFrame(rows).set_index("Datum")

            sl_cols = [c for c in hist_df.columns if c.startswith("s@")]

            # Grafiek
            if sl_cols:
                import matplotlib.pyplot as plt
                import matplotlib.ticker as ticker

                fig, ax = plt.subplots(figsize=(10, 4))
                for col in sl_cols:
                    ax.plot(hist_df.index, hist_df[col], marker="o", linewidth=2, label=col)
                    for x, y in zip(hist_df.index, hist_df[col]):
                        ax.annotate(str(int(y)), (x, y), textcoords="offset points",
                                    xytext=(0, 6), ha="center", fontsize=8)

                ax.set_xlabel("Update date", fontsize=11)
                ax.set_ylabel("Total base stock (units)", fontsize=11)
                ax.set_title("Total BPA base stock per update moment", fontsize=12)
                ax.legend(fontsize=9)
                ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
                ax.grid(True, alpha=0.3)
                plt.xticks(rotation=30, ha="right")
                fig.tight_layout()
                st.pyplot(fig)
                plt.close(fig)

            # Tabel
            st.write("**Snapshots:**")
            st.dataframe(hist_df.reset_index(), use_container_width=True)

            # Snapshot handmatig toevoegen (huidige staat)
            st.divider()
            if st.button("📸 Voeg snapshot toe van huidige staat"):
                from bpa_beheer import _sla_history_snapshot
                _sla_history_snapshot(cfg)
                st.success("Snapshot toegevoegd.")
                st.rerun()

    # ── Sensitivity grafieken ──────────────────────────────────────────────
    st.divider()
    st.subheader("Sensitivity grafieken")

    if "overzicht_df" not in st.session_state or st.session_state.overzicht_df.empty:
        st.info("Laad het overzicht (tabblad 📊) om de sensitivity grafieken te berekenen.")
    else:
        # Haal draaiknoppen op uit Kostenanalyse; gebruik defaults als nog niet berekend
        _kp = st.session_state.get('kosten_params', {})
        _ALPHA_DEF      = _kp.get('alpha',     0.15)
        _KAPPA_BPA_DEF  = _kp.get('kappa_bpa', 0.20)
        _KAPPA_C_DEF    = _kp.get('kappa_c',   0.25)

        st.caption(
            f"Vaste waarden buiten de gesweepte parameter: "
            f"α = **{_ALPHA_DEF:.0%}**, κ\\_BPA = **{_KAPPA_BPA_DEF:.0%}**, "
            f"κ\\_c = **{_KAPPA_C_DEF:.0%}**, N = standaard uit overzicht. "
            f"_(pas aan via tabblad 💰 Kostenanalyse)_"
        )

        _SL_SWEEP_S     = SERVICE_LEVELS
        _N_VALS         = [1, 2, 5, 10, 50]
        _ALPHA_SWEEP_S  = [0.05, 0.08, 0.10, 0.12, 0.15, 0.18, 0.20, 0.25, 0.30]
        _SL_ALPHA       = [sl for sl in SERVICE_LEVELS if sl >= 0.98]

        if st.button("📊 Bereken sensitivity grafieken"):
            _ov = st.session_state.overzicht_df
            _g1 = {n: [] for n in _N_VALS}
            _g2 = []
            _g3 = {sl: [] for sl in _SL_ALPHA}

            with st.spinner("Berekenen (kan even duren)…"):
                # Grafieken 1 & 2: sweep over service levels
                for _sl in _SL_SWEEP_S:
                    try:
                        _m2, _ = bouw_model_kosten(_ov, _ALPHA_DEF, _KAPPA_BPA_DEF, _KAPPA_C_DEF, _sl)
                        _g2.append({'sl': _sl, 'base': sum(_m2.calculate_base_stock_levels().values())})
                    except Exception:
                        _g2.append({'sl': _sl, 'base': None})
                    for _n in _N_VALS:
                        try:
                            _, _r1 = bouw_model_kosten(
                                _ov, _ALPHA_DEF, _KAPPA_BPA_DEF, _KAPPA_C_DEF, _sl,
                                n_klanten_override=_n,
                            )
                            _g1[_n].append({'sl': _sl, 'marge': _r1['bpa_margin']})
                        except Exception:
                            _g1[_n].append({'sl': _sl, 'marge': None})
                # Grafiek 3: sweep over alpha per SL ≥ 98%
                for _sl in _SL_ALPHA:
                    for _a in _ALPHA_SWEEP_S:
                        try:
                            _, _r3 = bouw_model_kosten(_ov, _a, _KAPPA_BPA_DEF, _KAPPA_C_DEF, _sl)
                            _g3[_sl].append({'alpha': _a, 'marge': _r3['bpa_margin']})
                        except Exception:
                            _g3[_sl].append({'alpha': _a, 'marge': None})

            st.session_state.sens_g1 = _g1
            st.session_state.sens_g2 = _g2
            st.session_state.sens_g3 = _g3

        if 'sens_g1' in st.session_state:
            import matplotlib.pyplot as _plt
            import matplotlib.ticker as _mt

            _COLORS5 = ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2', '#D32F2F']
            _COLORS4 = ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2']
            _fmt_eur = _mt.FuncFormatter(lambda v, _: f'€{v:,.0f}')
            _fmt_sl  = _mt.FuncFormatter(lambda v, _: f'{v:.2f}%')

            # ── Grafiek 1: service level vs marge per N ────────────────────
            _fig1, _ax1 = _plt.subplots(figsize=(10, 5))
            for _n, _col in zip(_N_VALS, _COLORS5):
                _pts = [(r['sl']*100, r['marge'])
                        for r in st.session_state.sens_g1[_n] if r['marge'] is not None]
                if _pts:
                    _ax1.plot([p[0] for p in _pts], [p[1] for p in _pts],
                              marker='o', linewidth=2, color=_col, label=f'N = {_n}')
            _ax1.axhline(0, color='grey', linewidth=0.8)
            _ax1.set_xlabel('Service level (%)', fontsize=11)
            _ax1.set_ylabel('Annual margin (€)', fontsize=11)
            _ax1.set_title(
                f'Margin vs. service level  '
                f'(α = {_ALPHA_DEF:.0%}, κ_BPA = {_KAPPA_BPA_DEF:.0%})',
                fontsize=12,
            )
            _ax1.yaxis.set_major_formatter(_fmt_eur)
            _ax1.xaxis.set_major_formatter(_fmt_sl)
            _ax1.set_xticks([sl*100 for sl in _SL_SWEEP_S])
            _ax1.legend(fontsize=9)
            _ax1.grid(True, alpha=0.3)
            _plt.setp(_ax1.get_xticklabels(), rotation=25, ha='right')
            _fig1.tight_layout()
            st.pyplot(_fig1)
            _plt.close(_fig1)

            # ── Grafiek 2: service level vs basisvoorraad ──────────────────
            _fig2, _ax2 = _plt.subplots(figsize=(10, 4))
            _pts2 = [(r['sl']*100, r['base'])
                     for r in st.session_state.sens_g2 if r['base'] is not None]
            if _pts2:
                _ax2.plot([p[0] for p in _pts2], [p[1] for p in _pts2],
                          marker='s', linewidth=2, color='#FF9800', label='Total S*')
                for _xv, _yv in _pts2:
                    _ax2.annotate(str(int(_yv)), (_xv, _yv),
                                  textcoords='offset points', xytext=(0, 7),
                                  ha='center', fontsize=9)
            _ax2.set_xlabel('Service level (%)', fontsize=11)
            _ax2.set_ylabel('Total base stock (units)', fontsize=11)
            _ax2.set_title(
                f'Base stock vs. service level  '
                f'(α = {_ALPHA_DEF:.0%}, N = standard)',
                fontsize=12,
            )
            _ax2.xaxis.set_major_formatter(_fmt_sl)
            _ax2.set_xticks([sl*100 for sl in _SL_SWEEP_S])
            _ax2.yaxis.set_major_locator(_mt.MaxNLocator(integer=True))
            _ax2.legend(fontsize=9)
            _ax2.grid(True, alpha=0.3)
            _plt.setp(_ax2.get_xticklabels(), rotation=25, ha='right')
            _fig2.tight_layout()
            st.pyplot(_fig2)
            _plt.close(_fig2)

            # ── Grafiek 3: alpha vs marge per service level ────────────────
            _fig3, _ax3 = _plt.subplots(figsize=(10, 5))
            for _sl3, _col3 in zip(_SL_ALPHA, _COLORS4):
                _pts3 = [(r['alpha']*100, r['marge'])
                         for r in st.session_state.sens_g3[_sl3] if r['marge'] is not None]
                if _pts3:
                    _ax3.plot([p[0] for p in _pts3], [p[1] for p in _pts3],
                              marker='o', linewidth=2, color=_col3, label=f'SL = {_sl3:.1%}')
            _ax3.axhline(0, color='grey', linewidth=0.8)
            _ax3.set_xlabel('Subscription rate α (%)', fontsize=11)
            _ax3.set_ylabel('Annual margin (€)', fontsize=11)
            _ax3.set_title(
                f'Margin vs. subscription rate  '
                f'(κ_BPA = {_KAPPA_BPA_DEF:.0%}, N = standard)',
                fontsize=12,
            )
            _ax3.yaxis.set_major_formatter(_fmt_eur)
            _ax3.xaxis.set_major_formatter(_mt.FuncFormatter(lambda v, _: f'{v:.0f}%'))
            _ax3.set_xticks([a*100 for a in _ALPHA_SWEEP_S])
            _ax3.legend(fontsize=9)
            _ax3.grid(True, alpha=0.3)
            _plt.setp(_ax3.get_xticklabels(), rotation=25, ha='right')
            _fig3.tight_layout()
            st.pyplot(_fig3)
            _plt.close(_fig3)

        # ── N vs. haalbaarheid per α ──────────────────────────────────────────────
        st.divider()
        st.subheader("Z vs. haalbaarheid per α")
        st.caption(
            "Effect van het aantal subscripties op de BPA-marge en haalbaarheid "
            "voor verschillende abonnementstarieven. "
            "SL, κ_BPA en κ_c worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        _N_FEAS_VALS = [1, 2, 3, 5, 7, 10, 15, 20, 30, 50, 75, 100]
        _ALPHA_FEAS  = [0.05, 0.08, 0.10, 0.12, 0.15, 0.18, 0.20, 0.25, 0.30]
        _COLORS_FEAS = [
            '#D32F2F', '#F57C00', '#FBC02D', '#8BC34A', '#388E3C',
            '#1976D2', '#7B1FA2', '#0097A7', '#5D4037',
        ]

        if st.button("📊 Bereken N vs. haalbaarheid"):
            _kp_f = st.session_state.get('kosten_params', {})
            _sl_f = _kp_f.get('service_level', 0.990)
            _kb_f = _kp_f.get('kappa_bpa',    0.20)
            _kc_f = _kp_f.get('kappa_c',      0.25)

            _nfeas = {a: [] for a in _ALPHA_FEAS}
            with st.spinner("Berekenen N vs. haalbaarheid…"):
                for _a_f in _ALPHA_FEAS:
                    for _n_f in _N_FEAS_VALS:
                        try:
                            _, _rf = bouw_model_kosten(
                                st.session_state.overzicht_df,
                                _a_f, _kb_f, _kc_f, _sl_f,
                                n_klanten_override=_n_f,
                            )
                            _nfeas[_a_f].append({
                                'n': _n_f,
                                'marge': _rf['bpa_margin'],
                                'feasible': _rf['feasible'],
                            })
                        except Exception:
                            _nfeas[_a_f].append({'n': _n_f, 'marge': None, 'feasible': False})

            st.session_state.sens_nfeas = _nfeas
            st.session_state.sens_nfeas_params = {
                'sl': _sl_f, 'kappa_bpa': _kb_f, 'kappa_c': _kc_f,
            }

        if 'sens_nfeas' in st.session_state:
            import matplotlib.pyplot as _plt_nf
            import matplotlib.ticker as _mt_nf
            import numpy as _np_nf

            _nfd   = st.session_state.sens_nfeas
            _nfp   = st.session_state.sens_nfeas_params
            _n_std = representatieve_z()

            # ── Grafiek 1: marge vs N per α ──────────────────────────────────────────
            _fig_nf, _ax_nf = _plt_nf.subplots(figsize=(11, 6))
            for _a_f, _col_f in zip(_ALPHA_FEAS, _COLORS_FEAS):
                _pts_nf = [(p['n'], p['marge']) for p in _nfd[_a_f] if p['marge'] is not None]
                if _pts_nf:
                    _xs_nf, _ys_nf = zip(*_pts_nf)
                    _ax_nf.plot(_xs_nf, _ys_nf, marker='o', linewidth=2,
                                color=_col_f, label=f'α = {_a_f:.0%}')
            _ax_nf.axhline(0, color='grey', linewidth=1.2, linestyle='--', label='Break-even')
            _ax_nf.axvline(_n_std, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std}')
            _ax_nf.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_nf.set_ylabel('Annual BPA margin (€)', fontsize=11)
            _ax_nf.set_title(
                f'BPA margin vs. Z per α  '
                f'(SL = {_nfp["sl"]:.1%}, '
                f'κ_BPA = {_nfp["kappa_bpa"]:.0%}, '
                f'κ_c = {_nfp["kappa_c"]:.0%})',
                fontsize=12,
            )
            _ax_nf.yaxis.set_major_formatter(
                _mt_nf.FuncFormatter(lambda v, _: f'€{v:,.0f}')
            )
            _ax_nf.set_xticks(_N_FEAS_VALS)
            _plt_nf.setp(_ax_nf.get_xticklabels(), rotation=30, ha='right')
            _ax_nf.legend(fontsize=9, ncol=2)
            _ax_nf.grid(True, alpha=0.3)
            _fig_nf.tight_layout()
            st.pyplot(_fig_nf)
            _plt_nf.close(_fig_nf)

            # ── Grafiek 2: haalbaarheids-heatmap (N × α) ───────────────────────────────
            _heat = _np_nf.array([
                [1.0 if p['feasible'] else 0.0 for p in _nfd[_a_f]]
                for _a_f in _ALPHA_FEAS
            ])
            _fig_hm, _ax_hm = _plt_nf.subplots(figsize=(11, 4))
            _ax_hm.imshow(_heat, aspect='auto', cmap='RdYlGn', vmin=0, vmax=1)
            _ax_hm.set_xticks(range(len(_N_FEAS_VALS)))
            _ax_hm.set_xticklabels(_N_FEAS_VALS, fontsize=9)
            _ax_hm.set_yticks(range(len(_ALPHA_FEAS)))
            _ax_hm.set_yticklabels([f'{a:.0%}' for a in _ALPHA_FEAS], fontsize=9)
            _ax_hm.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_hm.set_ylabel('α', fontsize=12)
            _ax_hm.set_title(
                'BPA feasibility per (N, α)  (✓ = feasible, ✗ = infeasible)',
                fontsize=12,
            )
            for _i in range(len(_ALPHA_FEAS)):
                for _j in range(len(_N_FEAS_VALS)):
                    _ax_hm.text(_j, _i,
                                '✓' if _heat[_i, _j] else '✗',
                                ha='center', va='center', fontsize=11,
                                color='#1a5c1a' if _heat[_i, _j] else '#7a0000')
            try:
                _n_idx = min(range(len(_N_FEAS_VALS)),
                             key=lambda k: abs(_N_FEAS_VALS[k] - _n_std))
                _ax_hm.axvline(_n_idx, color='black', linewidth=2.0, linestyle=':')
                _ax_hm.text(_n_idx + 0.15, -0.6, f'N={_n_std}', fontsize=8, color='black')
            except Exception:
                pass
            _fig_hm.tight_layout()
            st.pyplot(_fig_hm)
            _plt_nf.close(_fig_hm)
        # ── Haalbaarheid BPA per (N, SL) – heatmap ────────────────────────────────
        st.divider()
        st.subheader("Haalbaarheid BPA per (Z, serviceniveau)")
        st.caption(
            "Groen = BPA is haalbaar (marge ≥ 0), rood = niet haalbaar. "
            "α wordt overgenomen uit tabblad 💰 Kostenanalyse; κ_BPA en κ_c idem."
        )

        _N_NSL_VALS = [1, 2, 3, 5, 7, 10, 15, 20, 30, 50, 75, 100]

        if st.button("📊 Bereken haalbaarheid (N × SL)"):
            _kp_nsl = st.session_state.get('kosten_params', {})
            _a_nsl  = _kp_nsl.get('alpha',     0.15)
            _kb_nsl = _kp_nsl.get('kappa_bpa', 0.20)
            _kc_nsl = _kp_nsl.get('kappa_c',   0.25)

            _nsl_grid = {}
            with st.spinner("Berekenen haalbaarheid (N × SL)…"):
                for _n_nsl in _N_NSL_VALS:
                    _nsl_grid[_n_nsl] = {}
                    for _sl_nsl in SERVICE_LEVELS:
                        try:
                            _, _r_nsl = bouw_model_kosten(
                                st.session_state.overzicht_df,
                                _a_nsl, _kb_nsl, _kc_nsl, _sl_nsl,
                                n_klanten_override=_n_nsl,
                            )
                            _nsl_grid[_n_nsl][_sl_nsl] = {
                                'feasible': _r_nsl['feasible'],
                                'margin':   _r_nsl['bpa_margin'],
                            }
                        except Exception:
                            _nsl_grid[_n_nsl][_sl_nsl] = {'feasible': False, 'margin': None}

            st.session_state.sens_nsl_grid  = _nsl_grid
            st.session_state.sens_nsl_alpha = _a_nsl
            st.session_state.sens_nsl_kb    = _kb_nsl

        if 'sens_nsl_grid' in st.session_state:
            import matplotlib.pyplot as _plt_nsl
            import matplotlib.colors as _mcolors_nsl
            import numpy as _np_nsl

            _grid   = st.session_state.sens_nsl_grid
            _a_lbl  = st.session_state.sens_nsl_alpha
            _kb_lbl = st.session_state.sens_nsl_kb
            _n_std_nsl = representatieve_z()

            _rows_nsl = SERVICE_LEVELS       # y-as
            _cols_nsl = _N_NSL_VALS          # x-as

            # Bouw matrices: haalbaarheid (0/1) en genormaliseerde marge
            _feas_mat = _np_nsl.zeros((len(_rows_nsl), len(_cols_nsl)))
            _marg_mat = _np_nsl.full((len(_rows_nsl), len(_cols_nsl)), float('nan'))

            for _ci, _n_v in enumerate(_cols_nsl):
                for _ri, _sl_v in enumerate(_rows_nsl):
                    _cell = _grid.get(_n_v, {}).get(_sl_v, {})
                    _feas_mat[_ri, _ci] = 1.0 if _cell.get('feasible') else 0.0
                    if _cell.get('margin') is not None:
                        _marg_mat[_ri, _ci] = _cell['margin']

            # Kleurschaal: rood → geel → groen via marge-waarden
            _valid = _marg_mat[~_np_nsl.isnan(_marg_mat)]
            if len(_valid) > 0:
                _abs_max = max(abs(_valid.min()), abs(_valid.max()), 1)
            else:
                _abs_max = 1
            _norm_nsl = _mcolors_nsl.TwoSlopeNorm(
                vmin=-_abs_max, vcenter=0, vmax=_abs_max
            )

            _fig_nsl, _ax_nsl = _plt_nsl.subplots(figsize=(13, 5))
            _im_nsl = _ax_nsl.imshow(
                _marg_mat, aspect='auto',
                cmap='RdYlGn', norm=_norm_nsl,
                interpolation='nearest',
            )
            _plt_nsl.colorbar(_im_nsl, ax=_ax_nsl, label='BPA margin (€)', fraction=0.03, pad=0.02)

            # Annotaties per cel
            for _ri, _sl_v in enumerate(_rows_nsl):
                for _ci, _n_v in enumerate(_cols_nsl):
                    _cell = _grid.get(_n_v, {}).get(_sl_v, {})
                    _feas = _cell.get('feasible', False)
                    _mg   = _cell.get('margin')
                    _sym  = '✓' if _feas else '✗'
                    _tc   = '#1a5c1a' if _feas else '#7a0000'
                    _ax_nsl.text(_ci, _ri, _sym,
                                 ha='center', va='center' if _mg is None else 'bottom',
                                 fontsize=13, color=_tc, fontweight='bold')
                    if _mg is not None:
                        _ax_nsl.text(_ci, _ri + 0.28, f'€{_mg:,.0f}',
                                     ha='center', va='center', fontsize=6.5, color=_tc)

            # Assen
            _ax_nsl.set_xticks(range(len(_cols_nsl)))
            _ax_nsl.set_xticklabels([str(n) for n in _cols_nsl], fontsize=9)
            _ax_nsl.set_yticks(range(len(_rows_nsl)))
            _ax_nsl.set_yticklabels([f'{sl:.1%}' for sl in _rows_nsl], fontsize=9)
            _ax_nsl.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_nsl.set_ylabel('Service level', fontsize=11)
            _ax_nsl.set_title(
                f'BPA feasibility per (Z, service level)  '
                f'(α = {_a_lbl:.0%}, κ_BPA = {_kb_lbl:.0%})',
                fontsize=12,
            )

            # Markeer huidige N
            try:
                _ni_std = min(range(len(_cols_nsl)),
                              key=lambda k: abs(_cols_nsl[k] - _n_std_nsl))
                _ax_nsl.axvline(_ni_std, color='black', linewidth=2.0, linestyle=':')
                _ax_nsl.text(_ni_std + 0.15, -0.7, f'Z={_n_std_nsl}',
                             fontsize=8, color='black')
            except Exception:
                pass

            _fig_nsl.tight_layout()
            st.pyplot(_fig_nsl)
            _plt_nsl.close(_fig_nsl)

        # ── N vs. maximaal haalbaar serviceniveau ───────────────────────────────────
        st.divider()
        st.subheader("Z vs. maximaal haalbaar serviceniveau")
        st.caption(
            "Voor elk aantal subscripties (Z): wat is het hoogste service level waarbij het "
            "model nog haalbaar is? α, κ_BPA en κ_c worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        if st.button("📊 Bereken N vs. max. haalbaar SL"):
            _kp_sl  = st.session_state.get('kosten_params', {})
            _a_sl   = _kp_sl.get('alpha',    0.15)
            _kb_sl  = _kp_sl.get('kappa_bpa', 0.20)
            _kc_sl  = _kp_sl.get('kappa_c',   0.25)

            _nsl_results = []
            with st.spinner("Berekenen N vs. max. haalbaar SL…"):
                for _n_sl in _N_FEAS_VALS:
                    _row_sl = {'n': _n_sl}
                    _max_sl = None
                    _base_stocks = {}
                    for _sl_v in SERVICE_LEVELS:
                        try:
                            _m_sl, _r_sl = bouw_model_kosten(
                                st.session_state.overzicht_df,
                                _a_sl, _kb_sl, _kc_sl, _sl_v,
                                n_klanten_override=_n_sl,
                            )
                            _row_sl[_sl_v] = _r_sl['feasible']
                            if _r_sl['feasible']:
                                _max_sl = _sl_v
                            _base_stocks[_sl_v] = sum(_m_sl.calculate_base_stock_levels().values())
                        except Exception:
                            _row_sl[_sl_v] = False
                            _base_stocks[_sl_v] = None
                    _row_sl['max_feasible_sl'] = _max_sl
                    _row_sl['base_stocks'] = _base_stocks
                    _nsl_results.append(_row_sl)

            st.session_state.sens_nsl = _nsl_results
            st.session_state.sens_nsl_params = {
                'alpha': _a_sl, 'kappa_bpa': _kb_sl, 'kappa_c': _kc_sl,
            }

        if 'sens_nsl' in st.session_state:
            import matplotlib.pyplot as _plt_sl
            import matplotlib.ticker as _mt_sl
            import numpy as _np_sl

            _nsl_d    = st.session_state.sens_nsl
            _nsl_p    = st.session_state.sens_nsl_params
            _n_std_sl = representatieve_z()

            # ── Grafiek 2: totale S* vs N per service level ─────────────────────────
            _COLORS_SL = ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2']
            _fig_bs, _ax_bs = _plt_sl.subplots(figsize=(11, 5))
            for _sl_v, _col_sl in zip(SERVICE_LEVELS, _COLORS_SL):
                _bs_pts = [(r['n'], r['base_stocks'].get(_sl_v))
                           for r in _nsl_d
                           if r['base_stocks'].get(_sl_v) is not None]
                if _bs_pts:
                    _xb, _yb = zip(*_bs_pts)
                    _ax_bs.plot(_xb, _yb, marker='o', linewidth=2,
                                color=_col_sl, label=f'SL = {_sl_v:.1%}')
                    for _xv, _yv in zip(_xb, _yb):
                        _ax_bs.annotate(str(int(_yv)), (_xv, _yv),
                                        textcoords='offset points', xytext=(0, 7),
                                        ha='center', fontsize=8)
            _ax_bs.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std_sl}')
            _ax_bs.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_bs.set_ylabel('Total base stock S* (units)', fontsize=11)
            _ax_bs.set_title(
                f'Total S* vs. Z per service level  '
                f'(α = {_nsl_p["alpha"]:.0%})',
                fontsize=12,
            )
            _ax_bs.set_xticks(_N_FEAS_VALS)
            _plt_sl.setp(_ax_bs.get_xticklabels(), rotation=30, ha='right')
            _ax_bs.yaxis.set_major_locator(_mt_sl.MaxNLocator(integer=True))
            _ax_bs.legend(fontsize=9)
            _ax_bs.grid(True, alpha=0.3)
            _fig_bs.tight_layout()
            st.pyplot(_fig_bs)
            _plt_sl.close(_fig_bs)


            # ── ΔS* per extra subscriptie (pooling-effect) ────────────────────────────
            _fig_ds, _ax_ds = _plt_sl.subplots(figsize=(11, 5))
            for _sl_v, _col_sl in zip(SERVICE_LEVELS, _COLORS_SL):
                _bs_all = [
                    (r['n'], r['base_stocks'].get(_sl_v))
                    for r in _nsl_d
                    if r['base_stocks'].get(_sl_v) is not None
                ]
                if len(_bs_all) >= 2:
                    _nd, _dd = [], []
                    for (_n1, _s1), (_n2, _s2) in zip(_bs_all[:-1], _bs_all[1:]):
                        _nd.append((_n1 + _n2) / 2)
                        _dd.append((_s2 - _s1) / (_n2 - _n1))
                    _ax_ds.plot(_nd, _dd, marker='o', linewidth=2,
                                color=_col_sl, label=f'SL = {_sl_v:.1%}')
                    for _xv, _yv in zip(_nd, _dd):
                        _ax_ds.annotate(f'{_yv:.3f}', (_xv, _yv),
                                        textcoords='offset points', xytext=(0, 7),
                                        ha='center', fontsize=7)
            _ax_ds.axhline(0, color='grey', linewidth=0.8, linestyle='--')
            _ax_ds.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std_sl}')
            _ax_ds.set_xlabel('Number of subscriptions Z (interval midpoint)', fontsize=11)
            _ax_ds.set_ylabel('ΔS* / ΔN  (units per extra subscription)', fontsize=11)
            _ax_ds.set_title(
                f'Pooling effect: extra stock per extra subscription  '
                f'(α = {_nsl_p["alpha"]:.0%})',
                fontsize=12,
            )
            _ax_ds.legend(fontsize=9)
            _ax_ds.grid(True, alpha=0.3)
            _fig_ds.tight_layout()
            st.pyplot(_fig_ds)
            _plt_sl.close(_fig_ds)

            # ── Pooling-effect: analytische curves ──────────────────────────────
            # Bereken mu_per_sub = sum_i( lambda_i/N_i * L_i ) uit overzicht
            _ov_pa = st.session_state.overzicht_df.reset_index()
            _mu_per_sub = 0.0
            for _, _rpa in _ov_pa.iterrows():
                _ni = float(_rpa.get('n_klanten', 0) or 0)
                _li = float(_rpa.get('lambda_jr', 0) or 0)
                _lt = float(_rpa.get('LT_dagen', 0) or 0)
                if _ni > 0 and _li > 0 and _lt > 0:
                    _mu_per_sub += _li / _ni * _lt / 365.0

            _N_arr   = [float(r['n']) for r in _nsl_d]
            _mu_arr  = [n * _mu_per_sub for n in _N_arr]  # mu_total(N)

            # ── Plot 1: CV(N) = 1/sqrt(mu_total(N)) ───────────────────────
            import math as _math_pa
            _cv_arr = [1.0 / _math_pa.sqrt(mu) if mu > 0 else None for mu in _mu_arr]

            _fig_cv, _ax_cv = _plt_sl.subplots(figsize=(11, 4))
            _n_cv_ok = [n for n, cv in zip(_N_arr, _cv_arr) if cv is not None]
            _cv_ok   = [cv for cv in _cv_arr if cv is not None]
            if _n_cv_ok:
                _ax_cv.plot(_n_cv_ok, _cv_ok, marker='o', linewidth=2.5,
                            color='#1976D2', label='CV(Z) = 1/√(μₜₒₜ(Z))')
                for _xv, _yv in zip(_n_cv_ok, _cv_ok):
                    _ax_cv.annotate(f'{_yv:.3f}', (_xv, _yv),
                                    textcoords='offset points', xytext=(0, 7),
                                    ha='center', fontsize=8)
            _ax_cv.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std_sl}')
            _ax_cv.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_cv.set_ylabel('CV lead-time demand', fontsize=11)
            _ax_cv.set_title(
                'Relative uncertainty lead-time demand  '
                'CV(Z) = 1 / √(Z · Σ λᵢ · Lᵢ / Zᵢ)',
                fontsize=12,
            )
            _ax_cv.set_xticks(_N_FEAS_VALS)
            _plt_sl.setp(_ax_cv.get_xticklabels(), rotation=30, ha='right')
            _ax_cv.legend(fontsize=9)
            _ax_cv.grid(True, alpha=0.3)
            _fig_cv.tight_layout()
            st.pyplot(_fig_cv)
            _plt_sl.close(_fig_cv)

            # ── Plot 2: S*(X,N) / N per service level ──────────────────────
            _fig_sp, _ax_sp = _plt_sl.subplots(figsize=(11, 5))
            for _sl_v, _col_sl in zip(SERVICE_LEVELS, _COLORS_SL):
                _sp_pts = [
                    (r['n'], r['base_stocks'].get(_sl_v) / r['n'])
                    for r in _nsl_d
                    if r['base_stocks'].get(_sl_v) is not None and r['n'] > 0
                ]
                if _sp_pts:
                    _xp, _yp = zip(*_sp_pts)
                    _ax_sp.plot(_xp, _yp, marker='o', linewidth=2,
                                color=_col_sl, label=f'SL = {_sl_v:.1%}')
                    for _xv, _yv in zip(_xp, _yp):
                        _ax_sp.annotate(f'{_yv:.2f}', (_xv, _yv),
                                        textcoords='offset points', xytext=(0, 7),
                                        ha='center', fontsize=7)
            # Also plot mu/N = mu_per_sub as reference (pooling limit)
            _ax_sp.axhline(_mu_per_sub, color='grey', linewidth=1.0,
                           linestyle='--', label=f'μ/Z (lead-time demand per sub., ≈{_mu_per_sub:.3f})')
            _ax_sp.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std_sl}')
            _ax_sp.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_sp.set_ylabel('S*(X,Z) / Z  (stock per subscription)', fontsize=11)
            _ax_sp.set_title('Required stock per subscription  S*(X,Z) / Z',
                             fontsize=12)
            _ax_sp.set_xticks(_N_FEAS_VALS)
            _plt_sl.setp(_ax_sp.get_xticklabels(), rotation=30, ha='right')
            _ax_sp.legend(fontsize=9)
            _ax_sp.grid(True, alpha=0.3)
            _fig_sp.tight_layout()
            st.pyplot(_fig_sp)
            _plt_sl.close(_fig_sp)

            # ── S*(X,N)/N detail: N = 1 … 20 ────────────────────────────────
            st.caption(
                "Detail: S*(X,Z) / Z voor Z = 1 … 20, "
                "berekend per component via inverse_service_level."
            )
            if st.button("📊 Bereken S*/Z voor Z = 1 … 20"):
                _ov_det = st.session_state.overzicht_df.reset_index()
                # Collect (lambda_per_N, lt_jr) per component
                _comp_det = []
                for _, _rd in _ov_det.iterrows():
                    _ni = float(_rd.get('n_klanten', 0) or 0)
                    _li = float(_rd.get('lambda_jr', 0) or 0)
                    _lt = float(_rd.get('LT_dagen', 0) or 0)
                    if _ni > 0 and _li > 0 and _lt > 0:
                        _comp_det.append((_li / _ni, _lt / 365.0))

                _N_DET = list(range(1, 21))
                _det_results = {sl: [] for sl in SERVICE_LEVELS}
                with st.spinner("Berekenen S*/Z voor Z = 1 … 20…"):
                    for _n_det in _N_DET:
                        for _sl_det in SERVICE_LEVELS:
                            _s_tot = sum(
                                BPAOptimizationModel.inverse_service_level(
                                    _sl_det, _lam_pn * _n_det, _lt_c
                                )
                                for _lam_pn, _lt_c in _comp_det
                            )
                            _det_results[_sl_det].append(
                                {'n': _n_det, 's_per_n': _s_tot / _n_det}
                            )
                st.session_state.sens_spn_det = _det_results
                st.session_state.sens_spn_mu = _mu_per_sub

            if 'sens_spn_det' in st.session_state:
                _spn_d = st.session_state.sens_spn_det
                _mu_ref = st.session_state.sens_spn_mu
                _N_DET_x = list(range(1, 21))

                _fig_det, _ax_det = _plt_sl.subplots(figsize=(11, 5))
                for _sl_det, _col_det in zip(SERVICE_LEVELS, _COLORS_SL):
                    _pts_det = [(r['n'], r['s_per_n']) for r in _spn_d[_sl_det]]
                    if _pts_det:
                        _xd, _yd = zip(*_pts_det)
                        _ax_det.plot(_xd, _yd, marker='o', linewidth=2,
                                     color=_col_det, label=f'SL = {_sl_det:.1%}')
                        for _xv, _yv in zip(_xd, _yd):
                            _ax_det.annotate(f'{_yv:.2f}', (_xv, _yv),
                                             textcoords='offset points', xytext=(0, 7),
                                             ha='center', fontsize=7)
                _ax_det.axhline(_mu_ref, color='grey', linewidth=1.0, linestyle='--',
                                label=f'μ/Z (lower bound, ≈{_mu_ref:.3f})')
                _ax_det.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                                label=f'Z current = {_n_std_sl}')
                _ax_det.set_xlabel('Number of subscriptions Z', fontsize=11)
                _ax_det.set_ylabel('S*(X,Z) / Z  (stock per subscription)', fontsize=11)
                _ax_det.set_title(
                    'Required stock per subscription  S*(X,Z) / Z  (Z = 1… 20)',
                    fontsize=12,
                )
                _ax_det.set_xticks(_N_DET_x)
                _ax_det.legend(fontsize=9)
                _ax_det.grid(True, alpha=0.3)
                _fig_det.tight_layout()
                st.pyplot(_fig_det)
                _plt_sl.close(_fig_det)


            # ── Plot 3: Safety stock / N per service level ──────────────────
            _fig_ss, _ax_ss = _plt_sl.subplots(figsize=(11, 5))
            for _sl_v, _col_sl in zip(SERVICE_LEVELS, _COLORS_SL):
                _ss_pts = []
                for r in _nsl_d:
                    _bs = r['base_stocks'].get(_sl_v)
                    _mu_n = r['n'] * _mu_per_sub
                    if _bs is not None and r['n'] > 0:
                        _ss_pts.append((r['n'], (_bs - _mu_n) / r['n']))
                if _ss_pts:
                    _xs, _ys = zip(*_ss_pts)
                    _ax_ss.plot(_xs, _ys, marker='o', linewidth=2,
                                color=_col_sl, label=f'SL = {_sl_v:.1%}')
                    for _xv, _yv in zip(_xs, _ys):
                        _ax_ss.annotate(f'{_yv:.3f}', (_xv, _yv),
                                        textcoords='offset points', xytext=(0, 7),
                                        ha='center', fontsize=7)
            _ax_ss.axhline(0, color='grey', linewidth=0.8, linestyle='--')
            _ax_ss.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'Z current = {_n_std_sl}')
            _ax_ss.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_ss.set_ylabel('(S*(X,Z) − μ(Z)) / Z  (safety stock per subscription)',
                              fontsize=11)
            _ax_ss.set_title(
                'Safety stock per subscription  (S*(X,Z) − Z · ΣλᵢLᵢ/Zᵢ) / Z',
                fontsize=12,
            )
            _ax_ss.set_xticks(_N_FEAS_VALS)
            _plt_sl.setp(_ax_ss.get_xticklabels(), rotation=30, ha='right')
            _ax_ss.legend(fontsize=9)
            _ax_ss.grid(True, alpha=0.3)
            _fig_ss.tight_layout()
            st.pyplot(_fig_ss)
            _plt_sl.close(_fig_ss)

        # ── Marginale kosten vs. N (pooling-effect) ─────────────────────
        st.divider()
        st.subheader("Marginale kosten vs. Z")
        st.caption(
            "Hoe nemen de incrementele kosten per extra subscriptie af naarmate N groeit? "
            "Dit visualiseert het pooling-effect: elke extra subscriptie vereist minder "
            "extra inventariskosten dan de vorige. "
            "α, κ_BPA, κ_c en SL worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        _N_MC_VALS   = [1, 2, 3, 5, 7, 10, 15, 20, 30, 50, 75, 100]
        _SL_MC_COLORS = ['#2ca02c', '#ff7f0e', '#1f77b4', '#d62728', '#9467bd', '#8c564b']

        if st.button("📊 Bereken marginale kosten vs. Z"):
            _kp_mc = st.session_state.get('kosten_params', {})
            _a_mc  = _kp_mc.get('alpha',     0.15)
            _kb_mc = _kp_mc.get('kappa_bpa', 0.20)
            _kc_mc = _kp_mc.get('kappa_c',   0.25)

            _mc_results_by_sl = {}
            with st.spinner("Berekenen marginale kosten voor alle serviceniveaus…"):
                for _sl_mc in SERVICE_LEVELS:
                    _mc_rows = []
                    for _n_mc in _N_MC_VALS:
                        try:
                            _m_mc, _r_mc = bouw_model_kosten(
                                st.session_state.overzicht_df,
                                _a_mc, _kb_mc, _kc_mc, _sl_mc,
                                n_klanten_override=_n_mc,
                            )
                            _mc_rows.append({
                                'n':      _n_mc,
                                'cost':   _r_mc['bpa_costs'],
                                'margin': _r_mc['bpa_margin'],
                            })
                        except Exception:
                            _mc_rows.append({'n': _n_mc, 'cost': None, 'margin': None})
                    _mc_results_by_sl[_sl_mc] = _mc_rows

            st.session_state.sens_mc = _mc_results_by_sl
            st.session_state.sens_mc_params = {
                'alpha': _a_mc, 'kappa_bpa': _kb_mc, 'kappa_c': _kc_mc,
            }

        if 'sens_mc' in st.session_state:
            import matplotlib.pyplot as _plt_mc
            import matplotlib.ticker as _mt_mc

            _mc_by_sl = st.session_state.sens_mc
            _mc_p     = st.session_state.sens_mc_params
            _n_std_mc = representatieve_z()
            _fmt_mc   = _mt_mc.FuncFormatter(lambda v, _: f'€{v:,.0f}')

            # Backward compatibility: oud formaat was een lijst voor één SL
            if isinstance(_mc_by_sl, list):
                _mc_by_sl = {0.990: _mc_by_sl}

            _fig_mc, (_ax_mc1, _ax_mc2) = _plt_mc.subplots(
                2, 1, figsize=(11, 10), sharex=False
            )

            for (_sl_mc, _mc_d), _color in zip(_mc_by_sl.items(), _SL_MC_COLORS):
                _mc_ok = [(r['n'], r['cost']) for r in _mc_d if r['cost'] is not None]
                if not _mc_ok:
                    continue
                _ns_mc, _cs_mc = zip(*_mc_ok)
                _sl_lbl = f'SL {_sl_mc:.1%}'

                _ax_mc1.plot(_ns_mc, _cs_mc, marker='o', linewidth=2,
                             color=_color, label=_sl_lbl)

                _n_mid_mc, _delta_c_mc = [], []
                for (_n1, _c1), (_n2, _c2) in zip(_mc_ok[:-1], _mc_ok[1:]):
                    _n_mid_mc.append((_n1 + _n2) / 2)
                    _delta_c_mc.append((_c2 - _c1) / (_n2 - _n1))
                _ax_mc2.plot(_n_mid_mc, _delta_c_mc, marker='o', linewidth=2,
                             color=_color, label=_sl_lbl)

            _ax_mc1.axvline(_n_std_mc, color='black', linewidth=1.0,
                            linestyle=':', label=f'Z current = {_n_std_mc}')
            _ax_mc1.set_ylabel('Total BPA cost (€)', fontsize=11)
            _ax_mc1.set_title(
                f'BPA cost vs. Z  '
                f'(α = {_mc_p["alpha"]:.0%}, κ_BPA = {_mc_p["kappa_bpa"]:.0%})',
                fontsize=12,
            )
            _ax_mc1.yaxis.set_major_formatter(_fmt_mc)
            _ax_mc1.set_xticks(_N_MC_VALS)
            _ax_mc1.legend(fontsize=9)
            _ax_mc1.grid(True, alpha=0.3)

            _ax_mc2.axhline(0, color='grey', linewidth=0.8, linestyle='--')
            _ax_mc2.axvline(_n_std_mc, color='black', linewidth=1.0,
                            linestyle=':', label=f'Z current = {_n_std_mc}')
            _ax_mc2.set_xlabel('Number of subscriptions (Z)', fontsize=11)
            _ax_mc2.set_ylabel('ΔC / ΔZ  (extra cost per extra subscription, €)',
                               fontsize=11)
            _ax_mc2.set_title(
                'Pooling effect: marginal inventory cost per extra subscription',
                fontsize=12,
            )
            _ax_mc2.yaxis.set_major_formatter(_fmt_mc)
            _ax_mc2.set_xticks(_N_MC_VALS)
            _ax_mc2.legend(fontsize=9)
            _ax_mc2.grid(True, alpha=0.3)

            _fig_mc.tight_layout(h_pad=3)
            st.pyplot(_fig_mc)
            _plt_mc.close(_fig_mc)

        # ── Investering vs. N ─────────────────────────────────────────────────
        st.divider()
        st.subheader("Investering vs. aantal subscripties")
        st.caption(
            "Totale voorraadwaarde (Σ S\u002a × inkoopprijs) als functie van het aantal "
            "subscripties per service level. Toont hoeveel kapitaal BPA in voorraad "
            "moet investeren naarmate het klantenbestand groeit. De x-as toont het "
            "TOTALE aantal subscripties over alle componenten en start bij de som van "
            "de verwachte sub-aantallen (E[Z]); alle componenten schalen van daaruit "
            "proportioneel mee omhoog."
        )

        # Baseline per component = het geconfigureerde n_klanten, dat de
        # verwachte E[Z_i(α,X)] weergeeft zodra die via de tab Verwachte
        # subscripties is doorgezet. De x-as toont het TOTAAL aantal
        # subscripties over alle componenten (= som van de baselines bij
        # factor 1.0); alle componenten schalen proportioneel mee.
        _sim_base_inv = {}
        # Groeifactoren: 20 punten van 1.0x (huidig totaal) tot 2.9x in stappen van 0.1.
        _INV_FACTORS = [round(1.0 + 0.1 * _k, 1) for _k in range(20)]
        _COLORS_INV = ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2']

        if st.button("📊 Bereken investering vs. totaal subs"):
            _ov_inv = st.session_state.overzicht_df.reset_index()
            # Verzamel per component: lambda per subscriptie, baseline-subs, LT, IP, VP, code
            _comp_inv = []
            for _, _ri in _ov_inv.iterrows():
                _ni = float(_ri.get('n_klanten', 0) or 0)
                _li = float(_ri.get('lambda_jr', 0) or 0)
                _lt = float(_ri.get('LT_dagen', 0) or 0)
                _ip = float(_ri.get('IP', 0) or 0)
                _vp = float(_ri.get('VP', 0) or 0)
                if _ni > 0 and _li > 0 and _lt > 0:
                    _code = str(_ri.get('Code', ''))
                    _comp_inv.append({
                        'code':        _code,
                        'descr':       str(_ri.get('Descr', '')),
                        'lam_per_sub': _li / _ni,
                        'n_base':      float(_sim_base_inv.get(_code, _ni)),
                        'lt_jr':       _lt / 365,
                        'ip':          _ip,
                        'vp':          _vp,
                    })

            # Totaal subs bij factor 1.0 = som van de (verwachte) baselines.
            _T0_inv = sum(_c['n_base'] for _c in _comp_inv)

            _inv_results = {sl: [] for sl in SERVICE_LEVELS}
            # Per-component resultaten voor top-5 grafiek (alle SL's)
            _sl_top = st.session_state.get('kosten_params', {}).get('service_level', 0.990)
            _inv_per_comp = {sl: {c['code']: [] for c in _comp_inv} for sl in SERVICE_LEVELS}

            with st.spinner("Berekenen investering vs. totaal subs…"):
                for _f_inv in _INV_FACTORS:
                    _tot_subs = int(round(_T0_inv * _f_inv))   # x-waarde: totaal subscripties
                    for _sl_inv in SERVICE_LEVELS:
                        _totaal = sum(
                            BPAOptimizationModel.inverse_service_level(
                                _sl_inv, _c['lam_per_sub'] * _c['n_base'] * _f_inv, _c['lt_jr']
                            ) * _c['ip']
                            for _c in _comp_inv
                        )
                        _inv_results[_sl_inv].append({'n': _tot_subs, 'inv': _totaal})
                    # Per-component per SL (voor top-5/top-10 grafiek)
                    for _c in _comp_inv:
                        for _sl_c in SERVICE_LEVELS:
                            _s = BPAOptimizationModel.inverse_service_level(
                                _sl_c, _c['lam_per_sub'] * _c['n_base'] * _f_inv, _c['lt_jr']
                            )
                            _inv_per_comp[_sl_c][_c['code']].append({'n': _tot_subs, 'inv': _s * _c['ip']})

            # Top 5 / Top 10 duurste componenten op VP
            _top5_codes  = sorted(_comp_inv, key=lambda c: c['vp'], reverse=True)[:5]
            _top10_codes = sorted(_comp_inv, key=lambda c: c['vp'], reverse=True)[:10]

            st.session_state.sens_inv        = _inv_results
            st.session_state.sens_inv_comp   = _inv_per_comp
            st.session_state.sens_inv_top5   = _top5_codes
            st.session_state.sens_inv_top10  = _top10_codes
            st.session_state.sens_inv_sl_top = _sl_top
            st.session_state.sens_inv_t0     = int(round(_T0_inv))

        if 'sens_inv' in st.session_state:
            import matplotlib.pyplot as _plt_inv
            import matplotlib.ticker as _mt_inv

            _inv_d    = st.session_state.sens_inv
            _tot0_inv = int(st.session_state.get('sens_inv_t0', 0))
            _x_ticks_inv = [r['n'] for r in _inv_d[SERVICE_LEVELS[0]]]
            _fmt_inv  = _mt_inv.FuncFormatter(lambda v, _: f'€{v:,.0f}')

            _fig_inv, _ax_inv = _plt_inv.subplots(figsize=(11, 5))
            for _sl_inv, _col_inv in zip(SERVICE_LEVELS, _COLORS_INV):
                _pts_inv = [(r['n'], r['inv']) for r in _inv_d[_sl_inv] if r['inv'] is not None]
                if _pts_inv:
                    _xi, _yi = zip(*_pts_inv)
                    _ax_inv.plot(_xi, _yi, marker='o', linewidth=2,
                                 color=_col_inv, label=f'SL = {_sl_inv:.1%}')

            _ax_inv.axvline(_tot0_inv, color='black', linewidth=1.0, linestyle=':',
                            label=f'Total subs (sim) = {_tot0_inv}')
            _ax_inv.set_xlabel('Total number of subscriptions (all components)', fontsize=11)
            _ax_inv.set_ylabel('Total inventory value (€)', fontsize=11)
            _ax_inv.set_title(
                'Required investment in base stock vs. total number of subscriptions',
                fontsize=12,
            )
            _ax_inv.yaxis.set_major_formatter(_fmt_inv)
            _ax_inv.set_xticks(_x_ticks_inv)
            _plt_inv.setp(_ax_inv.get_xticklabels(), rotation=30, ha='right')
            _ax_inv.legend(fontsize=9)
            _ax_inv.grid(True, alpha=0.3)
            _fig_inv.tight_layout()
            st.pyplot(_fig_inv)
            _plt_inv.close(_fig_inv)

            # Tabel: investering per totaal aantal subs en SL
            _inv_tbl_rows = []
            for _n_v in _x_ticks_inv:
                _row_t = {'Totaal subs': _n_v}
                for _sl_v in SERVICE_LEVELS:
                    _pts = [r for r in _inv_d[_sl_v] if r['n'] == _n_v]
                    _row_t[f'SL {_sl_v:.1%}'] = f"€{_pts[0]['inv']:,.0f}" if _pts else '—'
                _inv_tbl_rows.append(_row_t)
            st.dataframe(pd.DataFrame(_inv_tbl_rows).set_index('Totaal subs'), use_container_width=False)

            # ── Top-5 duurste componenten per VP ──────────────────────────
            if 'sens_inv_top5' in st.session_state:
                _top5    = st.session_state.sens_inv_top5
                _comp_d  = st.session_state.sens_inv_comp
                _sl_lbl  = st.session_state.sens_inv_sl_top

                _COLORS_TOP5 = ['#D32F2F', '#F57C00', '#FBC02D', '#388E3C', '#1976D2']
                st.subheader("Top 5 duurste componenten (VP) — investering vs. totaal subs (per service level)")
                st.caption(
                    "Gesommeerde investeringswaarde (S\u002a × IP) van de top 5 duurste componenten "
                    "(op verkoopprijs) als functie van het totaal aantal subscripties, per service level."
                )

                # Haal x-waarden op uit de data
                _comp_d_sl0 = _comp_d.get(SERVICE_LEVELS[0], {})
                _x5_vals = [p['n'] for p in _comp_d_sl0.get(_top5[0]['code'], [])] if _top5 else []

                if _x5_vals:
                    _fig_t5, _ax_t5 = _plt_inv.subplots(figsize=(11, 5))

                    for _sl_t5, _col_t5, _ls_t5 in zip(
                            SERVICE_LEVELS, _COLORS_INV, ['-', '--', '-.', ':']):
                        _cd_sl = _comp_d.get(_sl_t5, {})
                        _tot_sl = [
                            sum(
                                next((p['inv'] for p in _cd_sl.get(_c5['code'], [])
                                      if p['n'] == _nv), 0)
                                for _c5 in _top5
                            )
                            for _nv in _x5_vals
                        ]
                        _ax_t5.plot(_x5_vals, _tot_sl, color=_col_t5, marker='o',
                                    linewidth=2.0, linestyle=_ls_t5,
                                    label=f'SL {_sl_t5:.1%}')

                    _ax_t5.axvline(_tot0_inv, color='black', linewidth=1.0, linestyle=':',
                                   label=f'Total subs (sim) = {_tot0_inv}')
                    _ax_t5.set_xlabel('Total number of subscriptions (all components)', fontsize=11)
                    _ax_t5.set_ylabel('Summed investment value top 5 (€)', fontsize=11)
                    _ax_t5.set_title(
                        'Top 5 most expensive components (VP): summed investment vs. total subs per SL',
                        fontsize=12,
                    )
                    _ax_t5.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t5.set_xticks(_x5_vals)
                    _plt_inv.setp(_ax_t5.get_xticklabels(), rotation=30, ha='right')
                    _ax_t5.legend(fontsize=9, loc='upper left')
                    _ax_t5.grid(True, alpha=0.3)
                    _fig_t5.tight_layout()
                    st.pyplot(_fig_t5)
                    _plt_inv.close(_fig_t5)

            # ── Top-10 duurste componenten — gesommeerde lijnen per SL ──────
            if 'sens_inv_top10' in st.session_state:
                import matplotlib as _mpl_inv
                _top10   = st.session_state.sens_inv_top10
                _comp_d  = st.session_state.sens_inv_comp

                _comp_d_sl0_t10 = _comp_d.get(SERVICE_LEVELS[0], {})
                _x10_sum_vals = [p['n'] for p in _comp_d_sl0_t10.get(_top10[0]['code'], [])] if _top10 else []

                st.subheader("Top 10 duurste componenten (VP) — gesommeerde investering vs. totaal subs (per service level)")
                st.caption(
                    "Gesommeerde investeringswaarde (S\u002a × IP) van de top 10 duurste componenten "
                    "(op verkoopprijs) als functie van het totaal aantal subscripties, per service level."
                )

                if _x10_sum_vals:
                    _fig_t10s, _ax_t10s = _plt_inv.subplots(figsize=(11, 5))
                    for _sl_t10s, _col_t10s, _ls_t10s in zip(
                            SERVICE_LEVELS, _COLORS_INV, ['-', '--', '-.', ':']):
                        _cd10s = _comp_d.get(_sl_t10s, {})
                        _tot10s = [
                            sum(
                                next((p['inv'] for p in _cd10s.get(_c10['code'], [])
                                      if p['n'] == _nv), 0)
                                for _c10 in _top10
                            )
                            for _nv in _x10_sum_vals
                        ]
                        _ax_t10s.plot(_x10_sum_vals, _tot10s, color=_col_t10s, marker='o',
                                      linewidth=2.0, linestyle=_ls_t10s,
                                      label=f'SL {_sl_t10s:.1%}')
                    _ax_t10s.axvline(_tot0_inv, color='black', linewidth=1.0, linestyle=':',
                                     label=f'Total subs (sim) = {_tot0_inv}')
                    _ax_t10s.set_xlabel('Total number of subscriptions (all components)', fontsize=11)
                    _ax_t10s.set_ylabel('Summed investment value top 10 (€)', fontsize=11)
                    _ax_t10s.set_title(
                        'Top 10 most expensive components (VP): summed investment vs. total subs per SL',
                        fontsize=12,
                    )
                    _ax_t10s.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t10s.set_xticks(_x10_sum_vals)
                    _plt_inv.setp(_ax_t10s.get_xticklabels(), rotation=30, ha='right')
                    _ax_t10s.legend(fontsize=9, loc='upper left')
                    _ax_t10s.grid(True, alpha=0.3)
                    _fig_t10s.tight_layout()
                    st.pyplot(_fig_t10s)
                    _plt_inv.close(_fig_t10s)

                # ── Top-10 duurste componenten — individuele lijnen ──────────
                _sl_opts_t10 = [f'SL {s:.1%}' for s in SERVICE_LEVELS]
                _sl_sel_t10  = st.selectbox(
                    'Service level voor top-10 grafiek',
                    _sl_opts_t10,
                    index=1,
                    key='top10_sl_select',
                )
                _sl_val_t10 = SERVICE_LEVELS[_sl_opts_t10.index(_sl_sel_t10)]

                st.subheader("Top 10 duurste componenten (VP) — investering per component vs. totaal subs")
                st.caption(
                    "Investeringswaarde (S\u002a \u00d7 IP) per component als functie van het totaal aantal subscripties "
                    "voor het geselecteerde service level."
                )

                _cd10_sl = _comp_d.get(_sl_val_t10, {})
                _x10_vals = [p['n'] for p in _cd10_sl.get(_top10[0]['code'], [])] if _top10 else []

                if _x10_vals:
                    _cmap10  = _mpl_inv.colormaps['tab10']
                    _fig_t10, _ax_t10 = _plt_inv.subplots(figsize=(12, 5))

                    for _ci, _c10 in enumerate(_top10):
                        _pts10 = [p['inv'] for p in _cd10_sl.get(_c10['code'], [])]
                        if _pts10:
                            _lbl10 = f"{_c10['code']} – {_c10.get('descr', '')[:25]}"
                            _ax_t10.plot(
                                _x10_vals, _pts10,
                                color=_cmap10(_ci / 10),
                                marker='o', linewidth=1.8, markersize=5,
                                label=_lbl10,
                            )

                    _ax_t10.axvline(_tot0_inv, color='black', linewidth=1.0, linestyle=':',
                                    label=f'Total subs (sim) = {_tot0_inv}')
                    _ax_t10.set_xlabel('Total number of subscriptions (all components)', fontsize=11)
                    _ax_t10.set_ylabel('Investment value per component (€)', fontsize=11)
                    _ax_t10.set_title(
                        f'Top 10 most expensive components — investment vs. total subs  ({_sl_sel_t10})',
                        fontsize=12,
                    )
                    _ax_t10.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t10.set_xticks(_x10_vals)
                    _plt_inv.setp(_ax_t10.get_xticklabels(), rotation=30, ha='right')
                    _ax_t10.legend(fontsize=8, loc='upper left', ncol=2)
                    _ax_t10.grid(True, alpha=0.3)
                    _fig_t10.tight_layout()
                    st.pyplot(_fig_t10)
                    _plt_inv.close(_fig_t10)

        # ── Marge over tijd (groeiend klantenbestand) ─────────────────────────
        st.divider()
        st.subheader("Marge over tijd (groeiend klantenbestand)")
        st.caption(
            "Cumulatieve cashflow rekening houdend met de initiële investering in basisvoorraad, "
            "jaarlijkse voorraadkosten (κ\\_BPA × S\\* × IP) en groeiende abonnementsinkomsten "
            "(α × VP × Z). Z groeit lineair van start naar doelstelling. "
            "α, κ\\_BPA en SL worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        _col_mt1, _col_mt2, _col_mt3 = st.columns(3)
        with _col_mt1:
            _N_start_mt = st.number_input(
                "Z start (jaar 0)", min_value=1,
                value=representatieve_z(), step=1,
                key="marge_tijd_n_start",
            )
        with _col_mt2:
            _N_end_mt = st.number_input(
                "Z eind (doelstelling)", min_value=1,
                value=max(representatieve_z() * 3, 10), step=1,
                key="marge_tijd_n_end",
            )
        with _col_mt3:
            _T_mt = st.number_input(
                "Tijdshorizon (jaar)", min_value=1, max_value=20,
                value=5, step=1,
                key="marge_tijd_T",
            )

        if st.button("📊 Bereken marge over tijd"):
            _kp_mt    = st.session_state.get('kosten_params', {})
            _alpha_mt = _kp_mt.get('alpha',         0.15)
            _kbpa_mt  = _kp_mt.get('kappa_bpa',     0.20)
            _sl_mt    = _kp_mt.get('service_level',  0.990)

            _ov_mt = st.session_state.overzicht_df.reset_index()
            _comp_mt = []
            for _, _rmt in _ov_mt.iterrows():
                _ni = float(_rmt.get('n_klanten', 0) or 0)
                _li = float(_rmt.get('lambda_jr',  0) or 0)
                _lt = float(_rmt.get('LT_dagen',   0) or 0)
                _ip = float(_rmt.get('IP',          0) or 0)
                _vp = float(_rmt.get('VP',          0) or 0)
                if _ni > 0 and _li > 0 and _lt > 0:
                    _comp_mt.append({
                        'code':        str(_rmt.get('Code', '')),
                        'descr':       str(_rmt.get('Descr', '')),
                        'lam_per_sub': _li / _ni,
                        'lt_jr':       _lt / 365,
                        'ip':          _ip,
                        'vp':          _vp,
                    })

            _comp_mt_top10 = sorted(_comp_mt, key=lambda c: c['vp'], reverse=True)[:10]

            def _N_t_mt(t):
                if _T_mt == 0:
                    return float(_N_start_mt)
                return _N_start_mt + (_N_end_mt - _N_start_mt) * t / _T_mt

            def _inv_waarde_for(comps, n):
                return sum(
                    BPAOptimizationModel.inverse_service_level(
                        _sl_mt, _c['lam_per_sub'] * n, _c['lt_jr']
                    ) * _c['ip'] 
                    for _c in comps
                )

            def _inv_waarde_mt(n):
                return _inv_waarde_for(_comp_mt, n)

            def _rev_jaar_for(comps, n):
                return sum(_c['vp'] * _alpha_mt * n for _c in comps)

            def _rev_jaar_mt(n):
                return _rev_jaar_for(_comp_mt, n)

            def _calc_cashflow(comps):
                _res_inv, _res_hold, _res_rev, _res_cum, _res_N = [], [], [], [], []
                _cum_cf = 0.0
                for _t in _t_arr_mt:
                    _n   = _N_t_mt(_t)
                    _iw  = _inv_waarde_for(comps, _n)
                    _hld = _kbpa_mt * _iw
                    _rv  = _rev_jaar_for(comps, _n)
                    if _t == 0:
                        _di = _iw
                    else:
                        _di = max(0.0, _iw - _inv_waarde_for(comps, _N_t_mt(_t - 1)))
                    _cum_cf += _rv - _hld - _di
                    _res_N.append(round(_n, 1))
                    _res_inv.append(-_di)
                    _res_hold.append(-_hld)
                    _res_rev.append(_rv)
                    _res_cum.append(_cum_cf)
                return _res_inv, _res_hold, _res_rev, _res_cum, _res_N

            _t_arr_mt = list(range(int(_T_mt) + 1))
            _params_mt = {
                'alpha':     _alpha_mt,
                'kappa_bpa': _kbpa_mt,
                'sl':        _sl_mt,
                'N_start':   int(_N_start_mt),
                'N_end':     int(_N_end_mt),
                'T':         int(_T_mt),
            }

            with st.spinner("Berekenen marge over tijd…"):
                _inv_mt_arr, _hold_mt_arr, _rev_mt_arr, _cum_mt_arr, _N_mt_arr = \
                    _calc_cashflow(_comp_mt)

                # Per top-10 component afzonderlijk
                _top10_comp_cf = {}
                for _c10 in _comp_mt_top10:
                    _i10, _h10, _r10, _cum10, _ = _calc_cashflow([_c10])
                    _top10_comp_cf[_c10['code']] = {
                        'descr': _c10['descr'],
                        'inv':   _i10, 'hold': _h10, 'rev': _r10, 'cum': _cum10,
                    }

                # Gesommeerde top-10
                _i10s, _h10s, _r10s, _cum10s, _ = _calc_cashflow(_comp_mt_top10)

            st.session_state.sens_marge_tijd = {
                't':      _t_arr_mt,
                'inv':    _inv_mt_arr,
                'hold':   _hold_mt_arr,
                'rev':    _rev_mt_arr,
                'cum':    _cum_mt_arr,
                'N':      _N_mt_arr,
                'params': _params_mt,
            }
            st.session_state.sens_marge_tijd_top10 = {
                't':        _t_arr_mt,
                'sum_inv':  _i10s,
                'sum_hold': _h10s,
                'sum_rev':  _r10s,
                'sum_cum':  _cum10s,
                'N':        _N_mt_arr,
                'comp':     _top10_comp_cf,
                'params':   _params_mt,
            }

        if 'sens_marge_tijd' in st.session_state:
            import matplotlib.pyplot as _plt_mt
            import matplotlib.ticker as _mt_tick
            import numpy as _np_mt

            _mtd    = st.session_state.sens_marge_tijd
            _mtp    = _mtd['params']
            _fmt_mt = _mt_tick.FuncFormatter(lambda v, _: f'€{v:,.0f}')

            _t_arr    = _mtd['t']
            _x_pos    = _np_mt.arange(len(_t_arr))
            _x_lbl    = [f'Jaar {t}' for t in _t_arr]
            _inv_arr  = _np_mt.array(_mtd['inv'])
            _hold_arr = _np_mt.array(_mtd['hold'])
            _rev_arr  = _np_mt.array(_mtd['rev'])
            _cum_arr  = _np_mt.array(_mtd['cum'])
            _net_arr  = _rev_arr + _hold_arr + _inv_arr

            _fig_mt, (_ax_mt1, _ax_mt2) = _plt_mt.subplots(2, 1, figsize=(12, 10))

            # ── Grafiek 1: jaarlijkse cashflow ──────────────────────────────────
            _w = 0.55
            _ax_mt1.bar(_x_pos, _rev_arr,  _w,
                        label='Revenue',      color='#388E3C', alpha=0.85)
            _ax_mt1.bar(_x_pos, _hold_arr, _w,
                        label='Inventory cost', color='#F57C00', alpha=0.85)
            _ax_mt1.bar(_x_pos, _inv_arr,  _w, bottom=_hold_arr,
                        label='Investment',    color='#D32F2F', alpha=0.85)
            _ax_mt1.plot(_x_pos, _net_arr, color='black', marker='o',
                         linewidth=1.8, linestyle='--', label='Net year', zorder=5)
            _ax_mt1.axhline(0, color='grey', linewidth=0.8)

            _ax_mt1b = _ax_mt1.twinx()
            _ax_mt1b.plot(_x_pos, _mtd['N'], color='#1976D2', marker='s',
                          linewidth=1.5, linestyle=':', alpha=0.7, label='Z')
            _ax_mt1b.set_ylabel('Z (subscriptions)', fontsize=10, color='#1976D2')
            _ax_mt1b.tick_params(axis='y', labelcolor='#1976D2')

            _ax_mt1.set_xticks(_x_pos)
            _ax_mt1.set_xticklabels(_x_lbl, rotation=30, ha='right', fontsize=9)
            _ax_mt1.set_ylabel('Cash flow per year (€)', fontsize=11)
            _ax_mt1.set_title(
                f'Annual cash flow  (α = {_mtp["alpha"]:.0%}, '
                f'κ_BPA = {_mtp["kappa_bpa"]:.0%}, SL = {_mtp["sl"]:.1%}, '
                f'Z: {_mtp["N_start"]} → {_mtp["N_end"]})',
                fontsize=12,
            )
            _ax_mt1.yaxis.set_major_formatter(_fmt_mt)
            _h1, _l1 = _ax_mt1.get_legend_handles_labels()
            _h2, _l2 = _ax_mt1b.get_legend_handles_labels()
            _ax_mt1.legend(_h1 + _h2, _l1 + _l2, fontsize=9, loc='lower right')
            _ax_mt1.grid(True, axis='y', alpha=0.3)

            # ── Grafiek 2: cumulatieve cashflow ─────────────────────────────────
            _bar_colors_mt = ['#388E3C' if v >= 0 else '#D32F2F' for v in _cum_arr]
            _ax_mt2.bar(_x_pos, _cum_arr, _w, color=_bar_colors_mt, alpha=0.65)
            _ax_mt2.plot(_x_pos, _cum_arr, color='#1976D2', marker='o',
                         linewidth=2.0, linestyle='-', label='Cumulative CF')
            _ax_mt2.axhline(0, color='grey', linewidth=1.0, linestyle='--')

            _be_titel = 'Break-even not reached within time horizon'
            for _bi in range(1, len(_cum_arr)):
                if _cum_arr[_bi - 1] < 0 <= _cum_arr[_bi]:
                    _frac     = -_cum_arr[_bi - 1] / (_cum_arr[_bi] - _cum_arr[_bi - 1])
                    _be_x     = _bi - 1 + _frac
                    _be_titel = f'Break-even ≈ year {_be_x:.1f}'
                    _ax_mt2.axvline(_be_x, color='#F57C00', linewidth=1.8,
                                    linestyle=':', label=_be_titel)
                    break
            if _cum_arr[0] >= 0:
                _be_titel = 'Break-even in year 0 (immediately profitable)'

            _ax_mt2.set_xticks(_x_pos)
            _ax_mt2.set_xticklabels(_x_lbl, rotation=30, ha='right', fontsize=9)
            _ax_mt2.set_ylabel('Cumulative cash flow (€)', fontsize=11)
            _ax_mt2.set_title(f'Cumulative cash flow — {_be_titel}', fontsize=12)
            _ax_mt2.yaxis.set_major_formatter(_fmt_mt)
            _ax_mt2.legend(fontsize=9)
            _ax_mt2.grid(True, axis='y', alpha=0.3)

            _fig_mt.tight_layout(h_pad=3)
            st.pyplot(_fig_mt)
            _plt_mt.close(_fig_mt)

            st.dataframe(pd.DataFrame([
                {
                    'Jaar':               _t_arr[_ti],
                    'Z':                  _mtd['N'][_ti],
                    'Investering (€)':    f"€{-_inv_arr[_ti]:,.0f}",
                    'Voorraadkosten (€)': f"€{-_hold_arr[_ti]:,.0f}",
                    'Inkomsten (€)':      f"€{_rev_arr[_ti]:,.0f}",
                    'Netto (€)':          f"€{_net_arr[_ti]:+,.0f}",
                    'Cumulatief (€)':     f"€{_cum_arr[_ti]:+,.0f}",
                }
                for _ti in range(len(_t_arr))
            ]).set_index('Jaar'), use_container_width=False)

        # ── Marge over tijd — top 10 duurste componenten ───────────────────────
        if 'sens_marge_tijd_top10' in st.session_state:
            import matplotlib.pyplot as _plt_mt10
            import matplotlib.ticker as _mt10_tick
            import matplotlib as _mt10_mpl
            import numpy as _np_mt10

            _mt10d  = st.session_state.sens_marge_tijd_top10
            _mt10p  = _mt10d['params']
            _fmt10  = _mt10_tick.FuncFormatter(lambda v, _: f'€{v:,.0f}')
            _t10arr = _mt10d['t']
            _xp10   = _np_mt10.arange(len(_t10arr))
            _xl10   = [f'Jaar {t}' for t in _t10arr]
            _w10    = 0.55

            _si10  = _np_mt10.array(_mt10d['sum_inv'])
            _sh10  = _np_mt10.array(_mt10d['sum_hold'])
            _sr10  = _np_mt10.array(_mt10d['sum_rev'])
            _sc10  = _np_mt10.array(_mt10d['sum_cum'])
            _sn10  = _mt10d['sum_rev']  # not used below, using _sr10
            _net10 = _sr10 + _sh10 + _si10

            st.divider()
            st.subheader("Marge over tijd — top 10 duurste componenten (gesommeerd)")
            st.caption(
                f"Dezelfde cashflow-analyse maar enkel voor de 10 duurste componenten op VP. "
                f"SL = {_mt10p['sl']:.1%}, α = {_mt10p['alpha']:.0%}, "
                f"κ_BPA = {_mt10p['kappa_bpa']:.0%}."
            )

            _fig10s, (_ax10s1, _ax10s2) = _plt_mt10.subplots(2, 1, figsize=(12, 10))

            _ax10s1.bar(_xp10, _sr10, _w10, label='Revenue',      color='#388E3C', alpha=0.85)
            _ax10s1.bar(_xp10, _sh10, _w10, label='Inventory cost', color='#F57C00', alpha=0.85)
            _ax10s1.bar(_xp10, _si10, _w10, bottom=_sh10, label='Investment', color='#D32F2F', alpha=0.85)
            _ax10s1.plot(_xp10, _net10, color='black', marker='o', linewidth=1.8, linestyle='--', label='Net year', zorder=5)
            _ax10s1.axhline(0, color='grey', linewidth=0.8)
            _ax10s1b = _ax10s1.twinx()
            _ax10s1b.plot(_xp10, _mt10d['N'], color='#1976D2', marker='s', linewidth=1.5, linestyle=':', alpha=0.7, label='Z')
            _ax10s1b.set_ylabel('Z (subscriptions)', fontsize=10, color='#1976D2')
            _ax10s1b.tick_params(axis='y', labelcolor='#1976D2')
            _ax10s1.set_xticks(_xp10)
            _ax10s1.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10s1.set_ylabel('Cash flow per year (€)', fontsize=11)
            _ax10s1.set_title(
                f'Annual cash flow top 10  (α={_mt10p["alpha"]:.0%}, '
                f'κ_BPA={_mt10p["kappa_bpa"]:.0%}, SL={_mt10p["sl"]:.1%}, '
                f'Z: {_mt10p["N_start"]} → {_mt10p["N_end"]})', fontsize=12)
            _ax10s1.yaxis.set_major_formatter(_fmt10)
            _h10a, _l10a = _ax10s1.get_legend_handles_labels()
            _h10b, _l10b = _ax10s1b.get_legend_handles_labels()
            _ax10s1.legend(_h10a + _h10b, _l10a + _l10b, fontsize=9, loc='lower right')
            _ax10s1.grid(True, axis='y', alpha=0.3)

            _bc10 = ['#388E3C' if v >= 0 else '#D32F2F' for v in _sc10]
            _ax10s2.bar(_xp10, _sc10, _w10, color=_bc10, alpha=0.65)
            _ax10s2.plot(_xp10, _sc10, color='#1976D2', marker='o', linewidth=2.0, label='Cumulative CF')
            _ax10s2.axhline(0, color='grey', linewidth=1.0, linestyle='--')
            _be10t = 'Break-even not reached within time horizon'
            for _bi10 in range(1, len(_sc10)):
                if _sc10[_bi10 - 1] < 0 <= _sc10[_bi10]:
                    _f10   = -_sc10[_bi10 - 1] / (_sc10[_bi10] - _sc10[_bi10 - 1])
                    _bex10 = _bi10 - 1 + _f10
                    _be10t = f'Break-even ≈ year {_bex10:.1f}'
                    _ax10s2.axvline(_bex10, color='#F57C00', linewidth=1.8, linestyle=':', label=_be10t)
                    break
            if _sc10[0] >= 0:
                _be10t = 'Break-even in year 0 (immediately profitable)'
            _ax10s2.set_xticks(_xp10)
            _ax10s2.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10s2.set_ylabel('Cumulative cash flow (€)', fontsize=11)
            _ax10s2.set_title(f'Cumulative cash flow top 10 — {_be10t}', fontsize=12)
            _ax10s2.yaxis.set_major_formatter(_fmt10)
            _ax10s2.legend(fontsize=9)
            _ax10s2.grid(True, axis='y', alpha=0.3)
            _fig10s.tight_layout(h_pad=3)
            st.pyplot(_fig10s)
            _plt_mt10.close(_fig10s)

            # ── Cumulatieve cashflow per component ──────────────────────────
            st.subheader("Cumulatieve cashflow per component (top 10)")
            st.caption("Elke lijn toont de cumulatieve netto cashflow van één component afzonderlijk.")

            _comp_cf10 = _mt10d['comp']
            _cmap10mt  = _mt10_mpl.colormaps['tab10']

            _fig10c, _ax10c = _plt_mt10.subplots(figsize=(12, 5))
            for _ci10, (_code10, _cdata10) in enumerate(_comp_cf10.items()):
                _cum10c = _np_mt10.array(_cdata10['cum'])
                _lbl10c = f"{_code10} – {_cdata10['descr'][:25]}"
                _ax10c.plot(_xp10, _cum10c, color=_cmap10mt(_ci10 / 10),
                            marker='o', linewidth=1.8, markersize=5, label=_lbl10c)
            _ax10c.axhline(0, color='grey', linewidth=1.0, linestyle='--')
            _ax10c.set_xticks(_xp10)
            _ax10c.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10c.set_ylabel('Cumulative cash flow (€)', fontsize=11)
            _ax10c.set_title(
                f'Cumulative cash flow per component — top 10  (SL={_mt10p["sl"]:.1%})',
                fontsize=12,
            )
            _ax10c.yaxis.set_major_formatter(_fmt10)
            _ax10c.legend(fontsize=8, loc='lower right', ncol=2)
            _ax10c.grid(True, alpha=0.3)
            _fig10c.tight_layout()
            st.pyplot(_fig10c)
            _plt_mt10.close(_fig10c)


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 7 – KOSTENANALYSE
# ─────────────────────────────────────────────────────────────────────────────

with tab_kosten:
    st.subheader("Kostenanalyse BPA")
    st.caption(
        "Berekent BPA-kosten, omzet, marge en α-interval per component "
        "op basis van het huidige overzicht en de gekozen draaiknoppen."
    )

    if "overzicht_df" not in st.session_state or st.session_state.overzicht_df.empty:
        st.warning("Laad eerst het overzicht via het tabblad 📊 Overzicht.")
    else:
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            k_alpha = st.number_input(
                "α (abonnementstarief, %)",
                min_value=1.0, max_value=50.0, value=15.0, step=1.0, format="%.0f",
                help="Abonnementsprijs als percentage van verkoopprijs",
            ) / 100
        with col_b:
            k_kappa_bpa = st.number_input(
                "κ_BPA (%)",
                min_value=1.0, max_value=100.0, value=20.0, step=1.0, format="%.0f",
                help="κ_BPA = financiering + opslag + obsolescence (BPA)",
            ) / 100
        with col_c:
            k_kappa_c = st.number_input(
                "κ_c (%)",
                min_value=1.0, max_value=100.0, value=25.0, step=1.0, format="%.0f",
                help="κ_c = financiering + opslag + obsolescence (klant)",
            ) / 100
        with col_d:
            k_sl = st.selectbox(
                "Service level",
                options=SERVICE_LEVELS,
                index=SERVICE_LEVELS.index(0.990) if 0.990 in SERVICE_LEVELS else 0,
                format_func=lambda v: f"{v:.1%}",
            )

        if st.button("💰 Bereken kosten"):
            with st.spinner("Kostenmodel berekenen…"):
                try:
                    _m, _r = bouw_model_kosten(
                        st.session_state.overzicht_df,
                        alpha=k_alpha,
                        kappa_bpa=k_kappa_bpa,
                        kappa_c=k_kappa_c,
                        service_level=k_sl,
                    )
                    st.session_state.kosten_result = (_m, _r)
                    st.session_state.kosten_params = {
                        'alpha': k_alpha, 'kappa_bpa': k_kappa_bpa,
                        'kappa_c': k_kappa_c, 'service_level': k_sl,
                    }
                except Exception as _e:
                    st.error(f"Fout bij berekening: {_e}")

        if "kosten_result" in st.session_state:
            _m, _r = st.session_state.kosten_result
            _iv = _r['alpha_intervals']
            _p  = st.session_state.kosten_params

            # ── Samenvatting ───────────────────────────────────────────────
            _c1, _c2, _c3, _c4 = st.columns(4)
            _c1.metric("Haalbaar",    "✓ JA"  if _r['feasible'] else "✗ NEE")
            _c2.metric("Totale omzet",  f"€ {_r['total_revenue']:,.0f}")
            _c3.metric("BPA kosten",    f"€ {_r['bpa_costs']:,.0f}")
            _c4.metric("Marge",         f"€ {_r['bpa_margin']:+,.0f}")

            _al = _iv['universal_alpha_L']
            _au = _iv['universal_alpha_U']
            if _al is not None:
                st.info(
                    f"Universeel α-interval: **[{_al:.4%} – {_au:.4%}]**  "
                    f"{'✓ Haalbaar' if _iv['universal_feasible'] else '✗ Niet haalbaar'}"
                )

            # ── Per-component kosten tabel ─────────────────────────────────
            st.subheader("Kosten per component")
            _det = _m.calculate_detailed_bpa_costs()
            _bsl = _m.calculate_base_stock_levels()
            _per = _iv['per_component']
            _lt  = _m.parameters['lead_time']

            _rows = []
            for _code in _m.sets['spare_parts']:
                _d = _det[_code]
                _pc = _per.get(_code, {})
                _al = _pc.get('alpha_L')
                _au = _pc.get('alpha_U')
                _ok = (
                    _al is not None and _au is not None
                    and _al <= _p['alpha'] <= _au
                )
                _rows.append({
                    'Code':       _code,
                    'S*':         _bsl.get(_code, 0),
                    'Λ_BPA':      round(_d['demand'], 4),
                    'μ=Λ·L':      round(_d['demand'] * _lt.get(_code, 0), 4),
                    'C_BPA (€)':  round(_d['total'], 2),
                    'Omzet (€)':  round(_r['revenue_by_part'].get(_code, 0), 2),
                    'Marge (€)':  round(_r['revenue_by_part'].get(_code, 0) - _d['total'], 2),
                    'α_L,i':      f"{_al:.3%}" if _al is not None else '—',
                    'α_U,i':      f"{_au:.3%}" if _au is not None else '—',
                    'OK':         '✓' if _ok else '✗',
                })
            _tbl = pd.DataFrame(_rows).set_index('Code')

            st.dataframe(
                _tbl.style.format({
                    'S*':        '{:.0f}',
                    'Λ_BPA':    '{:.4f}',
                    'μ=Λ·L':    '{:.4f}',
                    'C_BPA (€)': '€ {:,.2f}',
                    'Omzet (€)': '€ {:,.2f}',
                    'Marge (€)': '€ {:+,.2f}',
                }),
                use_container_width=True,
                height=420,
            )
            st.write(
                f"**Totaal:** S\\* = {int(_tbl['S*'].sum())}  |  "
                f"C\\_BPA = € {_tbl['C_BPA (€)'].sum():,.2f}  |  "
                f"Omzet = € {_tbl['Omzet (€)'].sum():,.2f}  |  "
                f"Marge = € {_tbl['Marge (€)'].sum():+,.2f}"
            )

            # ── Klantbesparingen ───────────────────────────────────────────
            with st.expander("Klantbesparingen"):
                _klant_rows = [
                    {
                        'Klant':              _cust,
                        'Eigen kosten (€)':   b['self_stocking_cost'],
                        'BPA abonnement (€)': b['bpa_service_cost'],
                        'Besparing (€)':      b['savings'],
                        'Voordeel':           '✓' if b['benefits'] else '✗',
                    }
                    for _cust, b in _r['customer_benefits'].items()
                ]
                st.dataframe(
                    pd.DataFrame(_klant_rows).set_index('Klant').style.format({
                        'Eigen kosten (€)':   '€ {:,.2f}',
                        'BPA abonnement (€)': '€ {:,.2f}',
                        'Besparing (€)':      '€ {:+,.2f}',
                    }),
                    use_container_width=True,
                )

# ─────────────────────────────────────────────────────────────────────────────────
#  TAB 8 – SUBSCRIPTIEDREMPEL
# ─────────────────────────────────────────────────────────────────────────────────

with tab_drempel:
    st.subheader("Subscriptiedrempel per component")
    st.caption(
        "Per component: hoeveel extra subscripties zijn er nodig voordat S\u002a met 1 stijgt? "
        "Aanname: λ schaalt lineair met Z (λ = Z × λ_huidig / Z_huidig). "
        "Van toepassing op MTBF-gebaseerde componenten."
    )

    if "overzicht_df" not in st.session_state or st.session_state.overzicht_df.empty:
        st.warning("Laad eerst het overzicht via het tabblad 📊 Overzicht.")
    else:
        _df_ov = st.session_state.overzicht_df.copy().reset_index()

        _sl_d = st.selectbox(
            "Service level",
            options=SERVICE_LEVELS,
            index=SERVICE_LEVELS.index(0.990) if 0.990 in SERVICE_LEVELS else 0,
            format_func=lambda v: f"{v:.1%}",
            key="drempel_sl",
        )
        _sl_col = f"s@{_sl_d:.1%}"

        _MAX_N_SEARCH = 100_000
        _drempel_rows = []

        for _, _row in _df_ov.iterrows():
            _code  = _row["Code"]
            _n     = int(_row["n_klanten"])
            _lam   = float(_row["lambda_jr"])
            _lt_jr = float(_row["LT_dagen"]) / 365
            _s_now = int(_row[_sl_col]) if _sl_col in _df_ov.columns else 0

            if _n > 0 and _lam > 0 and _lt_jr > 0:
                _lam_pn = _lam / _n
                # Binary search: kleinste N_drempel waarbij S* > _s_now
                _lo, _hi = _n + 1, _n + _MAX_N_SEARCH
                _s_hi = BPAOptimizationModel.inverse_service_level(
                    _sl_d, _lam_pn * _hi, _lt_jr
                )
                if _s_hi <= _s_now:
                    _n_drempel = None
                else:
                    while _lo < _hi:
                        _mid = (_lo + _hi) // 2
                        _s_mid = BPAOptimizationModel.inverse_service_level(
                            _sl_d, _lam_pn * _mid, _lt_jr
                        )
                        if _s_mid > _s_now:
                            _hi = _mid
                        else:
                            _lo = _mid + 1
                    _n_drempel = _lo
            else:
                _n_drempel = None

            _extra = (_n_drempel - _n) if _n_drempel is not None else None
            _drempel_rows.append({
                "Code":          _code,
                "Omschrijving":  str(_row.get("Descr", ""))[:35],
                "Z huidig":      _n,
                "S* huidig":     _s_now,
                "Z voor S*+1":   _n_drempel if _n_drempel is not None else f">{_n + _MAX_N_SEARCH}",
                "Extra Z nodig": _extra,
                "λ/jr":          round(_lam, 4),
                "μ = λ·L":       round(float(_row["mu"]), 4),
            })

        _tbl_d = pd.DataFrame(_drempel_rows).set_index("Code")
        _tbl_d_sorted = _tbl_d.sort_values("Extra Z nodig", na_position="last")
        # Styler.apply werkt niet met een niet-unieke index (dubbele 'Code').
        # Reset naar een unieke RangeIndex en verberg die in de weergave.
        if not _tbl_d_sorted.index.is_unique:
            _tbl_d_sorted = _tbl_d_sorted.reset_index()

        # Tabel weergeven met kleurcodering op basis van drempel
        def _kleur_drempel(row):
            v = row["Extra Z nodig"]
            if pd.isna(v):
                bg = "#d4edda"   # groen: geen drempel gevonden in zoekbereik
            elif int(v) <= 2:
                bg = "#f8d7da"   # rood: 1-2 extra subscripties
            elif int(v) <= 5:
                bg = "#fff3cd"   # oranje: 3-5 extra subscripties
            else:
                bg = "#d4edda"   # groen: 6+ extra subscripties
            return [f"background-color: {bg}"] * len(row)

        st.dataframe(
            _tbl_d_sorted.style
                .apply(_kleur_drempel, axis=1)
                .format({
                    "Z huidig":      "{:.0f}",
                    "S* huidig":     "{:.0f}",
                    "λ/jr":          "{:.4f}",
                    "μ = λ·L":       "{:.4f}",
                    "Extra Z nodig": lambda v: f"{int(v)}" if pd.notna(v) else "—",
                }),
            use_container_width=True,
            height=500,
        )

        # ── Bar chart: Extra N nodig per component ─────────────────────────
        _plot_d = _tbl_d_sorted[_tbl_d_sorted["Extra Z nodig"].notna()].copy()
        if not _plot_d.empty:
            import matplotlib.pyplot as _plt_d

            # Cap het aantal balken: bij honderden componenten wordt de grafiek
            # onleesbaar én PIL gooit een DecompressionBombError zodra het
            # gerenderde PNG > ~179 megapixels wordt. Toon de top-N met
            # de hoogste drempel (relevante "rode" gevallen eerst).
            _MAX_BARS = 60
            _n_total  = len(_plot_d)
            if _n_total > _MAX_BARS:
                _plot_d = _plot_d.nsmallest(_MAX_BARS, "Extra Z nodig")
                st.caption(
                    f"📊 Grafiek toont de **{_MAX_BARS}** componenten met de "
                    f"laagste drempel (van {_n_total} totaal). Volledige lijst "
                    f"staat in de tabel hierboven."
                )

            # Begrens figuur-breedte (max 32 inch) en zet expliciet dpi=100
            # om gegarandeerd onder de PIL-pixellimiet te blijven.
            _fig_w = min(max(8, len(_plot_d) * 0.55), 32)
            _fig_d, _ax_d = _plt_d.subplots(figsize=(_fig_w, 5), dpi=100)
            _ax_d.bar(
                range(len(_plot_d)),
                _plot_d["Extra Z nodig"].astype(int),
                color="#1976D2",
            )
            _ax_d.set_xticks(range(len(_plot_d)))
            _ax_d.set_xticklabels(
                _plot_d.index, rotation=45, ha="right", fontsize=9
            )
            _ax_d.set_ylabel("Extra subscriptions for S*+1", fontsize=11)
            _ax_d.set_title(
                f"Subscription threshold per component  (SL = {_sl_d:.1%})",
                fontsize=12,
            )
            _ax_d.grid(True, axis="y", alpha=0.3)
            _fig_d.tight_layout()
            st.pyplot(_fig_d)
            _plt_d.close(_fig_d)


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 9 – CLASSIFICATIE
# ─────────────────────────────────────────────────────────────────────────────

with tab_classificatie:
    st.subheader("Classificatie — selectie voor BPA-beheer")
    st.caption(
        "Score alle artikelen uit de bron-Excel op prijs, klantlocaties en order-frequentie. "
        "Pas de gewichten en drempel aan; de selectie wordt — na 'Toepassen' — als whitelist "
        "doorgezet naar het tabblad 📊 Overzicht."
    )

    # ── Bron-Excel (optioneel uploaden, anders EXCEL_PATH uit repo) ──
    _cls_upload = st.file_uploader(
        "Optioneel: upload een andere bron-Excel (anders wordt de repo-Excel gebruikt)",
        type=["xlsx"],
        key="cls_upload",
    )
    _cls_bron = _cls_upload if _cls_upload is not None else EXCEL_PATH
    _cls_sheet = st.text_input(
        "Sheet-naam (leeg = eerste sheet)",
        value="Filtered ",
        key="cls_sheet",
        help=("Tabblad in de Excel waar de classificatie op draait. "
              "Standaard 'Filtered ' (= zelfde sheet als BPA-overzicht) zodat "
              "MTBF(years) en de andere metadata correct meegenomen worden. "
              "De eerste sheet 'Final_data' bevat slechts een placeholder-MTBF "
              "in dagen en geeft daardoor te hoge λ-waarden voor "
              "classificatie-only componenten."),
    ).strip() or None

    st.divider()

    # ── Parameters ──
    st.markdown("**Gewichten** _(worden automatisch genormaliseerd)_")
    _c1, _c2, _c3 = st.columns(3)
    with _c1:
        _w_prijs = st.slider("Gewicht prijs",    0.0, 1.0, 1/3, 0.05, key="cls_w_prijs")
    with _c2:
        _w_loc   = st.slider("Gewicht locaties", 0.0, 1.0, 1/3, 0.05, key="cls_w_loc")
    with _c3:
        _w_ord   = st.slider("Gewicht orders",   0.0, 1.0, 1/3, 0.05, key="cls_w_ord")

    st.markdown("**Selectiemethode**")
    _sel_modus = st.radio(
        "Hoe wordt de selectie bepaald?",
        options=["top_n", "threshold", "top_pct_all"],
        format_func=lambda m: (
            "Top X componenten (hoogste gewogen score)"
            if m == "top_n"
            else "Drempelwaarde (gewogen score ≥ drempel)"
            if m == "threshold"
            else "Top X% per criterium (φ én χ én ψ)"
        ),
        horizontal=True,
        key="cls_sel_modus",
        help=("Top X: neem de X componenten met de hoogste gewogen score op. "
              "Drempelwaarde: neem álle componenten met een score ≥ drempel op. "
              "Top X% per criterium: neem alleen componenten op die tegelijk in "
              "de bovenste X% zitten voor prijs (φ), locaties (χ én orders (ψ)."),
    )
    _sm1, _sm2 = st.columns(2)
    with _sm1:
        if _sel_modus == "top_n":
            _top_n = st.number_input(
                "Aantal componenten (top X)", 1, 100_000, 100, 1, key="cls_top_n",
                help="De X componenten met de hoogste gewogen score (ná de harde "
                     "filters) worden opgenomen in de lijst.",
            )
            _thr = 0.0
            _top_pct = 20.0
        elif _sel_modus == "top_pct_all":
            _top_pct = st.slider(
                "Top X% per criterium", 1.0, 50.0, 20.0, 1.0, key="cls_top_pct",
                help="Een component wordt opgenomen als het tegelijk in de bovenste "
                     "X% valt voor φ (prijs), χ (locaties) én ψ (orders per locatie).",
            )
            _thr = 0.0
            _top_n = 100
        else:
            _thr = st.number_input("Drempel (≥ opnemen)", 0.0, 100.0, 55.0, 1.0, key="cls_thr")
            _top_n = 100
            _top_pct = 20.0

    st.markdown("**Niet-lineariteiten**")
    _ord_pow = st.slider("Orders-power", 1.0, 4.0, 2.0, 0.1, key="cls_ord_pow")

    st.markdown("**Min-filter drempels**")
    st.caption(
        "Artikelen onder deze drempels worden uitgesloten vóór de weging wordt toegepast."
    )
    _mf1, _mf2 = st.columns(2)
    with _mf1:
        _min_prijs = st.number_input(
            "Min. verkoopprijs (€)", 0.0, 100_000.0, 0.0, 10.0,
            key="cls_min_prijs",
            help="Artikelen met verkoopprijs < dit bedrag worden uitgesloten (harde filter).",
        )
    with _mf2:
        _min_orders = st.number_input(
            "Min. gem. orders/locatie", 0.0, 100.0, 0.0, 0.1,
            key="cls_min_orders",
            format="%.1f",
            help="Artikelen met gem. orders/locatie < deze waarde worden uitgesloten.",
        )

    st.markdown("**Aggregatiemethode**")
    _score_methode = st.radio(
        "Score-aggregatie",
        options=["arithmetisch", "geometrisch"],
        index=0,
        horizontal=True,
        key="cls_score_methode",
        help=(
            "Arithmetisch: gewogen som (lineair compensatoir). "
            "Geometrisch: gewogen geometrisch gemiddelde — een zeer lage score op "
            "één dimensie trekt de totaalscore sterker omlaag."
        ),
    )
    if _score_methode == "geometrisch":
        _epsilon = st.number_input(
            "ε (epsilon, verschuiving)", 0.001, 10.0, 1.0, 0.1,
            key="cls_epsilon",
            format="%.3f",
            help="Kleine constante waarmee elke score wordt verschoven voor de machtsverheffing "
                 "(s̅ᴵ = (s + ε) / (100 + ε)). Standaard 1.0."
        )
    else:
        _epsilon = 1.0

    st.markdown("**Harde filters**")
    _c8, _c9 = st.columns(2)
    with _c8:
        _min_loc = st.number_input("Min. klantlocaties", 0, 100, 5, 1, key="cls_min_loc")
    with _c9:
        _art_types_raw = st.text_input(
            "ArticleType-filter (komma-gescheiden, case-insensitief)",
            value="critical, onbekend",
            key="cls_art_types",
        )
    _art_types = tuple(s.strip().lower() for s in _art_types_raw.split(",") if s.strip())

    _params = ClassificatieParams(
        threshold=float(_thr),
        selectie_modus=_sel_modus,
        top_n=int(_top_n),
        top_pct=float(_top_pct),
        weight_prijs=float(_w_prijs),
        weight_locaties=float(_w_loc),
        weight_orders=float(_w_ord),
        orders_power=float(_ord_pow),
        min_prijs=float(_min_prijs),
        min_orders=float(_min_orders),
        min_klantlocaties=int(_min_loc),
        article_type_filter=_art_types,
        score_methode=_score_methode,
        epsilon=float(_epsilon),
    )

    st.divider()

    # ── Run-knop ──
    _col_run, _col_apply = st.columns([1, 1])
    with _col_run:
        _run_cls = st.button("🔄 Bereken classificatie", type="primary", key="cls_run")
    with _col_apply:
        _apply_cls = st.button("✅ Toepassen op BPA-overzicht", key="cls_apply",
                               disabled=("cls_result" not in st.session_state))

    if _run_cls:
        try:
            with st.spinner("Classificatie berekenen…"):
                # De (trage) Excel-parse wordt gecachet, zodat alleen de
                # gevectoriseerde scoring opnieuw draait bij parameter-tweaks.
                if _cls_upload is not None:
                    _df_raw = _cached_laad_ruwe_dataset(0.0, _cls_sheet, _cls_upload)
                    _bron_excel = None
                else:
                    _df_raw = _cached_laad_ruwe_dataset(
                        _file_mtime(EXCEL_PATH), _cls_sheet
                    )
                    _bron_excel = str(EXCEL_PATH)
                _miss = controleer_kolommen(_df_raw)
                if _miss:
                    raise ValueError(f"Ontbrekende kolommen: {_miss}")
                # Eerst basis-filteren, daarna scoren: de min-max-normalisatie
                # gaat zo over de artikelenset NÁ de harde filters. Top-n volgt
                # op de gescoorde set.
                _df_basis    = pas_basis_filters_toe(_df_raw, _params)
                _df_scored   = bereken_scores(_df_basis, _params)
                _df_filtered = pas_topn_selectie_toe(_df_scored, _params)
                _payload     = bouw_selectie_payload(
                    _df_filtered, _params, bron_excel=_bron_excel
                )
            st.session_state.cls_result   = _df_filtered
            st.session_state.cls_payload  = _payload
            st.session_state.cls_params   = _params
            st.session_state.cls_raw      = _df_raw
            _sel_info = (
                f"top {_params.top_n}" if _params.selectie_modus == "top_n"
                else f"top {_params.top_pct:.0f}% per criterium" if _params.selectie_modus == "top_pct_all"
                else f"drempel ≥ {_params.threshold}"
            )
            st.toast(f"{_payload['n_items']} componenten geselecteerd "
                     f"({_sel_info})", icon="✅")
        except Exception as e:
            st.error(f"Fout tijdens classificatie: {e}")

    # ── Resultaten ──
    if "cls_result" in st.session_state:
        _res = st.session_state.cls_result
        _pl  = st.session_state.cls_payload

        _n_tot     = len(_res)
        _n_opnemen = (_res["Classificatie_Beslissing"] == "Opnemen in lijst").sum()

        _m1, _m2, _m3, _m4 = st.columns(4)
        _m1.metric("Na harde filters", _n_tot)
        _m2.metric("Opnemen in lijst", int(_n_opnemen),
                   delta=f"{_n_opnemen/_n_tot*100:.0f}%" if _n_tot else "—")
        _m3.metric("LT geupdate",  _pl["lt_overzicht"]["geupdate"])
        _m4.metric("LT default / ontbreekt",
                   _pl["lt_overzicht"]["default"] + _pl["lt_overzicht"]["ontbreekt"])

        # Tabel — sorteer op score, kleurcodering op beslissing
        _show_cols = [c for c in [
            "Verkooporderregel artikel.Artikel.Artikelcode", "Artikelcode", "Code",
            "ABC_categorie", "ArticleType",
            "Standaard verkoopprijs",
            "Aantal_klantlocaties_met_orders_5jr",
            "Gem_orders_per_klantlocatie_5jr",
            # MTBF: bron-kolom (originele waarde + eenheid) + genormaliseerd in jaren
            "MTBF(years)", "MTBF (years)", "MTBF_years",
            "MTBF(jaren)", "MTBF (jaren)",
            "MTBF (dagen)", "MTBF(dagen)",
            "MTBF (days)", "MTBF(days)", "MTBF_days",
            "MTBF",
            "MTBF_jaren",
            "Lambda_jr",
            "Score_Prijs", "Score_Locaties", "Score_Orders",
            "Gewogen_Score", "Classificatie_Beslissing",
            "Hoofdleverancier.Levertijd",
        ] if c in _res.columns]
        _df_show = _res[_show_cols].sort_values("Gewogen_Score", ascending=False)

        def _kleur_beslissing(v):
            return ("background-color: #c8e6c9" if v == "Opnemen in lijst"
                    else "background-color: #ffcdd2")

        st.dataframe(
            _df_show.style.map(_kleur_beslissing, subset=["Classificatie_Beslissing"]),
            use_container_width=True, height=500,
        )

        # Download
        _csv = _df_show.to_csv(sep=";", decimal=",", index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download gescoorde tabel (CSV)",
            data=_csv, file_name=f"classificatie_{date.today()}.csv",
            mime="text/csv",
        )

        # ── Drempel-sweep analyse ──
        if "cls_raw" in st.session_state:
            with st.expander("📉 Drempel-sweep: stabiliteit van de selectielijst", expanded=False):
                st.caption(
                    "Per filter wordt één drempel gevarieerd; de andere twee blijven op hun huidige waarde. "
                    "Jaccard-similariteit en Kendall\u2019s \u03c4 meten hoeveel de selectielijst "
                    "afwijkt t.o.v. de huidige instelling (referentie = rode stippellijn)."
                )
                _sw_raw    = st.session_state.cls_raw
                _sw_params = st.session_state.cls_params
                _sw_types  = set(s.lower() for s in _sw_params.article_type_filter)
                _sw_mask   = (
                    _sw_raw["ArticleType"].astype(str).str.strip().str.lower()
                    .isin(_sw_types)
                )
                _sw_scored = bereken_scores(_sw_raw[_sw_mask].copy(), _sw_params)

                _SW_LOC = "Aantal_klantlocaties_met_orders_5jr"
                _SW_PRI = "Standaard verkoopprijs"
                _SW_ORD = "Gem_orders_per_klantlocatie_5jr"

                def _sw_sel(ml, mp, mo):
                    """Hard filters + topn; returns frozenset of index labels."""
                    _t = _sw_scored.copy()
                    if ml > 0 and _SW_LOC in _t.columns:
                        _t = _t[_t[_SW_LOC].fillna(0) >= ml]
                    if mp > 0 and _SW_PRI in _t.columns:
                        _t = _t[_t[_SW_PRI].fillna(0) >= mp]
                    if mo > 0 and _SW_ORD in _t.columns:
                        _t = _t[_t[_SW_ORD].fillna(0) >= mo]
                    return frozenset(pas_topn_selectie_toe(_t, _sw_params).index)

                def _sweep_stats(col_dim, thr_vals, base_loc, base_pri, base_ord):
                    """Returns list of (thr, jaccard, kendall_tau) vs baseline."""
                    try:
                        from scipy.stats import kendalltau as _kt
                    except ImportError:
                        _kt = None
                    import numpy as _np_sw
                    _idx     = list(_sw_scored.index)
                    _base    = _sw_sel(base_loc, base_pri, base_ord)
                    _v_base  = _np_sw.array([1 if i in _base else 0 for i in _idx], dtype=_np_sw.int8)
                    out = []
                    for tv in thr_vals:
                        ml = tv if col_dim == _SW_LOC else base_loc
                        mp = tv if col_dim == _SW_PRI else base_pri
                        mo = tv if col_dim == _SW_ORD else base_ord
                        _new   = _sw_sel(ml, mp, mo)
                        _v_new = _np_sw.array([1 if i in _new else 0 for i in _idx], dtype=_np_sw.int8)
                        _inter = int((_v_base & _v_new).sum())
                        _union = int((_v_base | _v_new).sum())
                        _jac   = _inter / _union if _union > 0 else 1.0
                        if _kt is not None:
                            try:
                                _tau = float(_kt(_v_base, _v_new)[0])
                            except Exception:
                                _tau = float("nan")
                        else:
                            _tau = float("nan")
                        out.append((tv, _jac, _tau))
                    return out

                import matplotlib.pyplot as _plt_sw
                _sw_configs = [
                    (_SW_LOC, "Min. klantlocaties",      float(_sw_params.min_klantlocaties)),
                    (_SW_PRI, "Min. verkoopprijs (€)",   float(_sw_params.min_prijs)),
                    (_SW_ORD, "Min. orders/locatie",     float(_sw_params.min_orders)),
                ]
                _sw_fig, _sw_axes = _plt_sw.subplots(2, 3, figsize=(13, 5.5), sharex="col")
                for _ci, (_sw_col, _sw_lbl, _sw_cur) in enumerate(_sw_configs):
                    _ax_j = _sw_axes[0][_ci]
                    _ax_t = _sw_axes[1][_ci]
                    if _sw_col not in _sw_scored.columns:
                        for _ax in (_ax_j, _ax_t):
                            _ax.text(0.5, 0.5, "kolom niet gevonden", ha="center",
                                     va="center", transform=_ax.transAxes, fontsize=9)
                        continue
                    _sw_vals  = _sw_scored[_sw_col].fillna(0)
                    _sw_p95   = float(_sw_vals.quantile(0.95))
                    _sw_max   = max(_sw_p95, _sw_cur * 1.5, 1.0)
                    _sw_steps = [_sw_max * k / 39 for k in range(40)]
                    _sw_res   = _sweep_stats(
                        _sw_col, _sw_steps,
                        float(_sw_params.min_klantlocaties),
                        float(_sw_params.min_prijs),
                        float(_sw_params.min_orders),
                    )
                    _xs   = [r[0] for r in _sw_res]
                    _jacs = [r[1] for r in _sw_res]
                    _taus = [r[2] for r in _sw_res]
                    _ax_j.plot(_xs, _jacs, color="#1976d2", linewidth=2)
                    _ax_j.set_ylabel("Jaccard", fontsize=9)
                    _ax_j.set_ylim(-0.05, 1.05)
                    _ax_j.set_title(_sw_lbl, fontsize=9)
                    _ax_t.plot(_xs, _taus, color="#388e3c", linewidth=2)
                    _ax_t.set_ylabel("Kendall's \u03c4", fontsize=9)
                    _ax_t.set_ylim(-0.05, 1.05)
                    _ax_t.set_xlabel(_sw_lbl, fontsize=9)
                    if _sw_cur > 0:
                        for _ax in (_ax_j, _ax_t):
                            _ax.axvline(_sw_cur, color="#b71c1c", linestyle="--",
                                        linewidth=1.5, label=f"huidig: {_sw_cur:g}")
                            _ax.legend(fontsize=8)
                    for _ax in (_ax_j, _ax_t):
                        _ax.tick_params(labelsize=8)
                _plt_sw.tight_layout()
                st.pyplot(_sw_fig)
                _plt_sw.close(_sw_fig)

        # ── Verdeling: 3D-visualisatie + bar charts per criterium ──
        # Twee weergaven: (a) genormaliseerde scores (0–100/200) en
        # (b) de daadwerkelijke ruwe componentdata (€, #locaties, #orders).
        _score_cols = ["Score_Prijs", "Score_Locaties", "Score_Orders"]
        _raw_cols   = [
            "Standaard verkoopprijs",
            "Aantal_klantlocaties_met_orders_5jr",
            "Gem_orders_per_klantlocatie_5jr",
        ]
        _has_scores = all(c in _res.columns for c in _score_cols)
        _has_raw    = all(c in _res.columns for c in _raw_cols)

        if _has_scores or _has_raw:
            st.divider()
            st.markdown("### 📐 Verdeling per criterium")
            st.caption(
                "Visualiseer hoe de componenten zich verhouden op de drie criteria. "
                "De 3D-scatter toont de spreiding over álle criteria tegelijk; "
                "de histogrammen tonen per criterium hoe scheef (skewed) de verdeling is."
            )

            import matplotlib.pyplot as _plt_cls
            import matplotlib.ticker as _mt_cls
            from mpl_toolkits.mplot3d import Axes3D  # noqa: F401  (registreert 3d-projectie)

            # Keuze databron: genormaliseerde scores vs ruwe componentdata.
            _bron_opties = []
            if _has_scores:
                _bron_opties.append("Genormaliseerde scores")
            if _has_raw:
                _bron_opties.append("Ruwe componentdata")
            _viz_bron = st.radio(
                "Databron voor visualisatie",
                _bron_opties,
                horizontal=True,
                key="cls_viz_bron",
            )

            if _viz_bron == "Ruwe componentdata":
                _viz_cols  = _raw_cols
                _crit_meta = {
                    _raw_cols[0]: ("Price (€)",      "#1976D2"),
                    _raw_cols[1]: ("Customer locations",  "#388E3C"),
                    _raw_cols[2]: ("Orders / location", "#F57C00"),
                }
                _eenheid_x = "Value"
            else:
                _viz_cols  = _score_cols
                _crit_meta = {
                    "Score_Prijs":    ("Price",    "#1976D2"),
                    "Score_Locaties": ("Locations", "#388E3C"),
                    "Score_Orders":   ("Orders",   "#F57C00"),
                }
                _eenheid_x = "Score"

            # Optionele log-schaal — handig bij sterk scheve ruwe data (prijs).
            _log_schaal = st.checkbox(
                "Log-schaal op assen (handig bij scheve ruwe data)",
                value=(_viz_bron == "Ruwe componentdata"),
                key="cls_viz_log",
            )

            _plot_df = _res.copy()
            for _c in _viz_cols:
                _plot_df[_c] = pd.to_numeric(_plot_df[_c], errors="coerce")
            _plot_df = _plot_df.dropna(subset=_viz_cols)
            if _log_schaal:
                # Log-schaal vereist strikt positieve waarden.
                _plot_df = _plot_df[(_plot_df[_viz_cols] > 0).all(axis=1)]

            if _plot_df.empty:
                st.info("Geen geldige data beschikbaar voor visualisatie "
                        "(controleer evt. de log-schaal-optie).")
            else:
                _cx, _cy, _cz = _viz_cols
                _lx, _ly, _lz = (_crit_meta[_cx][0], _crit_meta[_cy][0], _crit_meta[_cz][0])

                _opnemen_mask = (
                    _plot_df["Classificatie_Beslissing"] == "Opnemen in lijst"
                    if "Classificatie_Beslissing" in _plot_df.columns
                    else pd.Series(True, index=_plot_df.index)
                )

                # ── 3D-scatter: alle drie de criteria tegelijk ──
                _fig3d = _plt_cls.figure(figsize=(10, 8))
                _ax3d = _fig3d.add_subplot(111, projection="3d")

                for _mask, _kleur, _lbl in [
                    (_opnemen_mask,  "#2E7D32", "Include in list"),
                    (~_opnemen_mask, "#C62828", "Do not include"),
                ]:
                    _sub = _plot_df[_mask]
                    if not _sub.empty:
                        _ax3d.scatter(
                            _sub[_cx], _sub[_cy], _sub[_cz],
                            c=_kleur, label=_lbl, s=22, alpha=0.6,
                            edgecolors="none", depthshade=True,
                        )

                _ax3d.set_xlabel(_lx, fontsize=10, labelpad=12)
                _ax3d.set_ylabel(_ly, fontsize=10, labelpad=12)
                _ax3d.set_zlabel(_lz, fontsize=10, labelpad=12)
                if _log_schaal:
                    _ax3d.set_xscale("log")
                    _ax3d.set_yscale("log")
                    _ax3d.set_zscale("log")
                _ax3d.set_title(
                    f"3D distribution — {_viz_bron.lower()} ({len(_plot_df)} components)",
                    fontsize=12,
                )
                _ax3d.legend(fontsize=9, loc="upper left")
                _ax3d.view_init(elev=22, azim=-58)
                # tight_layout() knipt de z-as-label (orderfrequentie) van 3D-plots
                # weg. Zoom de 3D-box iets uit en gebruik ruime, gebalanceerde
                # marges zodat álle drie de assen + labels zichtbaar blijven
                # (de z-as 'Orders' staat bij deze kijkhoek aan de linkerkant).
                try:
                    _ax3d.set_box_aspect(None, zoom=0.82)
                except (TypeError, AttributeError):
                    pass  # oudere matplotlib zonder zoom-parameter
                _fig3d.subplots_adjust(left=0.06, right=0.96, top=0.95, bottom=0.06)
                st.pyplot(_fig3d)
                _plt_cls.close(_fig3d)

                # ── Bar charts (histogrammen) per criterium ──
                st.markdown("**Verdeling per criterium**")
                _hist_cols = st.columns(3)
                for _col_name, _slot in zip(_viz_cols, _hist_cols):
                    _vals = _plot_df[_col_name].astype(float)
                    _lbl, _kleur = _crit_meta[_col_name]
                    _figh, _axh = _plt_cls.subplots(figsize=(4.5, 3.2))
                    if _log_schaal and (_vals > 0).all():
                        _bins = np.logspace(
                            np.log10(_vals.min()), np.log10(_vals.max()), 20
                        )
                        _axh.set_xscale("log")
                    else:
                        _bins = 20
                    _axh.hist(
                        _vals, bins=_bins, color=_kleur,
                        edgecolor="white", alpha=0.85,
                    )
                    _mediaan = float(_vals.median())
                    _gem = float(_vals.mean())
                    _axh.axvline(_gem, color="#212121", linestyle="--",
                                 linewidth=1.2, label=f"Mean {_gem:,.1f}")
                    _axh.axvline(_mediaan, color="#757575", linestyle=":",
                                 linewidth=1.2, label=f"Median {_mediaan:,.1f}")
                    _axh.set_title(_lbl, fontsize=11)
                    _axh.set_xlabel(_eenheid_x, fontsize=9)
                    _axh.set_ylabel("Number of components", fontsize=9)
                    _axh.legend(fontsize=8)
                    _axh.grid(True, axis="y", alpha=0.3)
                    _figh.tight_layout()
                    with _slot:
                        st.pyplot(_figh)
                    _plt_cls.close(_figh)

                # ── Scheefheid (skewness) per criterium ──
                _skew_data = {
                    _crit_meta[c][0]: [
                        round(float(_plot_df[c].mean()), 2),
                        round(float(_plot_df[c].median()), 2),
                        round(float(_plot_df[c].skew()), 2),
                    ]
                    for c in _viz_cols
                }
                _skew_df = pd.DataFrame(
                    _skew_data, index=["Gemiddelde", "Mediaan", "Scheefheid"]
                ).T
                st.markdown("**Scheefheid (skewness) per criterium**")
                st.caption(
                    "Scheefheid > 0 = rechts-scheef (veel lage waarden, enkele uitschieters); "
                    "< 0 = links-scheef; ≈ 0 = symmetrisch."
                )
                st.dataframe(_skew_df, use_container_width=True)

    # ── Gewichten-sensitivity sweep ──────────────────────────────────────
    if "cls_result" in st.session_state:
        st.divider()
        st.markdown("### ⚖️ Gewichten-sensitivity")
        st.caption(
            "Varieer de drie criteria-gewichten (prijs / locaties / orders) over "
            "een simplex-raster en zie hoe stabiel de selectie is. Alle overige "
            "parameters (drempel/top-N, penalty's, harde filters) blijven gelijk. "
            "De huidige gewichten vormen de *baseline*."
        )

        _sw1, _sw2 = st.columns([1, 1])
        with _sw1:
            _sweep_step = st.select_slider(
                "Rasterresolutie (stap)",
                options=[0.5, 0.25, 0.2, 0.1, 0.05],
                value=0.1,
                key="cls_sweep_step",
                help="Kleiner = fijner raster en meer combinaties (langzamer). "
                     "0.1 ≈ 66 combinaties, 0.05 ≈ 231.",
            )
        with _sw2:
            _run_sweep = st.button("⚖️ Bereken gewichten-sweep", key="cls_run_sweep")

        if _run_sweep:
            st.session_state.cls_sweep_on = True

        if st.session_state.get("cls_sweep_on"):
            try:
                _params_json = json.dumps({
                    "threshold":               _params.threshold,
                    "selectie_modus":          _params.selectie_modus,
                    "top_n":                   _params.top_n,
                    "weight_prijs":            _params.weight_prijs,
                    "weight_locaties":         _params.weight_locaties,
                    "weight_orders":           _params.weight_orders,
                    "orders_power":            _params.orders_power,
                    "min_prijs":               _params.min_prijs,
                    "min_orders":              _params.min_orders,
                    "min_klantlocaties":       _params.min_klantlocaties,
                    "article_type_filter":     list(_params.article_type_filter),
                    "score_methode":           _params.score_methode,
                    "epsilon":                 _params.epsilon,
                }, sort_keys=True)
                _per_artikel, _per_combo = _cached_weight_sweep(
                    st.session_state.cls_result, _params_json, float(_sweep_step),
                    versie=2,
                )
            except Exception as e:
                st.error(f"Fout tijdens gewichten-sweep: {e}")
                _per_artikel = _per_combo = None

            if _per_artikel is not None and not _per_artikel.empty:
                _n_combos = len(_per_combo)
                _altijd = int((_per_artikel["Stabiliteit"] == "altijd").sum())
                _soms   = int((_per_artikel["Stabiliteit"] == "soms").sum())
                _nooit  = int((_per_artikel["Stabiliteit"] == "nooit").sum())
                _base_n = int(_per_artikel["In_baseline"].sum())

                _sm1, _sm2, _sm3, _sm4, _sm5 = st.columns(5)
                _sm1.metric("Combinaties", _n_combos)
                _sm2.metric("In baseline", _base_n)
                _sm3.metric("Altijd geselecteerd", _altijd)
                _sm4.metric("Soms (gevoelig)", _soms)
                _sm5.metric("Gem. selectiegrootte",
                            f"{_per_combo['n_opnemen'].mean():.0f}")

                import matplotlib.pyplot as _plt_sw

                _cc1, _cc2 = st.columns(2)

                # (a) Simplex-scatter: kleur = aantal opnemen per combinatie
                with _cc1:
                    st.markdown("**Selectiegrootte per gewicht-combinatie**")
                    _fig1, _ax1 = _plt_sw.subplots(figsize=(5, 4))
                    _sc = _ax1.scatter(
                        _per_combo["weight_prijs"], _per_combo["weight_orders"],
                        c=_per_combo["n_opnemen"], cmap="viridis", s=90,
                        edgecolor="white", linewidth=0.5,
                    )
                    _ax1.set_xlabel("gewicht prijs")
                    _ax1.set_ylabel("gewicht orders")
                    _ax1.set_title("# opnemen (locaties = rest)", fontsize=10)
                    _ax1.grid(True, alpha=0.3)
                    _fig1.colorbar(_sc, ax=_ax1, label="# opnemen")
                    _fig1.tight_layout()
                    st.pyplot(_fig1)
                    _plt_sw.close(_fig1)
                    st.caption("gewicht locaties = 1 − prijs − orders.")

                # (b) Stabiliteitsverdeling
                with _cc2:
                    st.markdown("**Stabiliteit van artikelen**")
                    _fig2, _ax2 = _plt_sw.subplots(figsize=(5, 4))
                    _ax2.bar(
                        ["altijd", "soms", "nooit"],
                        [_altijd, _soms, _nooit],
                        color=["#2ca02c", "#ff7f0e", "#d62728"],
                        edgecolor="white",
                    )
                    for _i, _v in enumerate([_altijd, _soms, _nooit]):
                        _ax2.text(_i, _v, str(_v), ha="center", va="bottom",
                                  fontsize=9)
                    _ax2.set_ylabel("aantal artikelen")
                    _ax2.set_title("'soms' = selectie hangt af van de weging",
                                   fontsize=10)
                    _ax2.grid(True, axis="y", alpha=0.3)
                    _fig2.tight_layout()
                    st.pyplot(_fig2)
                    _plt_sw.close(_fig2)

                # (c) Histogram van selectie-frequentie
                st.markdown("**Verdeling van de selectie-frequentie**")
                _fig3, _ax3 = _plt_sw.subplots(figsize=(9, 2.6))
                _ax3.hist(
                    _per_artikel["Selectie_frequentie"], bins=20,
                    range=(0, 1), color="#1f77b4", edgecolor="white", alpha=0.85,
                )
                _ax3.set_xlabel("fractie combinaties waarin geselecteerd")
                _ax3.set_ylabel("aantal artikelen")
                _ax3.grid(True, axis="y", alpha=0.3)
                _fig3.tight_layout()
                st.pyplot(_fig3)
                _plt_sw.close(_fig3)

                # ── Rangorde-robuustheid ──────────────────────────────────
                if "spearman" in _per_combo.columns:
                    st.markdown("#### 🔢 Effect op de volgorde (rangorde)")
                    st.caption(
                        "Naast wél/niet in de set (Jaccard) telt ook de volgorde. "
                        "Spearman/Kendall = rangcorrelatie over de hele kandidaat-set "
                        "(1 = identieke volgorde). RBO weegt de **top** zwaar — "
                        "relevant voor top-N-prioritering. 'Rank-shift' = aantal "
                        "posities dat een artikel opschuift t.o.v. de baseline."
                    )

                    _cm = _per_combo.copy()
                    _rk1, _rk2, _rk3, _rk4 = st.columns(4)
                    _rk1.metric("Laagste Spearman",
                                f"{_cm['spearman'].min():.2f}",
                                help="Worst-case rangcorrelatie over alle wegingen.")
                    _rk2.metric("Laagste RBO (top-zwaar)",
                                f"{_cm['rbo'].min():.2f}")
                    _rk3.metric("Grootste rank-shift",
                                f"{int(_cm['max_rank_shift'].max())}")
                    _rk4.metric("Gem. rank-shift",
                                f"{_cm['mean_rank_shift'].mean():.1f}")

                    # Per dominant criterium: welk criterium verstoort de
                    # volgorde het meest als het zwaarder weegt?
                    _crit_naam = {0: "prijs", 1: "locaties", 2: "orders"}
                    _wcols = ["weight_prijs", "weight_locaties", "weight_orders"]
                    _cm["_dominant"] = (
                        _cm[_wcols].values.argmax(axis=1)
                    )
                    _cm["Dominant criterium"] = _cm["_dominant"].map(_crit_naam)
                    _grp = (
                        _cm.groupby("Dominant criterium")[
                            ["spearman", "kendall", "rbo", "mean_rank_shift"]
                        ].mean().reindex(["prijs", "locaties", "orders"])
                    )

                    _rc1, _rc2 = st.columns(2)
                    with _rc1:
                        st.markdown("**Rang-stabiliteit per dominant criterium**")
                        _fig4, _ax4 = _plt_sw.subplots(figsize=(5, 4))
                        _xpos = np.arange(len(_grp))
                        _bw = 0.4
                        _ax4.bar(_xpos - _bw/2, _grp["spearman"], _bw,
                                 label="Spearman", color="#1f77b4")
                        _ax4.bar(_xpos + _bw/2, _grp["rbo"], _bw,
                                 label="RBO (top)", color="#ff7f0e")
                        _ax4.set_xticks(_xpos)
                        _ax4.set_xticklabels(_grp.index)
                        _ax4.set_ylim(0, 1)
                        _ax4.set_ylabel("gem. correlatie t.o.v. baseline")
                        _ax4.set_title("Hoger = volgorde blijft stabieler",
                                       fontsize=10)
                        _ax4.legend(fontsize=8)
                        _ax4.grid(True, axis="y", alpha=0.3)
                        _fig4.tight_layout()
                        st.pyplot(_fig4)
                        _plt_sw.close(_fig4)
                        st.caption("Het criterium met de **laagste** balken "
                                   "verstoort de volgorde het sterkst.")

                    with _rc2:
                        st.markdown("**Gem. rank-shift per dominant criterium**")
                        _fig5, _ax5 = _plt_sw.subplots(figsize=(5, 4))
                        _ax5.bar(_grp.index, _grp["mean_rank_shift"],
                                 color=["#2ca02c", "#9467bd", "#d62728"],
                                 edgecolor="white")
                        for _i, _v in enumerate(_grp["mean_rank_shift"]):
                            _ax5.text(_i, _v, f"{_v:.0f}", ha="center",
                                      va="bottom", fontsize=9)
                        _ax5.set_ylabel("gem. positieverschuiving")
                        _ax5.set_title("Hoger = volgorde schuift meer op",
                                       fontsize=10)
                        _ax5.grid(True, axis="y", alpha=0.3)
                        _fig5.tight_layout()
                        st.pyplot(_fig5)
                        _plt_sw.close(_fig5)
                        st.caption("Grotere verschuiving = minder robuuste "
                                   "prioritering.")

                    _worst = _grp["spearman"].idxmin()
                    _best = _grp["spearman"].idxmax()
                    st.info(
                        f"➡️ De volgorde is het **gevoeligst** voor het gewicht van "
                        f"**{_worst}** (laagste rangcorrelatie) en het **robuustst** "
                        f"voor **{_best}**. Onderbouw het gewicht van *{_worst}* het "
                        f"zorgvuldigst; daar bepaalt je keuze de prioritering."
                    )

                # (d) Resultatentabel per artikel
                st.markdown("**Resultaten per artikel** "
                            "_(gesorteerd op selectie-frequentie)_")
                _only_soms = st.checkbox(
                    "Toon alleen gevoelige artikelen (Stabiliteit = 'soms')",
                    value=False, key="cls_sweep_only_soms",
                )
                _tabel = (_per_artikel[_per_artikel["Stabiliteit"] == "soms"]
                          if _only_soms else _per_artikel)
                st.dataframe(
                    _tabel, use_container_width=True, height=420, hide_index=True,
                    column_config={
                        "Selectie_frequentie": st.column_config.ProgressColumn(
                            "Selectie_frequentie", min_value=0.0, max_value=1.0,
                            format="%.2f",
                        ),
                    },
                )

                with st.expander("Detail per gewicht-combinatie"):
                    st.dataframe(_per_combo, use_container_width=True,
                                 hide_index=True)

                _sweep_csv = _per_artikel.to_csv(
                    sep=";", decimal=",", index=False
                ).encode("utf-8")
                st.download_button(
                    "⬇️ Download sweep-resultaten (CSV)",
                    data=_sweep_csv,
                    file_name=f"gewichten_sweep_{date.today()}.csv",
                    mime="text/csv",
                    key="cls_sweep_dl",
                )

    # ── Apply: schrijf bpa_selectie.json + invalideer overzicht ──
    if _apply_cls and "cls_payload" in st.session_state:
        try:
            schrijf_selectie_json(st.session_state.cls_payload, SELECTIE_PATH)
            st.session_state.pop("overzicht_df", None)
            st.success(
                f"✅ Selectie opgeslagen in {SELECTIE_PATH}. "
                f"Open tab 📊 Overzicht — de basisvoorraden worden opnieuw berekend "
                f"met **{st.session_state.cls_payload['n_items']}** componenten als whitelist."
            )
        except Exception as e:
            st.error(f"Kon bpa_selectie.json niet schrijven: {e}")

    # ── Verwijder bestaande selectie (alle artikelen weer actief) ──
    st.divider()
    if os.path.exists(SELECTIE_PATH):
        if st.button("🗑️ Verwijder huidige classificatie-selectie (BPA gebruikt weer alle Excel-codes)"):
            try:
                os.remove(SELECTIE_PATH)
                st.session_state.pop("overzicht_df", None)
                st.toast("Selectie verwijderd — BPA gebruikt weer de standaard Excel-filters.", icon="🗑️")
                st.rerun()
            except Exception as e:
                st.error(f"Kon bestand niet verwijderen: {e}")
    else:
        st.info("Geen actieve classificatie-selectie. De BPA-tool gebruikt momenteel de standaard Excel-filters.")


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 10 – BUDGET-SCENARIO (greedy knapsack)
# ─────────────────────────────────────────────────────────────────────────────

with tab_budget:
    st.subheader("Budget-scenario — greedy selectie")
    st.caption(
        "Stel een maximaal investeringsbudget in en selecteer greedy de componenten "
        "met de hoogste **waarde per euro**. Investering per component = `S* × IP` "
        "bij het gekozen service level. Standaard-criterium = **Winst/jr ÷ investering "
        "(ROI)**, met `Winst/jr = Z·α·VP − κ_BPA·IP·S*` (identiek aan Kostenanalyse-tab)."
    )

    _ov_df = st.session_state.get("overzicht_df")
    if _ov_df is None or _ov_df.empty:
        try:
            _ov_df = get_overzicht_df(cfg)
            st.session_state["overzicht_df"] = _ov_df
        except Exception as e:
            st.error(f"Kon overzicht niet laden: {e}")
            _ov_df = pd.DataFrame()

    if _ov_df is None or _ov_df.empty:
        st.warning("Geen componenten beschikbaar. Voer eerst classificatie uit of laad de Excel.")
    else:
        _sl_cols_b = [c for c in _ov_df.columns if c.startswith("s@")]
        if not _sl_cols_b:
            st.error("Geen service-level kolommen gevonden in het overzicht.")
        else:
            # ── Parameters ──
            _c1, _c2 = st.columns([1, 1])
            with _c1:
                _sl_keuze = st.selectbox(
                    "Service level voor S*",
                    options=_sl_cols_b,
                    index=len(_sl_cols_b) // 2,
                    key="bud_sl",
                )
            _inv_per_comp = (_ov_df[_sl_keuze] * _ov_df["IP"]).round(2)
            _totale_inv   = float(_inv_per_comp.sum())

            with _c2:
                _budget = st.number_input(
                    "Maximaal budget (€)",
                    min_value=0.0,
                    max_value=max(_totale_inv * 1.2, 1_000_000.0),
                    value=float(round(_totale_inv * 0.5, 0)),
                    step=1000.0,
                    key="bud_max",
                )

            # ── Economische parameters (gedeeld met Kostenanalyse-tab) ──
            # Defaults komen uit st.session_state['kosten_params'] indien de
            # Kostenanalyse-tab al een berekening heeft gedaan; anders 15% en 20%.
            _kp = st.session_state.get("kosten_params", {})
            _p1, _p2 = st.columns(2)
            with _p1:
                _alpha_b = st.number_input(
                    "α (abonnementstarief, %)",
                    min_value=0.1, max_value=50.0,
                    value=float(_kp.get("alpha", 0.15)) * 100,
                    step=0.5, format="%.1f",
                    key="bud_alpha",
                    help=("Abonnementsprijs als % van VP. Standaardwaarde komt "
                          "uit de Kostenanalyse-tab; pas hier aan om scenario's "
                          "door te rekenen."),
                ) / 100
            with _p2:
                _kappa_b = st.number_input(
                    "κ_BPA (carrying rate, %)",
                    min_value=0.1, max_value=100.0,
                    value=float(_kp.get("kappa_bpa", 0.20)) * 100,
                    step=0.5, format="%.1f",
                    key="bud_kappa",
                    help=("BPA-voorraadkosten per jaar als % van IP "
                          "(financiering + opslag + obsolescence)."),
                ) / 100

            st.caption(
                f"Volledige voorraadwaarde bij {_sl_keuze}: **€ {_totale_inv:,.0f}** · "
                f"Budget = **€ {_budget:,.0f}** "
                f"({_budget/_totale_inv*100:.0f}% van totaal) · "
                f"α = **{_alpha_b:.1%}** · κ_BPA = **{_kappa_b:.1%}**"
            )

            # ── Economisch model (identiek aan Kostenanalyse-tab) ─────────
            #   Omzet_jr_i = N_i · α · VP_i              (abonnementsstroom)
            #   C_BPA_i    = κ_BPA · IP_i · S*_i         (carrying cost basisvoorraad)
            #   Winst_jr_i = Omzet_jr_i − C_BPA_i
            # Dit is dezelfde formule als calculate_detailed_bpa_costs() en
            # revenue_by_part gebruikt in bouw_model_kosten(). λ speelt hier
            # geen directe rol meer — λ zit impliciet in S* via Poisson.
            _df_b = _ov_df.copy()
            _df_b["S_star"]     = _ov_df[_sl_keuze].astype(float)
            _df_b["Inv"]        = _inv_per_comp
            _df_b["Marge_stuk"] = _df_b["VP"] - _df_b["IP"]
            _df_b["Omzet_jr"]   = _df_b["n_klanten"] * _alpha_b * _df_b["VP"]
            _df_b["C_BPA"]      = _kappa_b * _df_b["IP"] * _df_b["S_star"]
            _df_b["Winst_jr"]   = _df_b["Omzet_jr"] - _df_b["C_BPA"]

            # Waarschuwing voor structureel verliesgevende componenten
            _n_loss = int((_df_b["Winst_jr"] < 0).sum())
            if _n_loss > 0:
                st.warning(
                    f"⚠ {_n_loss} van {len(_df_b)} componenten zijn structureel "
                    f"verliesgevend bij α = {_alpha_b:.1%} en κ_BPA = {_kappa_b:.1%} "
                    f"(C_BPA > abonnementsomzet). Ze blijven zichtbaar in de tabel "
                    f"maar krijgen een negatieve ROI; greedy plaatst ze onderaan en "
                    f"selecteert ze in de regel niet — verhoog eventueel α of "
                    f"verlaag κ_BPA om ze rendabel te maken."
                )

            _waarde_keuze = st.radio(
                "Waarde-criterium",
                options=[
                    "Winst / investering (ROI)",
                    "Classificatie-score",
                    "λ × LT × VP (uitval-impact)",
                    "VP (verkoopprijs)",
                ],
                index=0,
                horizontal=True,
                key="bud_waarde",
                help="Standaard: ROI = jaarlijkse winst per geïnvesteerde euro. "
                     "Greedy selecteert dan de componenten die het hoogste rendement "
                     "op het werkkapitaal opleveren.",
            )

            if _waarde_keuze == "Winst / investering (ROI)":
                # Voor dit criterium is Waarde = jaarwinst, zodat
                # Ratio = Waarde/Inv = Winst_jr/Inv = ROI/jaar.
                _df_b["Waarde"] = _df_b["Winst_jr"]
            elif _waarde_keuze == "Classificatie-score":
                _waarde = pd.to_numeric(_df_b.get("Cls_score"), errors="coerce")
                # Fallback voor rijen zonder cls_score: λ × LT × VP
                _fallback = _df_b["lambda_jr"] * (_df_b["LT_dagen"] / 365) * _df_b["VP"]
                _df_b["Waarde"] = _waarde.fillna(_fallback)
            elif _waarde_keuze == "λ × LT × VP (uitval-impact)":
                _df_b["Waarde"] = _df_b["lambda_jr"] * (_df_b["LT_dagen"] / 365) * _df_b["VP"]
            else:
                _df_b["Waarde"] = _df_b["VP"]

            # Ratio waarde/€ — bescherming tegen IP=0
            _df_b["Ratio"] = np.where(
                _df_b["Inv"] > 0, _df_b["Waarde"] / _df_b["Inv"], np.inf
            )

            # ── Greedy ──
            # Primair op Ratio (waarde/€), afgerond op 3 decimalen zodat items
            # met hetzelfde getoonde ROI (1 dec., in %) als gelijk worden beschouwd;
            # bij gelijke afgeronde ratio kiest de greedy het component met de
            # laagste investering. mergesort houdt de volgorde stabiel.
            _df_sorted = (
                _df_b
                .assign(_Ratio_rd=lambda d: d["Ratio"].round(3))
                .sort_values(["_Ratio_rd", "Inv"], ascending=[False, True], kind="mergesort")
                .drop(columns=["_Ratio_rd"])
                .copy()
            )
            _cum = _df_sorted["Inv"].cumsum()
            _df_sorted["In_selectie"] = _cum <= _budget

            # Probeer optioneel nog kleinere items toe te voegen die nog wel passen
            # (na de eerste die niet meer past — kan totale waarde verhogen).
            # Positie-gebaseerd zodat dubbele Code-indexwaarden geen Series geven.
            _resterend = _budget - _df_sorted.loc[_df_sorted["In_selectie"], "Inv"].sum()
            _sel_arr = _df_sorted["In_selectie"].to_numpy().copy()
            _inv_arr = _df_sorted["Inv"].to_numpy()
            for _pos in np.where(~_sel_arr)[0]:
                _kost = float(_inv_arr[_pos])
                if _kost <= _resterend:
                    _sel_arr[_pos] = True
                    _resterend -= _kost
            _df_sorted["In_selectie"] = _sel_arr

            # ── Afgeleide marge- en ROI-kolommen ──
            # Marge_stuk, Omzet_jr, C_BPA en Winst_jr zijn al in _df_b berekend
            # en blijven geldig na sortering. Hier alleen de afgeleide ratio's.
            _df_sorted["Marge_pct"] = np.where(
                _df_sorted["Omzet_jr"] > 0,
                _df_sorted["Winst_jr"] / _df_sorted["Omzet_jr"] * 100,
                np.nan,
            )
            _df_sorted["ROI_jr"] = np.where(
                _df_sorted["Inv"] > 0,
                _df_sorted["Winst_jr"] / _df_sorted["Inv"] * 100,
                np.inf,
            )

            _in  = _df_sorted[_df_sorted["In_selectie"]]
            _uit = _df_sorted[~_df_sorted["In_selectie"]]

            _inv_gekozen  = float(_in["Inv"].sum())
            _waarde_geko  = float(_in["Waarde"].sum())
            _waarde_tot   = float(_df_sorted["Waarde"].sum())
            _winst_geko   = float(_in["Winst_jr"].sum())
            _winst_tot    = float(_df_sorted["Winst_jr"].sum())
            _omzet_geko   = float(_in["Omzet_jr"].sum())
            _cbpa_geko    = float(_in["C_BPA"].sum())
            _marge_gem    = (_winst_geko / _omzet_geko * 100) if _omzet_geko > 0 else 0.0
            _roi_port     = (_winst_geko / _inv_gekozen * 100) if _inv_gekozen > 0 else 0.0

            # ── Metrics rij 1 ──
            _m1, _m2, _m3, _m4 = st.columns(4)
            _m1.metric("Geselecteerd", f"{len(_in)} / {len(_df_sorted)}")
            _m2.metric("Investering", f"€ {_inv_gekozen:,.0f}",
                       delta=f"-€ {(_totale_inv - _inv_gekozen):,.0f}")
            _m3.metric("Waarde behouden",
                       f"{_waarde_geko/_waarde_tot*100:.1f}%" if _waarde_tot > 0 else "—")
            _m4.metric("Budget-benutting",
                       f"{_inv_gekozen/_budget*100:.1f}%" if _budget > 0 else "—")

            # ── Metrics rij 2: economie (BPA-formule) ──
            _w1, _w2, _w3, _w4 = st.columns(4)
            _w1.metric("Omzet / jaar", f"€ {_omzet_geko:,.0f}",
                       help="Σ Z·α·VP over geselecteerde componenten")
            _w2.metric("C_BPA / jaar", f"€ {_cbpa_geko:,.0f}",
                       help="Σ κ_BPA·IP·S* over geselecteerde componenten")
            _w3.metric("Winst / jaar", f"€ {_winst_geko:+,.0f}",
                       delta=(f"{_winst_geko/_winst_tot*100:.1f}% v/h totaal"
                              if _winst_tot != 0 else None))
            _w4.metric("ROI / jaar (portfolio)", f"{_roi_port:+.1f}%",
                       help="Winst/jaar ÷ totale investering")

            # ── Per-component tabel (kosten-analyse stijl) ──
            st.subheader("Winst & marge per component")
            _rows_b = []
            for _code, _r in _df_sorted.iterrows():
                _rows_b.append({
                    'Code':          str(_code),
                    'Omschrijving':  str(_r.get('Descr', ''))[:35],
                    'Z':             int(_r['n_klanten']),
                    'S*':            int(_r[_sl_keuze]),
                    'IP (€)':        round(float(_r['IP']), 2),
                    'VP (€)':        round(float(_r['VP']), 2),
                    'λ/jr':          round(float(_r['lambda_jr']), 4),
                    'Omzet/jr (€)':  round(float(_r['Omzet_jr']), 2),
                    'C_BPA/jr (€)':  round(float(_r['C_BPA']), 2),
                    'Winst/jr (€)':  round(float(_r['Winst_jr']), 2),
                    'Marge %':       (round(float(_r['Marge_pct']), 1)
                                      if pd.notna(_r['Marge_pct']) else np.nan),
                    'Inv. (€)':      round(float(_r['Inv']), 2),
                    'ROI/jr %':      (round(float(_r['ROI_jr']), 1)
                                      if np.isfinite(_r['ROI_jr']) else np.nan),
                    'Cls_score':     (round(float(_r['Cls_score']), 1)
                                      if pd.notna(_r.get('Cls_score')) else np.nan),
                    'In selectie':   '✓' if _r['In_selectie'] else '✗',
                })
            _tbl_b = pd.DataFrame(_rows_b).set_index('Code')

            def _kleur_sel2(v):
                return ("background-color: #c8e6c9" if v == '✓'
                        else "background-color: #ffcdd2")

            st.dataframe(
                _tbl_b.reset_index().style.format({
                    'IP (€)':         '€ {:,.2f}',
                    'VP (€)':         '€ {:,.2f}',
                    'λ/jr':           '{:.4f}',
                    'Omzet/jr (€)':   '€ {:,.0f}',
                    'C_BPA/jr (€)':   '€ {:,.0f}',
                    'Winst/jr (€)':   '€ {:+,.0f}',
                    'Marge %':        '{:+.1f}%',
                    'Inv. (€)':       '€ {:,.0f}',
                    'ROI/jr %':       '{:+.1f}%',
                    'Cls_score':      '{:.1f}',
                }, na_rep="—").map(_kleur_sel2, subset=['In selectie']),
                use_container_width=True,
                height=420,
            )

            # ── Totaalregels (kosten-stijl) ──
            _in_tbl  = _tbl_b[_tbl_b['In selectie'] == '✓']
            st.markdown(
                f"**Totaal selectie:** S\\* = {int(_in_tbl['S*'].sum())}  |  "
                f"Investering = € {_in_tbl['Inv. (€)'].sum():,.0f}  |  "
                f"Omzet/jr = € {_in_tbl['Omzet/jr (€)'].sum():,.0f}  |  "
                f"C_BPA/jr = € {_in_tbl['C_BPA/jr (€)'].sum():,.0f}  |  "
                f"Winst/jr = € {_in_tbl['Winst/jr (€)'].sum():+,.0f}  |  "
                f"ROI/jr = {_roi_port:+.1f}%"
            )
            st.markdown(
                f"**Totaal portfolio:** S\\* = {int(_tbl_b['S*'].sum())}  |  "
                f"Investering = € {_tbl_b['Inv. (€)'].sum():,.0f}  |  "
                f"Omzet/jr = € {_tbl_b['Omzet/jr (€)'].sum():,.0f}  |  "
                f"C_BPA/jr = € {_tbl_b['C_BPA/jr (€)'].sum():,.0f}  |  "
                f"Winst/jr = € {_tbl_b['Winst/jr (€)'].sum():+,.0f}"
            )

            # ── Greedy-rangschikking (waarde-criterium) ──
            with st.expander("📋 Greedy-rangschikking (waarde-criterium)"):
                _show = _df_sorted[
                    ['Descr', 'LT_dagen', _sl_keuze, 'Inv', 'Waarde',
                     'Ratio', 'Cls_score', 'In_selectie']
                ].copy()
                _show.columns = ['Omschrijving', 'LT(d)', f'S* @ {_sl_keuze}',
                                 'Inv. (€)', 'Waarde', 'Waarde/€',
                                 'Cls_score', 'In selectie']

                def _kleur_sel(v):
                    return ("background-color: #c8e6c9" if v
                            else "background-color: #ffcdd2")

                st.dataframe(
                    _show.reset_index(drop=True).style.format({
                        'Inv. (€)':  '€ {:,.0f}',
                        'Waarde':    '{:,.1f}',
                        'Waarde/€':  '{:.4f}',
                        'Cls_score': '{:.1f}',
                    }, na_rep="—").map(_kleur_sel, subset=['In selectie']),
                    use_container_width=True,
                    height=420,
                )

            # ── Curve: cumulatieve waarde vs budget ──
            with st.expander("📈 Cumulatieve waarde vs. cumulatieve investering"):
                import matplotlib.pyplot as _plt_bud
                _cum_inv  = _df_sorted["Inv"].cumsum().values
                _cum_wrd  = _df_sorted["Waarde"].cumsum().values
                _fig_b, _ax_b = _plt_bud.subplots(figsize=(9, 4.5))
                _ax_b.plot(_cum_inv, _cum_wrd, color="#1976D2", lw=2)
                _ax_b.axvline(_budget, color="#c62828", ls="--",
                              label=f"Budget € {_budget:,.0f}")
                _ax_b.fill_between(_cum_inv, _cum_wrd, where=(_cum_inv <= _budget),
                                   color="#c8e6c9", alpha=0.4, label="In selection")
                _ax_b.set_xlabel("Cumulative investment (€)")
                _ax_b.set_ylabel("Cumulative value")
                _ax_b.set_title("Greedy selection — value build-up with increasing investment")
                _ax_b.grid(True, alpha=0.3)
                _ax_b.legend()
                _fig_b.tight_layout()
                st.pyplot(_fig_b)
                _plt_bud.close(_fig_b)

            # ── Toepassen als uitsluitingen ──
            st.divider()
            st.markdown("**Selectie toepassen op model**")
            st.caption(
                "**Optie 1** — voeg de niet-geselecteerde componenten toe aan de "
                "uitsluitingslijst (handmatige componenten blijven onaangetast). "
                "**Optie 2** — herschrijf `bpa_selectie.json` zodat de BPA-overzicht-"
                "whitelist alleen de budget-geselecteerde componenten bevat. "
                "Optie 2 werkt hetzelfde als de knop in de Classificatie-tab."
            )
            _ba_col1, _ba_col2 = st.columns(2)
            with _ba_col1:
                _btn_excl = st.button(
                    "🚫 Pas toe via uitsluitingen",
                    key="bud_apply_excl",
                    help="Voeg niet-geselecteerde codes toe aan uitgesloten_componenten in de config.",
                )
            with _ba_col2:
                _btn_sel = st.button(
                    "✅ Toepassen op BPA-overzicht",
                    type="primary",
                    key="bud_apply_selectie",
                    help="Herschrijft bpa_selectie.json met alleen de budget-geselecteerde codes.",
                )

            if _btn_excl:
                _uit_codes = [str(c) for c in _uit.index
                              if str(c) not in cfg.get("handmatige_componenten", {})]
                cfg.setdefault("uitgesloten_componenten", [])
                for c in _uit_codes:
                    if c not in cfg["uitgesloten_componenten"]:
                        cfg["uitgesloten_componenten"].append(c)
                sla_config_op(cfg)
                st.session_state.pop("overzicht_df", None)
                st.success(
                    f"{len(_uit_codes)} componenten toegevoegd aan uitsluitingen. "
                    f"Tab 📊 Overzicht toont nu de budget-conforme selectie."
                )
                st.rerun()

            if _btn_sel:
                try:
                    _selected_codes = {str(c) for c in _in.index}
                    if not _selected_codes:
                        st.warning("Geen componenten geselecteerd — selectie niet weggeschreven.")
                    else:
                        # 1) Probeer bestaande selectie te filteren (behoudt
                        #    threshold/parameters/metadata van de classificatie).
                        _existing_raw = None
                        if os.path.exists(SELECTIE_PATH):
                            try:
                                with open(SELECTIE_PATH, encoding="utf-8") as _f:
                                    _existing_raw = json.load(_f)
                            except (json.JSONDecodeError, OSError):
                                _existing_raw = None

                        if _existing_raw and _existing_raw.get("items"):
                            _orig_items = _existing_raw.get("items", [])
                            _new_items = [
                                it for it in _orig_items
                                if str(it.get("code")) in _selected_codes
                            ]
                            _payload_new = dict(_existing_raw)
                        else:
                            # 2) Bouw minimaal payload uit het overzicht
                            #    (geen classificatie aanwezig).
                            _new_items = []
                            for _code in _selected_codes:
                                if _code not in _ov_df.index.astype(str).tolist():
                                    continue
                                _row = _ov_df.loc[_code] if _code in _ov_df.index else None
                                if _row is None:
                                    continue
                                _new_items.append({
                                    "code":              _code,
                                    "score":             float(_row["Cls_score"]) if pd.notna(_row.get("Cls_score")) else 0.0,
                                    "lt_dagen":          int(_row.get("LT_dagen", 30)),
                                    "lt_bron":           str(_row.get("LT_bron", "onbekend")),
                                    "abc":               "",
                                    "descr":             str(_row.get("Descr", ""))[:80],
                                    "ip":                float(_row.get("IP", 0.0)),
                                    "vp":                float(_row.get("VP", 0.0)),
                                    "mtbf":              None,
                                    "totaal_orders_5jr": None,
                                    "n_cust":            int(_row.get("n_klanten", 0)),
                                })
                            _payload_new = {
                                "bron_excel": None,
                                "threshold":  None,
                                "parameters": {},
                            }

                        # Hertel lt_overzicht voor de nieuwe set
                        _lt_ov_new = {"geupdate": 0, "default": 0, "ontbreekt": 0}
                        for _it in _new_items:
                            _b = _it.get("lt_bron", "onbekend")
                            if _b in _lt_ov_new:
                                _lt_ov_new[_b] += 1

                        _payload_new["items"]        = _new_items
                        _payload_new["n_items"]      = len(_new_items)
                        _payload_new["lt_overzicht"] = _lt_ov_new
                        _payload_new["gegenereerd"]  = (
                            f"{pd.Timestamp.today()} (budget-filter)"
                        )

                        schrijf_selectie_json(_payload_new, SELECTIE_PATH)
                        invalidate_caches()
                        st.session_state.pop("overzicht_df", None)
                        st.success(
                            f"✅ {len(_new_items)} componenten opgeslagen in "
                            f"`{SELECTIE_PATH}`. Tab 📊 Overzicht laadt nu alleen "
                            f"de budget-geselecteerde componenten."
                        )
                        st.rerun()
                except Exception as e:
                    st.error(f"Kon selectie niet schrijven: {e}")

            # ── Download ──
            _csv_b = _tbl_b.to_csv(sep=";", decimal=",").encode("utf-8")
            st.download_button(
                "⬇️ Download greedy-selectie (CSV)",
                data=_csv_b,
                file_name=f"budget_scenario_{date.today()}.csv",
                mime="text/csv",
            )

# ─────────────────────────────────────────────────────────────────────────────────
#  TAB 11 – VERWACHTE SUBSCRIPTIES  (E[Z_i(α,X)] uit adoption rate × historie)
# ─────────────────────────────────────────────────────────────────────────────────

with tab_subsim:
    st.subheader("Verwachte subscripties per component")
    st.markdown(
        "Het verwachte aantal subscripties per component $E[Z_i(α)]$ wordt "
        "**analytisch** bepaald uit de subscriptie-dataset (zonder 231-AS: RSPL) "
        "— zonder Monte Carlo. Alleen de componenten uit de "
        "**classificatie-selectie** tellen mee.\n\n"
        "Elk van de $N_i$ historische klanten van een component abonneert "
        "onafhankelijk met de globale **logit-adoptiekans** $q(α)$, dus "
        "$Z_i(α)\\sim\\mathrm{Binomiaal}(N_i,\\,q(α))$ en "
        "$E[Z_i(α)]=N_i\\cdot q(α)$. De adoptiekans volgt een "
        "discrete-keuzemodel op basis van de kostenratio $κ_c/α$:\n\n"
        "$$q(α)=σ\\!\\big(β_0+β_r\\ln(κ_c/α)\\big),\\qquad "
        "β_0=\\operatorname{logit}(q_{eq})=\\ln\\tfrac{q_{eq}}{1-q_{eq}}.$$\n\n"
        "De intercept $β_0$ is geijkt op **kostenpariteit**: bij $α=κ_c$ geldt "
        "$q=q_{eq}$. Het service level $X$ zit in het onwaargenomen nut en "
        "beïnvloedt $q$ niet direct (alleen de voorraad-/kostenkant)."
    )

    _cls_codes = sorted(get_classificatie_info().get("items", {}).keys())
    if _cls_codes:
        st.caption(f"Classificatie-selectie actief: **{len(_cls_codes)}** componenten.")
    else:
        st.warning(
            "Geen classificatie-selectie (`bpa_selectie.json`) gevonden. "
            "Voer eerst de classificatie uit via tab 🏷️ Classificatie."
        )

    # ── Bron-Excel met de subscriptie-dataset ────────────────────────────
    _subsim_upload = st.file_uploader(
        "Bron-Excel met de subscriptie-dataset (leeg = standaard dataset zonder RSPL)",
        type=["xlsx"],
        key="subsim_upload",
    )
    _subsim_default = SUBSCRIPTIES_PATH if os.path.exists(SUBSCRIPTIES_PATH) else None
    _excel_bron = _subsim_upload or _subsim_default
    if _subsim_upload is not None:
        _bron_naam = getattr(_excel_bron, "name", "geüploade Excel")
        st.caption(f"Bron-Excel: **{_bron_naam}**")
    elif _subsim_default is not None:
        st.caption(f"Bron-Excel: subscriptie-dataset (`{os.path.basename(SUBSCRIPTIES_PATH)}`)")
    else:
        st.caption(f"Bron-Excel: repo-Excel (`{os.path.basename(EXCEL_PATH)}`) — upload de subscriptie-dataset voor RSPL-vrije analyse")

    def _excel_arg():
        """Geef een leesbare bron terug; reset de upload-buffer naar het begin."""
        if _excel_bron is None:
            return None
        try:
            _excel_bron.seek(0)
        except (AttributeError, ValueError):
            pass
        return _excel_bron

    # ── Adoptiemodel: globale logit-kans q(α) ─────────────────────────────
    st.markdown("**Adoptiemodel — logit discrete-keuze**")
    _kp_sim        = st.session_state.get("kosten_params", {})
    _alpha_def_sim = float(_kp_sim.get("alpha", 0.15))
    _kappa_c_sim   = float(_kp_sim.get("kappa_c", 0.25))
    _X_def_sim     = float(_kp_sim.get("service_level", 0.99))
    st.caption(
        f"$κ_c$ komt uit de Kostenanalyse: κ_c = **{_kappa_c_sim:.0%}** "
        "_(pas aan via 💰 Kostenanalyse)_. $q_{eq}$ en $β_r$ zijn hieronder "
        "instelbaar en worden meegenomen in de gevoeligheidsanalyse."
    )

    col_a1, col_a2, col_a3 = st.columns(3)
    with col_a1:
        _alpha_sim = st.number_input(
            "α (prijspercentage)", min_value=0.0001, max_value=1.0,
            value=_alpha_def_sim, step=0.01, format="%.2f", key="subsim_alpha",
        )
    with col_a2:
        _q_eq = st.number_input(
            "q_eq (adoptie bij kostenpariteit α=κ_c)", min_value=0.01, max_value=0.99,
            value=0.55, step=0.05, format="%.2f", key="subsim_q_eq",
            help="Kans dat een klant abonneert wanneer het abonnement precies "
                 "even duur is als zelf voorraad houden (α = κ_c).",
        )
    with col_a3:
        _beta_r = st.number_input(
            "β_r (gevoeligheid kostenratio)", min_value=0.0, max_value=20.0,
            value=1.0, step=0.1, format="%.2f", key="subsim_beta_r",
            help="Hoe sterk de adoptie reageert op de kostenratio ln(κ_c/α). "
                 "Groter β_r → adoptie daalt sneller als α richting κ_c stijgt.",
        )

    _q_ad = adoptie_kans(_alpha_sim, _kappa_c_sim, _q_eq, _beta_r)
    _c_q1, _c_q2 = st.columns(2)
    _c_q1.metric("q(α) — adoptiekans", f"{_q_ad:.3f}",
                 delta=f"{_q_ad - _q_eq:+.3f} vs q_eq")
    _c_q2.metric("κ_c / α — kostenratio", f"{_kappa_c_sim / max(_alpha_sim, 1e-9):.2f}")

    # ── Verdeling van Z (binomiale verdeling) ─────────────────────────────
    with st.expander("📊 Verdeling van Z (binomiale verdeling)"):
        st.caption(
            "Achter de verwachtingswaarde $E[Z_i]=N_i\\cdot q(α)$ zit een echte "
            "kansverdeling: $Z_i(α)\\sim\\mathrm{Binomiaal}(N_i, q(α))$. Omdat "
            "alle klanten dezelfde globale adoptiekans $q(α)$ delen, is het "
            "**totaal** $Z_{tot}=\\sum_i Z_i\\sim\\mathrm{Binomiaal}(\\sum_i N_i, q(α))$. "
            "Hieronder de kansverdeling (PMF) bij de huidige α, q_eq en β_r."
        )
        try:
            _n_series = aantal_klanten_per_component(_excel_arg(), _cls_codes)
        except (ValueError, FileNotFoundError, OSError) as _dist_err:
            _n_series = None
            st.info(
                "Verdeling niet beschikbaar: de bron-Excel bevat geen tab "
                f"'Adoptie' of is niet bereikbaar. ({_dist_err})"
            )
        if _n_series is not None and not _n_series.empty:
            import matplotlib.pyplot as _plt_bin
            _N_tot = int(round(float(_n_series.sum())))
            _mu    = _N_tot * _q_ad
            _sd    = (_N_tot * _q_ad * (1.0 - _q_ad)) ** 0.5

            # Totale verdeling Z_tot ~ Binomiaal(Σ N_i, q).
            _cB1, _cB2, _cB3 = st.columns(3)
            _cB1.metric("N totaal (klanten)", f"{_N_tot:,}")
            _cB2.metric("E[Z_tot] = N·q", f"{_mu:,.0f}")
            _cB3.metric("Std. dev. √(N·q·(1−q))", f"{_sd:,.1f}")
            _kt, _pt = binomiale_verdeling(_N_tot, _q_ad)
            _figb, _axb = _plt_bin.subplots(figsize=(10, 4))
            _axb.bar(_kt, _pt, width=1.0, color="#1f77b4", alpha=0.75,
                     edgecolor="none")
            _axb.axvline(_mu, color="#d62728", ls="--", lw=1.5,
                         label=f"E[Z] = {_mu:,.0f}")
            _axb.set_xlabel("total subscriptions  Z_tot")
            _axb.set_ylabel("probability  P(Z_tot = k)")
            _axb.set_title(f"Binomial(N={_N_tot:,}, q={_q_ad:.3f})")
            _axb.legend(fontsize=9)
            _axb.grid(True, alpha=0.3)
            _figb.tight_layout()
            st.pyplot(_figb)
            _plt_bin.close(_figb)

            # Verdeling voor één gekozen component Z_i ~ Binomiaal(N_i, q).
            st.markdown("**Verdeling voor één component**")
            _n_sorted = _n_series.sort_values(ascending=False)
            _codes_sorted = [str(c) for c in _n_sorted.index]
            _sel_code = st.selectbox(
                "Component (Code)", options=_codes_sorted,
                format_func=lambda c: f"{c}  (N={int(_n_series.get(c, 0))})",
                key="subsim_bin_code")
            _Ni   = int(round(float(_n_series.get(_sel_code, 0))))
            _mu_i = _Ni * _q_ad
            _ki, _pi = binomiale_verdeling(_Ni, _q_ad)
            _figc, _axc = _plt_bin.subplots(figsize=(10, 4))
            _axc.bar(_ki, _pi, width=1.0, color="#2ca02c", alpha=0.75,
                     edgecolor="none")
            _axc.axvline(_mu_i, color="#d62728", ls="--", lw=1.5,
                         label=f"E[Z_i] = {_mu_i:,.1f}")
            _axc.set_xlabel(f"subscriptions  Z_i  (component {_sel_code})")
            _axc.set_ylabel("probability  P(Z_i = k)")
            _axc.set_title(f"Binomial(N_i={_Ni}, q={_q_ad:.3f})")
            _axc.legend(fontsize=9)
            _axc.grid(True, alpha=0.3)
            _figc.tight_layout()
            st.pyplot(_figc)
            _plt_bin.close(_figc)

    # ── Automatische doorwerking naar alle tabs ───────────────────────────
    # E[Z_i(α)] = N_i · q(α) wordt analytisch (deterministisch) bepaald en als
    # integer Z-override in de configuratie gezet zodra α, q_eq of β_r wijzigt.
    _auto_z = st.checkbox(
        "Verwachte Z automatisch doorzetten naar alle tabs bij wijziging van α/q_eq/β_r",
        value=True, key="subsim_auto_z",
        help="Schrijft E[Z_i(α)] = N_i·q(α) per component als integer Z-override "
             "en herberekent de overige tabs.",
    )
    if _auto_z and _cls_codes:
        _auto_sig = (
            round(_alpha_sim, 6), round(_q_eq, 6), round(_beta_r, 6),
            round(_kappa_c_sim, 6), tuple(_cls_codes),
        )
        if st.session_state.get("subsim_auto_sig") != _auto_sig:
            try:
                _ez = verwacht_subscripties_per_component(
                    excel_file=_excel_arg(),
                    codes=_cls_codes,
                    alpha=float(_alpha_sim),
                    kappa_c=float(_kappa_c_sim),
                    q_eq=float(_q_eq),
                    beta_r=float(_beta_r),
                )
                if not _ez.empty:
                    _n_ov = dict(cfg.get("n_klanten_overrides", {}))
                    for _code, _val in _ez.items():
                        _n_ov[str(_code)] = max(1, int(round(float(_val))))
                    cfg["n_klanten_overrides"] = _n_ov
                    if "overzicht_df" in st.session_state:
                        st.session_state.overzicht_df_prev = st.session_state.overzicht_df.copy()
                    sla_config_op(cfg)
                    invalidate_caches()
                    st.session_state.pop("overzicht_df", None)
                    st.session_state["subsim_auto_sig"] = _auto_sig
                    st.success(
                        f"✅ Verwachte Z (α={_alpha_sim:.2f}, q(α)={_q_ad:.3f}) "
                        f"automatisch doorgezet naar {len(_ez)} componenten."
                    )
                    st.rerun()
            except (ValueError, FileNotFoundError, OSError) as _auto_err:
                st.info(
                    "Automatische doorwerking overgeslagen: de bron-Excel bevat "
                    f"geen tab 'Adoptie' of is niet bereikbaar. ({_auto_err})"
                )

    # ── Gevoeligheidsanalyse: Σ E[Z] vs. α, β_r en q_eq ───────────────────
    with st.expander("📈 Gevoeligheidsanalyse: verwachte Z vs. α, β_r en q_eq"):
        st.caption(
            "Toont het totaal verwachte aantal subscripties "
            "$\\sum_i E[Z_i]=q(\\cdot)\\cdot\\sum_i N_i$ als functie van het "
            "prijspercentage α, de kostenratio-gevoeligheid β_r en de "
            "pariteitskans q_eq. De stippellijn markeert de huidige instelling."
        )
        _cga, _cgb, _cgc = st.columns(3)
        with _cga:
            _a_min = st.number_input(
                "α-bereik min", min_value=0.0001, max_value=1.0, value=0.02,
                step=0.01, format="%.2f", key="subsim_sens_amin")
            _a_max = st.number_input(
                "α-bereik max", min_value=0.01, max_value=1.0, value=0.40,
                step=0.01, format="%.2f", key="subsim_sens_amax")
        with _cgb:
            _br_min = st.number_input(
                "β_r-bereik min", min_value=0.0, max_value=20.0, value=0.0,
                step=0.1, format="%.2f", key="subsim_sens_brmin")
            _br_max = st.number_input(
                "β_r-bereik max", min_value=0.1, max_value=20.0, value=4.0,
                step=0.1, format="%.2f", key="subsim_sens_brmax")
        with _cgc:
            _qe_min = st.number_input(
                "q_eq-bereik min", min_value=0.01, max_value=0.99, value=0.10,
                step=0.05, format="%.2f", key="subsim_sens_qemin")
            _qe_max = st.number_input(
                "q_eq-bereik max", min_value=0.02, max_value=0.99, value=0.90,
                step=0.05, format="%.2f", key="subsim_sens_qemax")
        _n_grid = st.slider(
            "Aantal gridpunten", min_value=5, max_value=60, value=25,
            key="subsim_sens_n")

        if st.button("Bereken gevoeligheid", key="subsim_sens_btn",
                     disabled=not _cls_codes):
            try:
                _a_grid  = list(np.linspace(float(_a_min),  float(_a_max),  int(_n_grid)))
                _br_grid = list(np.linspace(float(_br_min), float(_br_max), int(_n_grid)))
                _qe_grid = list(np.linspace(float(_qe_min), float(_qe_max), int(_n_grid)))
                with st.spinner("Gevoeligheid berekenen…"):
                    _z_vs_a = gevoeligheid_verwachte_z(
                        _a_grid, 'alpha', float(_alpha_sim), float(_kappa_c_sim),
                        float(_q_eq), float(_beta_r),
                        excel_file=_excel_arg(), codes=_cls_codes)
                    _z_vs_br = gevoeligheid_verwachte_z(
                        _br_grid, 'beta_r', float(_alpha_sim), float(_kappa_c_sim),
                        float(_q_eq), float(_beta_r),
                        excel_file=_excel_arg(), codes=_cls_codes)
                    _z_vs_qe = gevoeligheid_verwachte_z(
                        _qe_grid, 'q_eq', float(_alpha_sim), float(_kappa_c_sim),
                        float(_q_eq), float(_beta_r),
                        excel_file=_excel_arg(), codes=_cls_codes)
                st.session_state["subsim_sens_data"] = {
                    "a_grid": _a_grid, "z_vs_a": _z_vs_a,
                    "br_grid": _br_grid, "z_vs_br": _z_vs_br,
                    "qe_grid": _qe_grid, "z_vs_qe": _z_vs_qe,
                    "alpha": float(_alpha_sim), "beta_r": float(_beta_r),
                    "q_eq": float(_q_eq),
                }
            except (ValueError, FileNotFoundError, OSError) as _sens_err:
                st.warning(
                    "Gevoeligheid niet beschikbaar: de bron-Excel bevat geen "
                    f"tab 'Adoptie' of is niet bereikbaar. ({_sens_err})"
                )
                st.session_state.pop("subsim_sens_data", None)

        _sens = st.session_state.get("subsim_sens_data")
        if _sens:
            import matplotlib.pyplot as _plt_sens
            _figs, (_axa, _axb, _axc) = _plt_sens.subplots(1, 3, figsize=(14, 4))
            _axa.plot(_sens["a_grid"], _sens["z_vs_a"], color="#1f77b4", lw=2)
            _axa.axvline(_sens["alpha"], color="grey", ls="--", lw=1)
            _axa.set_xlabel("price percentage α")
            _axa.set_ylabel("expected total subscriptions  Σ E[Z]")
            _axa.set_title(f"Z vs. α  (β_r={_sens['beta_r']:.2f}, q_eq={_sens['q_eq']:.2f})")
            _axa.grid(True, alpha=0.3)
            _axb.plot(_sens["br_grid"], _sens["z_vs_br"], color="#ff7f0e", lw=2)
            _axb.axvline(_sens["beta_r"], color="grey", ls="--", lw=1)
            _axb.set_xlabel("cost-ratio sensitivity β_r")
            _axb.set_ylabel("Σ E[Z]")
            _axb.set_title(f"Z vs. β_r  (α={_sens['alpha']:.2f})")
            _axb.grid(True, alpha=0.3)
            _axc.plot(_sens["qe_grid"], _sens["z_vs_qe"], color="#2ca02c", lw=2)
            _axc.axvline(_sens["q_eq"], color="grey", ls="--", lw=1)
            _axc.set_xlabel("adoption at parity q_eq")
            _axc.set_ylabel("Σ E[Z]")
            _axc.set_title(f"Z vs. q_eq  (α={_sens['alpha']:.2f})")
            _axc.grid(True, alpha=0.3)
            _figs.tight_layout()
            st.pyplot(_figs)
            _plt_sens.close(_figs)

    # ── Pareto-efficiëntie: BPA-winst vs. klantsurplus over (α, X) ─────────
    with st.expander("🧭 Pareto-efficiëntie: BPA-winst vs. klantsurplus over (α, X)"):
        st.caption(
            "Elk punt is een combinatie van prijspercentage α en service level X. "
            "Via de keten $α\\to q(α)\\to E[Z_i]\\to$ kostenmodel worden twee "
            "concurrerende doelen berekend: de **BPA-marge** (€) en het totale "
            "**klantsurplus**. De adoptie q(α) hangt alleen van α af; X werkt via "
            "de voorraad-/kostenkant. De Pareto-frontier verbindt de "
            "niet-gedomineerde combinaties."
        )
        _kp_par = st.session_state.get("kosten_params", {})
        _kappa_bpa_par = float(_kp_par.get("kappa_bpa", 0.20))
        _kappa_c_par   = float(_kp_par.get("kappa_c", 0.25))
        st.caption(
            f"Kostenparameters uit Kostenanalyse: κ_BPA = **{_kappa_bpa_par:.0%}**, "
            f"κ_c = **{_kappa_c_par:.0%}** _(pas aan via 💰 Kostenanalyse)_."
        )

        _cp1, _cp2 = st.columns(2)
        with _cp1:
            _pa_min = st.number_input(
                "α-bereik min", min_value=0.0001, max_value=1.0, value=0.02,
                step=0.01, format="%.2f", key="subsim_par_amin")
            _pa_max = st.number_input(
                "α-bereik max", min_value=0.01, max_value=1.0, value=0.30,
                step=0.01, format="%.2f", key="subsim_par_amax")
            _pa_n = st.slider("Aantal α-waarden", 2, 12, 7, key="subsim_par_na")
        with _cp2:
            _px_min = st.number_input(
                "X-bereik min", min_value=0.50, max_value=0.9999, value=0.95,
                step=0.005, format="%.3f", key="subsim_par_xmin")
            _px_max = st.number_input(
                "X-bereik max", min_value=0.50, max_value=0.9999, value=0.999,
                step=0.005, format="%.3f", key="subsim_par_xmax")
            _px_n = st.slider("Aantal X-waarden", 2, 12, 5, key="subsim_par_nx")

        if st.button("Bereken Pareto-frontier", key="subsim_par_btn",
                     disabled=not _cls_codes):
            try:
                _ov_par = get_overzicht_df(cfg)
                if _ov_par is None or _ov_par.empty:
                    st.warning("Geen overzicht beschikbaar — laad eerst het overzicht (tab 📊).")
                else:
                    _a_vals = list(np.linspace(float(_pa_min), float(_pa_max), int(_pa_n)))
                    _x_vals = list(np.linspace(float(_px_min), float(_px_max), int(_px_n)))
                    with st.spinner("Pareto-frontier berekenen…"):
                        _par_df = pareto_alpha_X(
                            _ov_par, _a_vals, _x_vals,
                            float(_q_eq), float(_beta_r),
                            _kappa_bpa_par, _kappa_c_par,
                            excel_file=_excel_arg(), codes=_cls_codes)
                    if _par_df.empty:
                        st.warning("Geen resultaten — controleer de Adoptie-tab en selectie.")
                        st.session_state.pop("subsim_par_data", None)
                    else:
                        st.session_state["subsim_par_data"] = _par_df
            except (ValueError, FileNotFoundError, OSError) as _par_err:
                st.warning(
                    "Pareto-analyse niet beschikbaar: de bron-Excel bevat geen "
                    f"tab 'Adoptie' of is niet bereikbaar. ({_par_err})"
                )
                st.session_state.pop("subsim_par_data", None)

        _par_df = st.session_state.get("subsim_par_data")
        if _par_df is not None and not _par_df.empty:
            import matplotlib.pyplot as _plt_par
            _valid = _par_df.dropna(subset=["margin", "surplus"]).reset_index(drop=True)
            if _valid.empty:
                st.info("Geen geldige (haalbare) (α,X)-combinaties om te plotten.")
            else:
                _m = _valid["margin"].to_numpy()
                _s = _valid["surplus"].to_numpy()
                _eff = np.ones(len(_valid), dtype=bool)
                for _i in range(len(_valid)):
                    for _j in range(len(_valid)):
                        if _i == _j:
                            continue
                        if (_m[_j] >= _m[_i] and _s[_j] >= _s[_i]
                                and (_m[_j] > _m[_i] or _s[_j] > _s[_i])):
                            _eff[_i] = False
                            break
                _figp, _axp = _plt_par.subplots(figsize=(8, 5.5))
                _sc = _axp.scatter(
                    _valid["margin"], _valid["surplus"],
                    c=_valid["alpha"], cmap="viridis", s=70,
                    edgecolor="white", linewidth=0.6, zorder=3)
                _cb = _figp.colorbar(_sc, ax=_axp)
                _cb.set_label("price percentage α")
                _front = _valid[_eff].sort_values("margin")
                _axp.plot(
                    _front["margin"], _front["surplus"],
                    color="#d62728", lw=2, marker="o", ms=9,
                    markerfacecolor="none", markeredgecolor="#d62728",
                    label="Pareto-frontier", zorder=4)
                _infeas = _valid[~_valid["feasible"]]
                if not _infeas.empty:
                    _axp.scatter(
                        _infeas["margin"], _infeas["surplus"],
                        facecolors="none", edgecolors="red", s=130,
                        linewidth=1.2, label="infeasible", zorder=5)
                _axp.axhline(0, color="grey", lw=0.8, ls=":")
                _axp.axvline(0, color="grey", lw=0.8, ls=":")
                _axp.set_xlabel("BPA margin (€)")
                _axp.set_ylabel("total customer surplus (€)")
                _axp.set_title("Pareto efficiency over (α, X)")
                _axp.grid(True, alpha=0.3)
                _axp.legend(loc="best", fontsize=9)
                _figp.tight_layout()
                st.pyplot(_figp)
                _plt_par.close(_figp)

                st.markdown("**Pareto-efficiënte (α, X)-combinaties**")
                _tab = _front.copy()
                _tab["α"]              = _tab["alpha"].map(lambda v: f"{v:.0%}")
                _tab["X"]              = _tab["X"].map(lambda v: f"{v:.3f}")
                _tab["BPA-marge (€)"]  = _tab["margin"].map(lambda v: f"{v:,.0f}")
                _tab["Klantsurplus (€)"] = _tab["surplus"].map(lambda v: f"{v:,.0f}")
                _tab["Σ E[Z]"]         = _tab["total_Z"].map(lambda v: f"{v:,.0f}")
                _tab["Haalbaar"]       = _tab["feasible"].map(lambda b: "✓" if b else "✗")
                st.dataframe(
                    _tab[["α", "X", "BPA-marge (€)", "Klantsurplus (€)",
                          "Σ E[Z]", "Haalbaar"]],
                    hide_index=True, use_container_width=True)
                st.download_button(
                    "⬇️ Download alle (α,X)-resultaten (CSV)",
                    _par_df.to_csv(index=False).encode("utf-8"),
                    file_name="pareto_alpha_X.csv", mime="text/csv",
                    key="subsim_par_dl")

    # ── Optimale α bij vast service level X ───────────────────────────────
    with st.expander("🎯 Optimale α bij vast service level X (marge maximaliseren)"):
        st.caption(
            "Zet het service level X vast en zoek het prijspercentage α dat de "
            "BPA-marge maximaliseert. Een hogere α verhoogt de omzet per klant "
            "maar verlaagt de adoptie q(α), dus er bestaat doorgaans een "
            "inwendig optimum."
        )
        _kp_opt = st.session_state.get("kosten_params", {})
        _kappa_bpa_opt = float(_kp_opt.get("kappa_bpa", 0.20))
        _kappa_c_opt   = float(_kp_opt.get("kappa_c", 0.25))
        st.caption(
            f"Kostenparameters uit Kostenanalyse: κ_BPA = **{_kappa_bpa_opt:.0%}**, "
            f"κ_c = **{_kappa_c_opt:.0%}** _(pas aan via 💰 Kostenanalyse)_."
        )

        _co1, _co2 = st.columns(2)
        with _co1:
            _X_opt = st.number_input(
                "Vast service level X", min_value=0.50, max_value=0.9999,
                value=float(_X_def_sim), step=0.005, format="%.3f",
                key="subsim_opt_X")
            _oa_min = st.number_input(
                "α-bereik min", min_value=0.0001, max_value=1.0, value=0.02,
                step=0.01, format="%.2f", key="subsim_opt_amin")
        with _co2:
            _oa_n = st.slider(
                "Aantal α-waarden", 5, 60, 25, key="subsim_opt_na")
            _oa_max = st.number_input(
                "α-bereik max", min_value=0.01, max_value=1.0, value=0.40,
                step=0.01, format="%.2f", key="subsim_opt_amax")
        _opt_feas = st.checkbox(
            "Alleen haalbare combinaties meenemen (marge ≥ 0 én alle klanten profiteren)",
            value=False, key="subsim_opt_feas")

        if st.button("🎯 Bereken optimale α", key="subsim_opt_btn",
                     disabled=not _cls_codes):
            try:
                _ov_opt = get_overzicht_df(cfg)
                if _ov_opt is None or _ov_opt.empty:
                    st.warning("Geen overzicht beschikbaar — laad eerst het overzicht (tab 📊).")
                else:
                    _oa_grid = list(np.linspace(float(_oa_min), float(_oa_max), int(_oa_n)))
                    with st.spinner("Optimale α zoeken…"):
                        _opt_curve, _opt_best = optimale_alpha_bij_X(
                            _ov_opt, float(_X_opt), _oa_grid,
                            float(_q_eq), float(_beta_r),
                            _kappa_bpa_opt, _kappa_c_opt,
                            excel_file=_excel_arg(), codes=_cls_codes,
                            alleen_haalbaar=bool(_opt_feas))
                    if _opt_curve is None or _opt_curve.empty or _opt_best is None:
                        st.warning("Geen geldige marge berekend — controleer de Adoptie-tab en selectie.")
                        st.session_state.pop("subsim_opt_data", None)
                    else:
                        st.session_state["subsim_opt_data"] = {
                            "curve": _opt_curve,
                            "best": _opt_best.to_dict(),
                            "X": float(_X_opt),
                        }
            except (ValueError, FileNotFoundError, OSError) as _opt_err:
                st.warning(
                    "Optimalisatie niet beschikbaar: de bron-Excel bevat geen "
                    f"tab 'Adoptie'. ({_opt_err})"
                )
                st.session_state.pop("subsim_opt_data", None)

        _opt_data = st.session_state.get("subsim_opt_data")
        if _opt_data:
            import matplotlib.pyplot as _plt_opt
            _ocurve = _opt_data["curve"].dropna(subset=["margin"]).sort_values("alpha")
            _obest  = _opt_data["best"]
            _oX     = _opt_data["X"]
            _z_lbl  = "Σ E[Z] (analytisch)"
            _cm1, _cm2, _cm3 = st.columns(3)
            _cm1.metric("Optimale α", f"{_obest['alpha']:.1%}")
            _cm2.metric("BPA-marge", f"€ {_obest['margin']:,.0f}")
            _cm3.metric(_z_lbl, f"{_obest['total_Z']:,.0f}",
                        help=f"Haalbaar: {'ja' if _obest['feasible'] else 'nee'}")
            _figo, _axo = _plt_opt.subplots(figsize=(9, 4.5))
            _axo.plot(_ocurve["alpha"], _ocurve["margin"],
                      color="#1f77b4", lw=2, marker="o", ms=4,
                      label="BPA margin")
            _axo.axvline(_obest["alpha"], color="#d62728", ls="--", lw=1.5,
                         label=f"optimal α = {_obest['alpha']:.1%}")
            _axo.axhline(0, color="grey", lw=0.8, ls=":")
            _axo.set_xlabel("price percentage α")
            _axo.set_ylabel("BPA margin (€)")
            _axo.set_title(f"Margin vs. α at fixed X = {_oX:.3f}  (analytical E[Z])")
            _axo.grid(True, alpha=0.3)
            _axo2 = _axo.twinx()
            _axo2.plot(_ocurve["alpha"], _ocurve["total_Z"],
                       color="#2ca02c", lw=1.5, ls="-.", alpha=0.7,
                       label=_z_lbl)
            _axo2.set_ylabel(f"total subscriptions  {_z_lbl}", color="#2ca02c")
            _axo2.tick_params(axis="y", labelcolor="#2ca02c")
            _l1, _lab1 = _axo.get_legend_handles_labels()
            _l2, _lab2 = _axo2.get_legend_handles_labels()
            _axo.legend(_l1 + _l2, _lab1 + _lab2, loc="best", fontsize=9)
            _figo.tight_layout()
            st.pyplot(_figo)
            _plt_opt.close(_figo)

            st.markdown("**Revenue, costs en stocklevels per α** (bij vast X)")
            _tab_opt = _ocurve.copy()
            _tab_opt["α"]              = _tab_opt["alpha"].map(lambda v: f"{v:.1%}")
            _tab_opt["BPA-marge (€)"]  = _tab_opt["margin"].map(lambda v: f"{v:,.0f}")
            _tab_opt["Revenue (€)"]    = _tab_opt["revenue"].map(lambda v: f"{v:,.0f}")
            _tab_opt["Costs (€)"]      = _tab_opt["costs"].map(lambda v: f"{v:,.0f}")
            _tab_opt["Stocklevel (units)"] = _tab_opt["stock_level"].map(
                lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
            st.dataframe(
                _tab_opt[["α", "BPA-marge (€)", "Revenue (€)", "Costs (€)",
                          "Stocklevel (units)"]],
                hide_index=True, use_container_width=True)
            st.download_button(
                "⬇️ Download marge-vs-α curve (CSV)",
                _opt_data["curve"].to_csv(index=False).encode("utf-8"),
                file_name="optimale_alpha_bij_X.csv", mime="text/csv",
                key="subsim_opt_dl")

    # ── β_r-onzekerheidsband: uniforme scenario-parameter ─────────────────
    with st.expander("🎲 β_r-onzekerheidsband: bandbreedte van winst en optimale α"):
        st.caption(
            "$β_r$ kan **niet** uit historische subscriptie-data worden "
            "geschat en wordt daarom als **onzekere scenario-parameter** "
            "behandeld. Bij gebrek aan voorkennis over welke waarde binnen een "
            "plausibel bereik het meest waarschijnlijk is, gebruiken we een "
            "**uniforme** verdeling $β_r\\sim U(β_r^{min}, β_r^{max})$. Per "
            "trekking wordt (bij vast $X$) de volledige keten "
            "$α\\to q(α)\\to E[Z_i]\\to$ kostenmodel doorgerekend. Zo ontstaat "
            "per α een verdeling van de verwachte BPA-winst $E[Π_{BPA}(α)]$ "
            "(P5/P50/P95-band) en een verdeling van de winst-maximaliserende "
            "$α^*$."
        )
        _kp_bru = st.session_state.get("kosten_params", {})
        _kappa_bpa_bru = float(_kp_bru.get("kappa_bpa", 0.20))
        _kappa_c_bru   = float(_kp_bru.get("kappa_c", 0.25))
        st.caption(
            f"Kostenparameters uit Kostenanalyse: κ_BPA = **{_kappa_bpa_bru:.0%}**, "
            f"κ_c = **{_kappa_c_bru:.0%}**. $q_{{eq}}$ = **{_q_eq:.2f}** "
            "(uit het adoptiemodel hierboven)."
        )

        _cbr1, _cbr2 = st.columns(2)
        with _cbr1:
            _X_bru = st.number_input(
                "Vast service level X", min_value=0.50, max_value=0.9999,
                value=float(_X_def_sim), step=0.005, format="%.3f",
                key="subsim_bru_X")
            _bra_min = st.number_input(
                "α-bereik min", min_value=0.0001, max_value=1.0, value=0.02,
                step=0.01, format="%.2f", key="subsim_bru_amin")
            _bra_max = st.number_input(
                "α-bereik max", min_value=0.01, max_value=1.0, value=0.40,
                step=0.01, format="%.2f", key="subsim_bru_amax")
        with _cbr2:
            _br_lo = st.number_input(
                "β_r^min", min_value=0.0, max_value=20.0, value=0.5,
                step=0.1, format="%.2f", key="subsim_bru_brlo",
                help="Ondergrens van het plausibele β_r-bereik.")
            _br_hi = st.number_input(
                "β_r^max", min_value=0.1, max_value=20.0, value=3.5,
                step=0.1, format="%.2f", key="subsim_bru_brhi",
                help="Bovengrens van het plausibele β_r-bereik.")
            _bra_n = st.slider(
                "Aantal α-waarden", 5, 60, 25, key="subsim_bru_na")
        _cbr3, _cbr4 = st.columns(2)
        with _cbr3:
            _bru_K = st.slider(
                "Aantal β_r-trekkingen (K)", 20, 1000, 200, step=20,
                key="subsim_bru_K",
                help="Meer trekkingen → gladdere band, maar langere rekentijd.")
        with _cbr4:
            _bru_feas = st.checkbox(
                "α* alleen over haalbare α zoeken",
                value=False, key="subsim_bru_feas",
                help="Bepaal de optimale α* per trekking alleen over haalbare α "
                     "(marge ≥ 0 én alle klanten profiteren).")
        _bru_seed = st.checkbox(
            "Reproduceerbare trekkingen (vaste seed)", value=True,
            key="subsim_bru_seed")

        if st.button("🎲 Bereken β_r-band", key="subsim_bru_btn",
                     disabled=not _cls_codes):
            try:
                _ov_bru = get_overzicht_df(cfg)
                if _ov_bru is None or _ov_bru.empty:
                    st.warning("Geen overzicht beschikbaar — laad eerst het overzicht (tab 📊).")
                elif float(_br_hi) <= float(_br_lo):
                    st.warning("β_r^max moet groter zijn dan β_r^min.")
                else:
                    _bra_grid = list(np.linspace(float(_bra_min), float(_bra_max), int(_bra_n)))
                    with st.spinner(
                            f"β_r-band berekenen ({int(_bru_K)} trekkingen × "
                            f"{int(_bra_n)} α-waarden)…"):
                        _bru_res = beta_r_winstband(
                            _ov_bru, float(_X_bru), _bra_grid,
                            float(_q_eq), float(_br_lo), float(_br_hi),
                            _kappa_bpa_bru, _kappa_c_bru,
                            n_samples=int(_bru_K),
                            excel_file=_excel_arg(), codes=_cls_codes,
                            seed=(42 if _bru_seed else None),
                            alleen_haalbaar=bool(_bru_feas))
                    if _bru_res is None:
                        st.warning("Geen resultaten — controleer de Adoptie-tab en selectie.")
                        st.session_state.pop("subsim_bru_data", None)
                    else:
                        _bru_res["X"] = float(_X_bru)
                        _bru_res["beta_r_min"] = float(_br_lo)
                        _bru_res["beta_r_max"] = float(_br_hi)
                        st.session_state["subsim_bru_data"] = _bru_res
            except (ValueError, FileNotFoundError, OSError) as _bru_err:
                st.warning(
                    "β_r-band niet beschikbaar: de bron-Excel bevat geen tab "
                    f"'Adoptie' of is niet bereikbaar. ({_bru_err})"
                )
                st.session_state.pop("subsim_bru_data", None)

        _bru = st.session_state.get("subsim_bru_data")
        if _bru:
            import matplotlib.pyplot as _plt_bru
            _ag   = _bru["alpha_grid"]
            _p5   = _bru["margin_pct"].get(5)
            _p50  = _bru["margin_pct"].get(50)
            _p95  = _bru["margin_pct"].get(95)
            _mean = _bru["margin_mean"]
            _oap  = _bru["opt_alpha_pct"]

            _cbm1, _cbm2, _cbm3 = st.columns(3)
            _cbm1.metric("α* — P5", f"{_oap.get(5, float('nan')):.1%}")
            _cbm2.metric("α* — mediaan (P50)", f"{_oap.get(50, float('nan')):.1%}")
            _cbm3.metric("α* — P95", f"{_oap.get(95, float('nan')):.1%}")
            st.caption(
                "Onder onzekerheid in de klantgevoeligheid β_r ligt de optimale "
                f"α meestal tussen **{_oap.get(5, float('nan')):.1%}** en "
                f"**{_oap.get(95, float('nan')):.1%}** "
                f"(mediaan **{_oap.get(50, float('nan')):.1%}**), bij "
                f"X = {_bru['X']:.3f} en β_r ~ U({_bru['beta_r_min']:.2f}, "
                f"{_bru['beta_r_max']:.2f})."
            )

            # ── Grafiek 1: winst-band vs α ────────────────────────────────
            _figbr, _axbr = _plt_bru.subplots(figsize=(9, 4.8))
            if _p5 is not None and _p95 is not None:
                _axbr.fill_between(
                    _ag, _p5, _p95, color="#1f77b4", alpha=0.20,
                    label="P5–P95 band")
            if _p50 is not None:
                _axbr.plot(_ag, _p50, color="#1f77b4", lw=2.2,
                           label="mediaan (P50)")
            if _mean is not None:
                _axbr.plot(_ag, _mean, color="#ff7f0e", lw=1.4, ls="--",
                           label="gemiddelde E[Π]")
            # Mediaan-optimum markeren.
            if _p50 is not None and np.isfinite(_p50).any():
                _ix = int(np.nanargmax(_p50))
                _axbr.axvline(_ag[_ix], color="#d62728", ls=":", lw=1.4,
                              label=f"α*(P50) = {_ag[_ix]:.1%}")
            _axbr.axhline(0, color="grey", lw=0.8, ls=":")
            _axbr.set_xlabel("α")
            _axbr.set_ylabel("expected BPA profit  E[Π_BPA] (€)")
            _axbr.set_title(
                f"Profit band under β_r ~ U({_bru['beta_r_min']:.2f}, "
                f"{_bru['beta_r_max']:.2f})  at X = {_bru['X']:.3f}")
            _axbr.grid(True, alpha=0.3)
            _axbr.legend(loc="best", fontsize=9)
            _figbr.tight_layout()
            st.pyplot(_figbr)
            _plt_bru.close(_figbr)

            # ── Grafiek 2: verdeling van optimale α* ──────────────────────
            _opt_a_valid = _bru["opt_alpha"][~np.isnan(_bru["opt_alpha"])]
            if _opt_a_valid.size:
                _fighx, _axhx = _plt_bru.subplots(figsize=(9, 3.6))
                _axhx.hist(_opt_a_valid * 100.0, bins=min(30, max(5, _opt_a_valid.size // 5)),
                           color="#2ca02c", alpha=0.75, edgecolor="white")
                for _p, _c, _lbl in ((5, "#d62728", "P5"), (50, "#000000", "P50"),
                                     (95, "#d62728", "P95")):
                    _axhx.axvline(_oap.get(_p, float('nan')) * 100.0, color=_c,
                                  ls="--", lw=1.4,
                                  label=f"{_lbl} = {_oap.get(_p, float('nan')):.1%}")
                _axhx.set_xlabel("optimal price percentage α* (%)")
                _axhx.set_ylabel("frequency")
                _axhx.set_title("Distribution of profit-maximising α* over β_r draws")
                _axhx.grid(True, alpha=0.3)
                _axhx.legend(loc="best", fontsize=9)
                _fighx.tight_layout()
                st.pyplot(_fighx)
                _plt_bru.close(_fighx)

            # ── Downloadbare band-tabel ───────────────────────────────────
            _band_df = pd.DataFrame({
                "alpha":      _ag,
                "P5":         _p5   if _p5   is not None else np.nan,
                "P50":        _p50  if _p50  is not None else np.nan,
                "P95":        _p95  if _p95  is not None else np.nan,
                "mean":       _mean if _mean is not None else np.nan,
            })
            st.download_button(
                "⬇️ Download winst-band per α (CSV)",
                _band_df.to_csv(index=False).encode("utf-8"),
                file_name="beta_r_winstband.csv", mime="text/csv",
                key="subsim_bru_dl")


# ─────────────────────────────────────────────────────────────────────────────────
#  TAB 12 – SENSITIVITY (WTP)  – elementen van de adoptie-/WTP-functie plotten
# ─────────────────────────────────────────────────────────────────────────────────

with tab_sensitivity:
    st.subheader("Sensitivity-analyse van de WTP-functie")
    st.markdown(
        "Kies zelf de **afhankelijke (y-as)** uitkomst en de **onafhankelijke (x-as)** "
        "variabele uit de adoptie-/kostenketen. Optioneel varieer je een **tweede** "
        "element om een familie van curves te tekenen, zodat je twee variabelen direct "
        "tegen elkaar kunt afwegen.\n\n"
        "Per parameterpunt loopt de volledige keten: "
        "$\\text{parameters}\\to q(α)\\to E[Z_i]\\to$ kostenmodel "
        "$\\to$ **BPA-marge / klantsurplus**. De adoptie volgt het globale "
        "logit-model $q(α)=σ\\!\\big(\\operatorname{logit}(q_{eq})+β_r\\ln(κ_c/α)\\big)$; "
        "de kans hangt alleen van α af (via κ_c/α), X werkt via de "
        "voorraad-/kostenkant."
    )

    # ── Registry van WTP-elementen (label, grenzen, stap, default) ─────────
    _WTP_PARAMS = {
        "alpha":   {"label": "α — prijspercentage",            "axis": "α — price percentage",         "min": 0.0001, "max": 1.0,    "step": 0.01,  "fmt": "%.3f"},
        "X":       {"label": "X — service level",              "axis": "X — service level",            "min": 0.50,   "max": 0.9999, "step": 0.005, "fmt": "%.3f"},
        "q_eq":    {"label": "q_eq — adoptie bij pariteit",     "axis": "q_eq — adoption at parity",    "min": 0.01,   "max": 0.99,   "step": 0.05,  "fmt": "%.3f"},
        "beta_r":  {"label": "β_r — kostenratio-gevoeligheid",  "axis": "β_r — cost-ratio sensitivity", "min": 0.0,    "max": 20.0,   "step": 0.1,   "fmt": "%.2f"},
        "kappa_c": {"label": "κ_c — kostenpariteit",            "axis": "κ_c — cost parity",            "min": 0.01,   "max": 1.0,    "step": 0.01,  "fmt": "%.3f"},
    }

    # Startwaarden overnemen uit de Subscriptie-simulatie / Kostenanalyse-tab.
    _kp_se = st.session_state.get("kosten_params", {})
    _seed_se = {
        "alpha":   float(_kp_se.get("alpha", 0.15)),
        "X":       float(_kp_se.get("service_level", 0.99)),
        "q_eq":    float(st.session_state.get("subsim_q_eq", 0.55)),
        "beta_r":  float(st.session_state.get("subsim_beta_r", 1.0)),
        "kappa_c": float(_kp_se.get("kappa_c", 0.25)),
    }

    def _clip_se(_name, _val):
        _spec = _WTP_PARAMS[_name]
        return float(min(max(_val, _spec["min"]), _spec["max"]))

    _labels_se = {k: v["label"] for k, v in _WTP_PARAMS.items()}

    # ── Registry van afhankelijke (y-as) uitkomsten ───────────────────────
    _Y_METRICS = {
        "bpa_margin": {"label": "Totale BPA-winst (€)",              "axis": "total BPA profit (€)",                  "tbl": "BPA-winst (€)",    "kind": "euro"},
        "surplus":    {"label": "Totaal klantsurplus (€)",           "axis": "total customer surplus (€)",            "tbl": "Klantsurplus (€)", "kind": "euro"},
        "total_Z":    {"label": "Verwacht aantal subscripties E[Z]", "axis": "expected number of subscriptions E[Z]", "tbl": "E[Z]",             "kind": "num"},
        "q":          {"label": "Adoptiekans q(α)",                  "axis": "adoption probability q(α)",             "tbl": "q(α)",             "kind": "pct"},
    }

    # ── Afhankelijke (y-as) variabele ─────────────────────────────────────
    _y_var = st.selectbox(
        "Y-as variabele (afhankelijk)", options=list(_Y_METRICS.keys()),
        format_func=lambda k: _Y_METRICS[k]["label"], index=0, key="se_y_var",
        help="De keten-uitkomst die tegen de gekozen x-as wordt geplot.")
    _y_spec = _Y_METRICS[_y_var]

    # ── Onafhankelijke as-keuzes ──────────────────────────────────────────
    _cx, _ccurve = st.columns(2)
    with _cx:
        _x_var = st.selectbox(
            "X-as variabele", options=list(_WTP_PARAMS.keys()),
            format_func=lambda k: _labels_se[k], index=0, key="se_x_var",
            help="Het WTP-element dat over de x-as wordt gevarieerd.")
    with _ccurve:
        _curve_opts = ["(geen)"] + [k for k in _WTP_PARAMS if k != _x_var]
        # Als de opgeslagen curve-variabele gelijk is aan de nieuw gekozen
        # x-variabele, reset dan naar "(geen)" — anders gooit Streamlit een
        # exception (waarde niet in opties) en valt de tab-layout uit elkaar.
        if st.session_state.get("se_curve_var", "(geen)") == _x_var:
            st.session_state["se_curve_var"] = "(geen)"
        _curve_var = st.selectbox(
            "Curve-variabele (optioneel)", options=_curve_opts,
            format_func=lambda k: _labels_se.get(k, k), index=0, key="se_curve_var",
            help="Een tweede element dat per curve verandert (familie van lijnen).")

    # ── X-as bereik ───────────────────────────────────────────────────────
    _spec_x = _WTP_PARAMS[_x_var]
    st.markdown(f"**X-as bereik — {_spec_x['label']}**")
    # Opgeslagen min/max-waarden van een vorige x-variabele kunnen buiten de
    # grenzen van de nieuwe variabele vallen -> Streamlit exception -> tab breekt.
    for _se_key, _se_def in (("se_x_min", _spec_x["min"]), ("se_x_max", _spec_x["max"])):
        if _se_key in st.session_state:
            _sv = st.session_state[_se_key]
            if not (_spec_x["min"] <= _sv <= _spec_x["max"]):
                st.session_state[_se_key] = _se_def
    _cx1, _cx2, _cx3 = st.columns(3)
    with _cx1:
        _x_min = st.number_input(
            "min", min_value=_spec_x["min"], max_value=_spec_x["max"],
            value=_spec_x["min"], step=_spec_x["step"], format=_spec_x["fmt"],
            key="se_x_min")
    with _cx2:
        _x_max = st.number_input(
            "max", min_value=_spec_x["min"], max_value=_spec_x["max"],
            value=_spec_x["max"], step=_spec_x["step"], format=_spec_x["fmt"],
            key="se_x_max")
    with _cx3:
        _x_n = st.slider("gridpunten", min_value=5, max_value=200, value=60,
                         key="se_x_n")

    # ── Curve-variabele waarden ───────────────────────────────────────────
    _curve_vals = [None]
    if _curve_var != "(geen)":
        _spec_c = _WTP_PARAMS[_curve_var]
        st.markdown(f"**Curve-waarden — {_spec_c['label']}**")
        # Zelfde guard als voor se_x_min/se_x_max: reset bij wisselen curve-var.
        for _se_key, _se_def in (("se_c_min", _clip_se(_curve_var, _seed_se[_curve_var])), ("se_c_max", _spec_c["max"])):
            if _se_key in st.session_state:
                _sv = st.session_state[_se_key]
                if not (_spec_c["min"] <= _sv <= _spec_c["max"]):
                    st.session_state[_se_key] = _se_def
        _cc1, _cc2, _cc3 = st.columns(3)
        with _cc1:
            _c_min = st.number_input(
                "curve min", min_value=_spec_c["min"], max_value=_spec_c["max"],
                value=_clip_se(_curve_var, _seed_se[_curve_var]),
                step=_spec_c["step"], format=_spec_c["fmt"], key="se_c_min")
        with _cc2:
            _c_max = st.number_input(
                "curve max", min_value=_spec_c["min"], max_value=_spec_c["max"],
                value=_spec_c["max"], step=_spec_c["step"], format=_spec_c["fmt"],
                key="se_c_max")
        with _cc3:
            _c_n = st.slider("aantal curves", min_value=1, max_value=8, value=4,
                             key="se_c_n")
        _curve_vals = list(np.linspace(float(_c_min), float(_c_max), int(_c_n)))

    # ── Vaste waarden voor de overige elementen ───────────────────────────
    _fixed = {k: _clip_se(k, _seed_se[k]) for k in _WTP_PARAMS}
    _vrij = [k for k in _WTP_PARAMS if k not in (_x_var, _curve_var)]
    with st.expander("⚙️ Vaste waarden voor de overige elementen", expanded=True):
        _fc = st.columns(2)
        for _i, _k in enumerate(_vrij):
            _spec = _WTP_PARAMS[_k]
            with _fc[_i % 2]:
                _fixed[_k] = st.number_input(
                    _spec["label"], min_value=_spec["min"], max_value=_spec["max"],
                    value=_clip_se(_k, _seed_se[_k]), step=_spec["step"],
                    format=_spec["fmt"], key=f"se_fix_{_k}")

    # (WTP-plafond vervallen: adoptie hangt via κ_c/α af van α; geen aparte poort.)

    # ── Kostenparameters + bron ───────────────────────────────────────────
    _kappa_bpa_se = float(_kp_se.get("kappa_bpa", 0.20))
    _kappa_c_se   = float(_kp_se.get("kappa_c", 0.25))
    st.caption(
        f"Kostenparameters uit Kostenanalyse: κ_BPA = **{_kappa_bpa_se:.0%}**, "
        f"κ_c = **{_kappa_c_se:.0%}** _(pas aan via 💰 Kostenanalyse)_."
    )

    _cls_codes_se = sorted(get_classificatie_info().get("items", {}).keys())
    if not _cls_codes_se:
        st.warning(
            "Geen classificatie-selectie (`bpa_selectie.json`) gevonden. "
            "Voer eerst de classificatie uit via tab 🏷️ Classificatie."
        )

    # Bron-Excel met de tab 'Adoptie': zelfde resolutie als de Subscriptie-
    # simulatie-tab — eerst de daar geüploade Excel (`subsim_upload`), dan de
    # classificatie-upload (`cls_upload`), anders de repo-Excel.
    _excel_se = st.session_state.get("subsim_upload") or st.session_state.get("cls_upload")
    if _excel_se is not None:
        st.caption(f"Bron-Excel: **{getattr(_excel_se, 'name', 'geüploade Excel')}**")
    else:
        st.caption(f"Bron-Excel: repo-Excel (`{os.path.basename(EXCEL_PATH)}`)")

    def _excel_arg_se():
        if _excel_se is None:
            return None
        try:
            _excel_se.seek(0)
        except (AttributeError, ValueError):
            pass
        return _excel_se

    _x_grid = list(np.linspace(float(_x_min), float(_x_max), int(_x_n)))

    # ── Param-dicts voor alle (x, curve)-punten ───────────────────────────
    def _bouw_param_dicts():
        _dicts = []
        for _cval in _curve_vals:
            for _xv in _x_grid:
                _p = dict(_fixed)
                _p[_x_var] = float(_xv)
                if _cval is not None:
                    _p[_curve_var] = float(_cval)
                _dicts.append(_p)
        return _dicts

    if st.button("📊 Bereken sensitivity", type="primary",
                 disabled=not _cls_codes_se, key="se_bereken"):
        try:
            _ov_se = get_overzicht_df(cfg)
        except Exception as _e:
            _ov_se = None
            st.error(f"Kon overzicht niet laden: {_e}")
        if _ov_se is None or _ov_se.empty:
            st.warning("Geen overzicht beschikbaar — laad eerst het overzicht (tab 📊).")
        else:
            with st.spinner("Winst-sensitivity berekenen via het kostenmodel…"):
                try:
                    _recs = metrieken_voor_wtp_grid(
                        _ov_se, _bouw_param_dicts(),
                        _kappa_bpa_se, _kappa_c_se,
                        excel_file=_excel_arg_se(), codes=_cls_codes_se,
                    )
                    _yvals = [r.get(_y_var, float("nan")) for r in _recs]
                    _nx = len(_x_grid)
                    _per_curve = [
                        _yvals[_i * _nx:(_i + 1) * _nx]
                        for _i in range(len(_curve_vals))
                    ]
                    st.session_state["se_resultaat"] = {
                        "x_grid":     _x_grid,
                        "per_curve":  _per_curve,
                        "x_var":      _x_var,
                        "curve_var":  _curve_var,
                        "curve_vals": _curve_vals,
                        "x_label":    _spec_x["axis"],
                        "x_fmt":      _spec_x["fmt"],
                        "x_cur":      _clip_se(_x_var, _seed_se[_x_var]),
                        "y_axis":     _y_spec["axis"],
                        "y_label":    _y_spec["axis"],
                        "y_tbl":      _y_spec["tbl"],
                        "y_kind":     _y_spec["kind"],
                    }
                except ValueError as _se_err:
                    st.warning(
                        "Berekening niet beschikbaar: de bron-Excel bevat geen "
                        f"tab 'Adoptie'. ({_se_err})"
                    )
                    st.session_state.pop("se_resultaat", None)

    # ── Plot van het laatst berekende resultaat ───────────────────────────
    _res_se = st.session_state.get("se_resultaat")
    if _res_se:
        import matplotlib.pyplot as _plt_se
        import matplotlib.ticker as _mt_se

        _xg        = _res_se["x_grid"]
        _per_curve = _res_se["per_curve"]
        _cvals     = _res_se["curve_vals"]
        _cv_var    = _res_se["curve_var"]
        _x_lbl     = _res_se["x_label"]
        _x_fmt     = _res_se["x_fmt"]
        _y_axis    = _res_se.get("y_axis", "total BPA profit (€)")
        _y_lbl     = _res_se.get("y_label", "total BPA profit (€)")
        _y_tbl     = _res_se.get("y_tbl", "BPA-winst (€)")
        _y_kind    = _res_se.get("y_kind", "euro")

        _fig_se, _ax_se = _plt_se.subplots(figsize=(10, 5))
        _cmap_se = _plt_se.cm.viridis
        for _ci, _cval in enumerate(_cvals):
            _ys = _per_curve[_ci]
            if _cval is None:
                _ax_se.plot(_xg, _ys, color="#1f77b4", lw=2.2, marker="o", ms=3)
            else:
                _col = _cmap_se(_ci / max(1, len(_cvals) - 1))
                _spec_c = _WTP_PARAMS[_cv_var]
                _ax_se.plot(_xg, _ys, lw=2.0, color=_col, marker="o", ms=3,
                            label=f"{_spec_c['label'].split(' ')[0]} = {_cval:{_spec_c['fmt'][1:]}}")

        _x_cur = _res_se["x_cur"]
        if _xg and min(_xg) <= _x_cur <= max(_xg):
            _ax_se.axvline(_x_cur, color="grey", ls="--", lw=1,
                           label=f"current {_x_lbl.split(' ')[0]} = {_x_cur:{_x_fmt[1:]}}")
        _ax_se.axhline(0.0, color="black", lw=0.8, alpha=0.6)
        _ax_se.set_xlabel(_x_lbl, fontsize=11)
        _ax_se.set_ylabel(_y_axis, fontsize=11)
        _ax_se.set_title(f"Sensitivity of {_y_lbl} w.r.t. the WTP elements", fontsize=12)
        if _y_kind == "euro":
            _yfmt_se = _mt_se.FuncFormatter(lambda v, _: f"€{v:,.0f}")
        elif _y_kind == "pct":
            _yfmt_se = _mt_se.FuncFormatter(lambda v, _: f"{v:.0%}")
        else:
            _yfmt_se = _mt_se.FuncFormatter(lambda v, _: f"{v:,.0f}")
        _ax_se.yaxis.set_major_formatter(_yfmt_se)
        _ax_se.grid(True, alpha=0.3)
        if _cv_var != "(geen)" or (_xg and min(_xg) <= _x_cur <= max(_xg)):
            _ax_se.legend(fontsize=9)
        _fig_se.tight_layout()
        st.pyplot(_fig_se)
        _plt_se.close(_fig_se)

        # ── Datatabel + download ──────────────────────────────────────────
        with st.expander("📋 Data achter de grafiek"):
            _tbl = {_x_lbl.split(" ")[0]: _xg}
            for _ci, _cval in enumerate(_cvals):
                if _cval is None:
                    _tbl[_y_tbl] = _per_curve[_ci]
                else:
                    _spec_c = _WTP_PARAMS[_cv_var]
                    _tbl[f"{_y_tbl} @ {_spec_c['label'].split(' ')[0]}={_cval:{_spec_c['fmt'][1:]}}"] = _per_curve[_ci]
            _df_se = pd.DataFrame(_tbl)
            st.dataframe(_df_se, use_container_width=True, height=320)
            st.download_button(
                "⬇️ Download sensitivity-data (CSV)",
                data=_df_se.to_csv(sep=";", decimal=",", index=False).encode("utf-8"),
                file_name=f"wtp_sensitivity_{date.today()}.csv",
                mime="text/csv",
            )
    else:
        st.info("Stel de parameters in en klik op **📊 Bereken sensitivity**.")



