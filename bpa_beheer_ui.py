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
import json

# Hergebruik alle logica uit bpa_beheer.py
from bpa_beheer import (
    laad_config,
    sla_config_op,
    bereken_overzicht,
    bouw_model_kosten,
    laad_excel_onderdelen,
    SERVICE_LEVELS,
    CONFIG_PATH,
    HISTORY_PATH,
    SCRIPT_DIR,
)
from model import BPAOptimizationModel

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
        _df = bereken_overzicht(cfg, _excel_file)
    if not _df.empty:
        st.session_state.overzicht_df = _df

# ══════════════════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════════════════

tab_overzicht, tab_subscripties, tab_toevoegen, tab_verwijderen, tab_config, tab_historie, tab_kosten, tab_drempel = st.tabs([
    "📊 Overzicht",
    "✏️ Subscripties aanpassen",
    "➕ Component toevoegen",
    "🗑️ Component verwijderen",
    "⚙️ Configuratie",
    "📈 Historiek",
    "💰 Kostenanalyse",
    "🔢 Subscriptiedrempel",
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
        f"Standaard N: **{cfg['standaard_n_klanten']}** · "
        f"Configuratie bijgewerkt: **{cfg['aangepast']}** · "
        f"Excel gewijzigd: **{_excel_mtime}**"
    )

    if st.button("🔄 Herbereken (laadt Excel opnieuw)"):
        with st.spinner("Berekenen…"):
            df = bereken_overzicht(cfg, _excel_file)
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
        for _sc in sl_cols:
            _dc = f"\u0394{_sc}"
            _delta_cols.append(_dc)
            def _calc_delta(row, _sc=_sc):
                _code = str(row.get('Code', ''))
                _prev = _prev_comp.get(_code, {})
                if _sc in _prev:
                    return int(row[_sc]) - int(_prev[_sc])
                return float('nan')
            _df_disp[_dc] = _df_disp.apply(_calc_delta, axis=1)

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
    st.subheader("Standaard aantal subscripties")

    nieuw_standaard = st.number_input(
        "Standaard N (geldt voor alle componenten zonder override)",
        min_value=1,
        value=cfg["standaard_n_klanten"],
        step=1,
    )
    if st.button("Opslaan standaard"):
        # Bewaar huidige overzicht_df als vorige snapshot vóór recompute
        if "overzicht_df" in st.session_state:
            st.session_state.overzicht_df_prev = st.session_state.overzicht_df.copy()
        cfg["standaard_n_klanten"] = int(nieuw_standaard)
        sla_config_op(cfg)
        st.toast(f"Standaard aangepast naar {cfg['standaard_n_klanten']}", icon="✅")
        st.session_state.pop("overzicht_df", None)
        st.rerun()

    st.divider()
    st.subheader("Overrides per artikelcode")
    st.caption("N = aantal subscripties, IP = inkoopprijs (€), LT = levertijd (dagen). "
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
            "N":           st.column_config.NumberColumn("N (subscripties)", min_value=1, step=1),
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
        st.toast(f"Overrides opgeslagen — {len(n_ov)} N, {len(ip_ov)} IP, {len(lt_ov)} LT.", icon="✅")
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
                "Aantal subscripties",
                min_value=1, value=cfg["standaard_n_klanten"], step=1,
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

    try:
        excel_codes = list(laad_excel_onderdelen(_excel_file).index)
    except Exception:
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
    col1.metric("Standaard N", cfg["standaard_n_klanten"])
    col2.metric("N-overrides", len(cfg.get("n_klanten_overrides", {})))
    col3.metric("IP/LT-overrides",
                max(len(cfg.get("ip_overrides", {})), len(cfg.get("lt_overrides", {}))))
    col4.metric("Uitgesloten", len(cfg.get("uitgesloten_componenten", [])))
    st.write(f"Aangemaakt: `{cfg['aangemaakt']}`  |  Aangepast: `{cfg['aangepast']}`")

    if cfg["n_klanten_overrides"]:
        st.write("**Overrides:**")
        st.dataframe(
            pd.DataFrame([{"Code": k, "N": v}
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
                    "N":           v.get("n_klanten", "std"),
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
                row = {"Datum": h["datum"], "N": h["n_klanten"], "# componenten": h["n_actief"]}
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

                ax.set_xlabel("Datum update", fontsize=11)
                ax.set_ylabel("Totale basisvoorraad (stuks)", fontsize=11)
                ax.set_title("Totale basisvoorraad BPA per update-moment", fontsize=12)
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
            _ax1.set_ylabel('Jaarlijkse marge (€)', fontsize=11)
            _ax1.set_title(
                f'Marge vs. service level  '
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
                          marker='s', linewidth=2, color='#FF9800', label='Totaal S*')
                for _xv, _yv in _pts2:
                    _ax2.annotate(str(int(_yv)), (_xv, _yv),
                                  textcoords='offset points', xytext=(0, 7),
                                  ha='center', fontsize=9)
            _ax2.set_xlabel('Service level (%)', fontsize=11)
            _ax2.set_ylabel('Totale basisvoorraad (stuks)', fontsize=11)
            _ax2.set_title(
                f'Basisvoorraad vs. service level  '
                f'(α = {_ALPHA_DEF:.0%}, N = standaard)',
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
            _ax3.set_xlabel('Abonnementstarief α (%)', fontsize=11)
            _ax3.set_ylabel('Jaarlijkse marge (€)', fontsize=11)
            _ax3.set_title(
                f'Marge vs. abonnementstarief  '
                f'(κ_BPA = {_KAPPA_BPA_DEF:.0%}, N = standaard)',
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
        st.subheader("N vs. haalbaarheid per α")
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
            _n_std = int(cfg['standaard_n_klanten'])

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
                           label=f'N huidig = {_n_std}')
            _ax_nf.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_nf.set_ylabel('Jaarlijkse BPA-marge (€)', fontsize=11)
            _ax_nf.set_title(
                f'BPA-marge vs. N per α  '
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
            _ax_hm.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_hm.set_ylabel('α', fontsize=12)
            _ax_hm.set_title(
                'Haalbaarheid BPA per (N, α)  (✓ = haalbaar, ✗ = niet haalbaar)',
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
        st.subheader("Haalbaarheid BPA per (N, serviceniveau)")
        st.caption(
            "Groen = BPA is haalbaar (marge ≥ 0), rood = niet haalbaar. "
            "α wordt overgenomen uit tabblad 💰 Kostenanalyse; κ_BPA en κ_c idem."
        )

        _N_NSL_VALS = [1, 2, 3, 5, 7, 10, 15, 20]

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
            _n_std_nsl = int(cfg['standaard_n_klanten'])

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
            _plt_nsl.colorbar(_im_nsl, ax=_ax_nsl, label='BPA-marge (€)', fraction=0.03, pad=0.02)

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
            _ax_nsl.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_nsl.set_ylabel('Service level', fontsize=11)
            _ax_nsl.set_title(
                f'Haalbaarheid BPA per (N, service level)  '
                f'(α = {_a_lbl:.0%}, κ_BPA = {_kb_lbl:.0%})',
                fontsize=12,
            )

            # Markeer huidige N
            try:
                _ni_std = min(range(len(_cols_nsl)),
                              key=lambda k: abs(_cols_nsl[k] - _n_std_nsl))
                _ax_nsl.axvline(_ni_std, color='black', linewidth=2.0, linestyle=':')
                _ax_nsl.text(_ni_std + 0.15, -0.7, f'N={_n_std_nsl}',
                             fontsize=8, color='black')
            except Exception:
                pass

            _fig_nsl.tight_layout()
            st.pyplot(_fig_nsl)
            _plt_nsl.close(_fig_nsl)

        # ── N vs. maximaal haalbaar serviceniveau ───────────────────────────────────
        st.divider()
        st.subheader("N vs. maximaal haalbaar serviceniveau")
        st.caption(
            "Voor elk aantal subscripties (N): wat is het hoogste service level waarbij het "
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
            _n_std_sl = int(cfg['standaard_n_klanten'])

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
                           label=f'N huidig = {_n_std_sl}')
            _ax_bs.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_bs.set_ylabel('Totale basisvoorraad S* (stuks)', fontsize=11)
            _ax_bs.set_title(
                f'Totale S* vs. N per service level  '
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
                           label=f'N huidig = {_n_std_sl}')
            _ax_ds.set_xlabel('Aantal subscripties N (midden interval)', fontsize=11)
            _ax_ds.set_ylabel('ΔS* / ΔN  (stuks per extra subscriptie)', fontsize=11)
            _ax_ds.set_title(
                f'Pooling-effect: extra voorraad per extra subscriptie  '
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
                            color='#1976D2', label='CV(N) = 1/√(μₜₒₜ(N))')
                for _xv, _yv in zip(_n_cv_ok, _cv_ok):
                    _ax_cv.annotate(f'{_yv:.3f}', (_xv, _yv),
                                    textcoords='offset points', xytext=(0, 7),
                                    ha='center', fontsize=8)
            _ax_cv.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'N huidig = {_n_std_sl}')
            _ax_cv.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_cv.set_ylabel('CV lead-time demand', fontsize=11)
            _ax_cv.set_title(
                'Relatieve onzekerheid lead-time demand  '
                'CV(N) = 1 / √(N · Σ λᵢ · Lᵢ / Nᵢ)',
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
                           linestyle='--', label=f'μ/N (lead-time demand per sub., ≈{_mu_per_sub:.3f})')
            _ax_sp.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                           label=f'N huidig = {_n_std_sl}')
            _ax_sp.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_sp.set_ylabel('S*(X,N) / N  (voorraad per subscriptie)', fontsize=11)
            _ax_sp.set_title('Benodigde voorraad per subscriptie  S*(X,N) / N',
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
                "Detail: S*(X,N) / N voor N = 1 … 20, "
                "berekend per component via inverse_service_level."
            )
            if st.button("📊 Bereken S*/N voor N = 1 … 20"):
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
                with st.spinner("Berekenen S*/N voor N = 1 … 20…"):
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
                                label=f'μ/N (ondergrens, ≈{_mu_ref:.3f})')
                _ax_det.axvline(_n_std_sl, color='black', linewidth=1.0, linestyle=':',
                                label=f'N huidig = {_n_std_sl}')
                _ax_det.set_xlabel('Aantal subscripties N', fontsize=11)
                _ax_det.set_ylabel('S*(X,N) / N  (voorraad per subscriptie)', fontsize=11)
                _ax_det.set_title(
                    'Benodigde voorraad per subscriptie  S*(X,N) / N  (N = 1… 20)',
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
                           label=f'N huidig = {_n_std_sl}')
            _ax_ss.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_ss.set_ylabel('(S*(X,N) − μ(N)) / N  (safety stock per subscriptie)',
                              fontsize=11)
            _ax_ss.set_title(
                'Safety stock per subscriptie  (S*(X,N) − N · ΣλᵢLᵢ/Nᵢ) / N',
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
        st.subheader("Marginale kosten vs. N")
        st.caption(
            "Hoe nemen de incrementele kosten per extra subscriptie af naarmate N groeit? "
            "Dit visualiseert het pooling-effect: elke extra subscriptie vereist minder "
            "extra inventariskosten dan de vorige. "
            "α, κ_BPA, κ_c en SL worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        _N_MC_VALS   = [1, 2, 3, 5, 7, 10, 15, 20, 30, 50, 75, 100]
        _SL_MC_COLORS = ['#2ca02c', '#ff7f0e', '#1f77b4', '#d62728', '#9467bd', '#8c564b']

        if st.button("📊 Bereken marginale kosten vs. N"):
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
            _n_std_mc = int(cfg['standaard_n_klanten'])
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
                            linestyle=':', label=f'N huidig = {_n_std_mc}')
            _ax_mc1.set_ylabel('Totale BPA-kosten (€)', fontsize=11)
            _ax_mc1.set_title(
                f'BPA-kosten vs. N  '
                f'(α = {_mc_p["alpha"]:.0%}, κ_BPA = {_mc_p["kappa_bpa"]:.0%})',
                fontsize=12,
            )
            _ax_mc1.yaxis.set_major_formatter(_fmt_mc)
            _ax_mc1.set_xticks(_N_MC_VALS)
            _ax_mc1.legend(fontsize=9)
            _ax_mc1.grid(True, alpha=0.3)

            _ax_mc2.axhline(0, color='grey', linewidth=0.8, linestyle='--')
            _ax_mc2.axvline(_n_std_mc, color='black', linewidth=1.0,
                            linestyle=':', label=f'N huidig = {_n_std_mc}')
            _ax_mc2.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_mc2.set_ylabel('ΔC / ΔN  (extra kosten per extra subscriptie, €)',
                               fontsize=11)
            _ax_mc2.set_title(
                'Pooling-effect: marginale inventariskosten per extra subscriptie',
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
            "moet investeren naarmate het klantenbestand groeit."
        )

        _N_INV_VALS = list(range(1, 21))
        _COLORS_INV = ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2']

        if st.button("📊 Bereken investering vs. N"):
            _ov_inv = st.session_state.overzicht_df.reset_index()
            # Verzamel per component: lambda per subscriptie, LT, IP, VP, code
            _comp_inv = []
            for _, _ri in _ov_inv.iterrows():
                _ni = float(_ri.get('n_klanten', 0) or 0)
                _li = float(_ri.get('lambda_jr', 0) or 0)
                _lt = float(_ri.get('LT_dagen', 0) or 0)
                _ip = float(_ri.get('IP', 0) or 0)
                _vp = float(_ri.get('VP', 0) or 0)
                if _ni > 0 and _li > 0 and _lt > 0:
                    _comp_inv.append({
                        'code':        str(_ri.get('Code', '')),
                        'descr':       str(_ri.get('Descr', '')),
                        'lam_per_sub': _li / _ni,
                        'lt_jr':       _lt / 365,
                        'ip':          _ip,
                        'vp':          _vp,
                    })

            _inv_results = {sl: [] for sl in SERVICE_LEVELS}
            # Per-component resultaten voor top-5 grafiek (alle SL's)
            _sl_top = st.session_state.get('kosten_params', {}).get('service_level', 0.990)
            _inv_per_comp = {sl: {c['code']: [] for c in _comp_inv} for sl in SERVICE_LEVELS}

            with st.spinner("Berekenen investering vs. N…"):
                for _n_inv in _N_INV_VALS:
                    for _sl_inv in SERVICE_LEVELS:
                        _totaal = sum(
                            BPAOptimizationModel.inverse_service_level(
                                _sl_inv, _c['lam_per_sub'] * _n_inv, _c['lt_jr']
                            ) * _c['ip']
                            for _c in _comp_inv
                        )
                        _inv_results[_sl_inv].append({'n': _n_inv, 'inv': _totaal})
                    # Per-component per SL (voor top-5 grafiek)
                    for _c in _comp_inv:
                        for _sl_c in SERVICE_LEVELS:
                            _s = BPAOptimizationModel.inverse_service_level(
                                _sl_c, _c['lam_per_sub'] * _n_inv, _c['lt_jr']
                            )
                            _inv_per_comp[_sl_c][_c['code']].append({'n': _n_inv, 'inv': _s * _c['ip']})

            # Top 5 / Top 10 duurste componenten op VP
            _top5_codes  = sorted(_comp_inv, key=lambda c: c['vp'], reverse=True)[:5]
            _top10_codes = sorted(_comp_inv, key=lambda c: c['vp'], reverse=True)[:10]

            st.session_state.sens_inv        = _inv_results
            st.session_state.sens_inv_comp   = _inv_per_comp
            st.session_state.sens_inv_top5   = _top5_codes
            st.session_state.sens_inv_top10  = _top10_codes
            st.session_state.sens_inv_sl_top = _sl_top

        if 'sens_inv' in st.session_state:
            import matplotlib.pyplot as _plt_inv
            import matplotlib.ticker as _mt_inv

            _inv_d    = st.session_state.sens_inv
            _n_std_inv = int(cfg['standaard_n_klanten'])
            _fmt_inv  = _mt_inv.FuncFormatter(lambda v, _: f'€{v:,.0f}')

            _fig_inv, _ax_inv = _plt_inv.subplots(figsize=(11, 5))
            for _sl_inv, _col_inv in zip(SERVICE_LEVELS, _COLORS_INV):
                _pts_inv = [(r['n'], r['inv']) for r in _inv_d[_sl_inv] if r['inv'] is not None]
                if _pts_inv:
                    _xi, _yi = zip(*_pts_inv)
                    _ax_inv.plot(_xi, _yi, marker='o', linewidth=2,
                                 color=_col_inv, label=f'SL = {_sl_inv:.1%}')

            _ax_inv.axvline(_n_std_inv, color='black', linewidth=1.0, linestyle=':',
                            label=f'N huidig = {_n_std_inv}')
            _ax_inv.set_xlabel('Aantal subscripties (N)', fontsize=11)
            _ax_inv.set_ylabel('Totale voorraadwaarde (€)', fontsize=11)
            _ax_inv.set_title(
                'Vereiste investering in basisvoorraad vs. aantal subscripties',
                fontsize=12,
            )
            _ax_inv.yaxis.set_major_formatter(_fmt_inv)
            _ax_inv.set_xticks(_N_INV_VALS)
            _plt_inv.setp(_ax_inv.get_xticklabels(), rotation=30, ha='right')
            _ax_inv.legend(fontsize=9)
            _ax_inv.grid(True, alpha=0.3)
            _fig_inv.tight_layout()
            st.pyplot(_fig_inv)
            _plt_inv.close(_fig_inv)

            # Tabel: investering per N en SL
            _inv_tbl_rows = []
            for _n_v in _N_INV_VALS:
                _row_t = {'N': _n_v}
                for _sl_v in SERVICE_LEVELS:
                    _pts = [r for r in _inv_d[_sl_v] if r['n'] == _n_v]
                    _row_t[f'SL {_sl_v:.1%}'] = f"€{_pts[0]['inv']:,.0f}" if _pts else '—'
                _inv_tbl_rows.append(_row_t)
            st.dataframe(pd.DataFrame(_inv_tbl_rows).set_index('N'), use_container_width=False)

            # ── Top-5 duurste componenten per VP ──────────────────────────
            if 'sens_inv_top5' in st.session_state:
                _top5    = st.session_state.sens_inv_top5
                _comp_d  = st.session_state.sens_inv_comp
                _sl_lbl  = st.session_state.sens_inv_sl_top

                _COLORS_TOP5 = ['#D32F2F', '#F57C00', '#FBC02D', '#388E3C', '#1976D2']
                st.subheader("Top 5 duurste componenten (VP) — investering vs. N (per service level)")
                st.caption(
                    "Gesommeerde investeringswaarde (S\u002a × IP) van de top 5 duurste componenten "
                    "(op verkoopprijs) als functie van N, per service level."
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

                    _ax_t5.axvline(_n_std_inv, color='black', linewidth=1.0, linestyle=':',
                                   label=f'N huidig = {_n_std_inv}')
                    _ax_t5.set_xlabel('Aantal subscripties (N)', fontsize=11)
                    _ax_t5.set_ylabel('Gesommeerde investeringswaarde top 5 (€)', fontsize=11)
                    _ax_t5.set_title(
                        'Top 5 duurste componenten (VP): gesommeerde investering vs. N per SL',
                        fontsize=12,
                    )
                    _ax_t5.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t5.set_xticks(_N_INV_VALS)
                    _plt_inv.setp(_ax_t5.get_xticklabels(), rotation=30, ha='right')
                    _ax_t5.legend(fontsize=9, loc='upper left')
                    _ax_t5.grid(True, alpha=0.3)
                    _fig_t5.tight_layout()
                    st.pyplot(_fig_t5)
                    _plt_inv.close(_fig_t5)

            # ── Top-10 duurste componenten — gesommeerde lijnen per SL ──────
            if 'sens_inv_top10' in st.session_state:
                import matplotlib.cm as _cm_inv
                _top10   = st.session_state.sens_inv_top10
                _comp_d  = st.session_state.sens_inv_comp

                _comp_d_sl0_t10 = _comp_d.get(SERVICE_LEVELS[0], {})
                _x10_sum_vals = [p['n'] for p in _comp_d_sl0_t10.get(_top10[0]['code'], [])] if _top10 else []

                st.subheader("Top 10 duurste componenten (VP) — gesommeerde investering vs. N (per service level)")
                st.caption(
                    "Gesommeerde investeringswaarde (S\u002a × IP) van de top 10 duurste componenten "
                    "(op verkoopprijs) als functie van N, per service level."
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
                    _ax_t10s.axvline(_n_std_inv, color='black', linewidth=1.0, linestyle=':',
                                     label=f'N huidig = {_n_std_inv}')
                    _ax_t10s.set_xlabel('Aantal subscripties (N)', fontsize=11)
                    _ax_t10s.set_ylabel('Gesommeerde investeringswaarde top 10 (€)', fontsize=11)
                    _ax_t10s.set_title(
                        'Top 10 duurste componenten (VP): gesommeerde investering vs. N per SL',
                        fontsize=12,
                    )
                    _ax_t10s.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t10s.set_xticks(_N_INV_VALS)
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

                st.subheader("Top 10 duurste componenten (VP) — investering per component vs. N")
                st.caption(
                    "Investeringswaarde (S\u002a \u00d7 IP) per component als functie van N "
                    "voor het geselecteerde service level."
                )

                _cd10_sl = _comp_d.get(_sl_val_t10, {})
                _x10_vals = [p['n'] for p in _cd10_sl.get(_top10[0]['code'], [])] if _top10 else []

                if _x10_vals:
                    _cmap10  = _cm_inv.get_cmap('tab10')
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

                    _ax_t10.axvline(_n_std_inv, color='black', linewidth=1.0, linestyle=':',
                                    label=f'N huidig = {_n_std_inv}')
                    _ax_t10.set_xlabel('Aantal subscripties (N)', fontsize=11)
                    _ax_t10.set_ylabel('Investeringswaarde per component (€)', fontsize=11)
                    _ax_t10.set_title(
                        f'Top 10 duurste componenten — investering vs. N  ({_sl_sel_t10})',
                        fontsize=12,
                    )
                    _ax_t10.yaxis.set_major_formatter(_fmt_inv)
                    _ax_t10.set_xticks(_N_INV_VALS)
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
            "(α × VP × N). N groeit lineair van start naar doelstelling. "
            "α, κ\\_BPA en SL worden overgenomen uit tabblad 💰 Kostenanalyse."
        )

        _col_mt1, _col_mt2, _col_mt3 = st.columns(3)
        with _col_mt1:
            _N_start_mt = st.number_input(
                "N start (jaar 0)", min_value=1,
                value=int(cfg['standaard_n_klanten']), step=1,
                key="marge_tijd_n_start",
            )
        with _col_mt2:
            _N_end_mt = st.number_input(
                "N eind (doelstelling)", min_value=1,
                value=max(int(cfg['standaard_n_klanten']) * 3, 10), step=1,
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
                        label='Inkomsten',      color='#388E3C', alpha=0.85)
            _ax_mt1.bar(_x_pos, _hold_arr, _w,
                        label='Voorraadkosten', color='#F57C00', alpha=0.85)
            _ax_mt1.bar(_x_pos, _inv_arr,  _w, bottom=_hold_arr,
                        label='Investering',    color='#D32F2F', alpha=0.85)
            _ax_mt1.plot(_x_pos, _net_arr, color='black', marker='o',
                         linewidth=1.8, linestyle='--', label='Netto jaar', zorder=5)
            _ax_mt1.axhline(0, color='grey', linewidth=0.8)

            _ax_mt1b = _ax_mt1.twinx()
            _ax_mt1b.plot(_x_pos, _mtd['N'], color='#1976D2', marker='s',
                          linewidth=1.5, linestyle=':', alpha=0.7, label='N')
            _ax_mt1b.set_ylabel('N (subscripties)', fontsize=10, color='#1976D2')
            _ax_mt1b.tick_params(axis='y', labelcolor='#1976D2')

            _ax_mt1.set_xticks(_x_pos)
            _ax_mt1.set_xticklabels(_x_lbl, rotation=30, ha='right', fontsize=9)
            _ax_mt1.set_ylabel('Cashflow per jaar (€)', fontsize=11)
            _ax_mt1.set_title(
                f'Jaarlijkse cashflow  (α = {_mtp["alpha"]:.0%}, '
                f'κ_BPA = {_mtp["kappa_bpa"]:.0%}, SL = {_mtp["sl"]:.1%}, '
                f'N: {_mtp["N_start"]} → {_mtp["N_end"]})',
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
                         linewidth=2.0, linestyle='-', label='Cumulatieve CF')
            _ax_mt2.axhline(0, color='grey', linewidth=1.0, linestyle='--')

            _be_titel = 'Break-even niet bereikt binnen tijdshorizon'
            for _bi in range(1, len(_cum_arr)):
                if _cum_arr[_bi - 1] < 0 <= _cum_arr[_bi]:
                    _frac     = -_cum_arr[_bi - 1] / (_cum_arr[_bi] - _cum_arr[_bi - 1])
                    _be_x     = _bi - 1 + _frac
                    _be_titel = f'Break-even ≈ jaar {_be_x:.1f}'
                    _ax_mt2.axvline(_be_x, color='#F57C00', linewidth=1.8,
                                    linestyle=':', label=_be_titel)
                    break
            if _cum_arr[0] >= 0:
                _be_titel = 'Break-even in jaar 0 (direct winstgevend)'

            _ax_mt2.set_xticks(_x_pos)
            _ax_mt2.set_xticklabels(_x_lbl, rotation=30, ha='right', fontsize=9)
            _ax_mt2.set_ylabel('Cumulatieve cashflow (€)', fontsize=11)
            _ax_mt2.set_title(f'Cumulatieve cashflow — {_be_titel}', fontsize=12)
            _ax_mt2.yaxis.set_major_formatter(_fmt_mt)
            _ax_mt2.legend(fontsize=9)
            _ax_mt2.grid(True, axis='y', alpha=0.3)

            _fig_mt.tight_layout(h_pad=3)
            st.pyplot(_fig_mt)
            _plt_mt.close(_fig_mt)

            st.dataframe(pd.DataFrame([
                {
                    'Jaar':               _t_arr[_ti],
                    'N':                  _mtd['N'][_ti],
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
            import matplotlib.cm as _mt10_cm
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

            _ax10s1.bar(_xp10, _sr10, _w10, label='Inkomsten',      color='#388E3C', alpha=0.85)
            _ax10s1.bar(_xp10, _sh10, _w10, label='Voorraadkosten', color='#F57C00', alpha=0.85)
            _ax10s1.bar(_xp10, _si10, _w10, bottom=_sh10, label='Investering', color='#D32F2F', alpha=0.85)
            _ax10s1.plot(_xp10, _net10, color='black', marker='o', linewidth=1.8, linestyle='--', label='Netto jaar', zorder=5)
            _ax10s1.axhline(0, color='grey', linewidth=0.8)
            _ax10s1b = _ax10s1.twinx()
            _ax10s1b.plot(_xp10, _mt10d['N'], color='#1976D2', marker='s', linewidth=1.5, linestyle=':', alpha=0.7, label='N')
            _ax10s1b.set_ylabel('N (subscripties)', fontsize=10, color='#1976D2')
            _ax10s1b.tick_params(axis='y', labelcolor='#1976D2')
            _ax10s1.set_xticks(_xp10)
            _ax10s1.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10s1.set_ylabel('Cashflow per jaar (€)', fontsize=11)
            _ax10s1.set_title(
                f'Jaarlijkse cashflow top 10  (α={_mt10p["alpha"]:.0%}, '
                f'κ_BPA={_mt10p["kappa_bpa"]:.0%}, SL={_mt10p["sl"]:.1%}, '
                f'N: {_mt10p["N_start"]} → {_mt10p["N_end"]})', fontsize=12)
            _ax10s1.yaxis.set_major_formatter(_fmt10)
            _h10a, _l10a = _ax10s1.get_legend_handles_labels()
            _h10b, _l10b = _ax10s1b.get_legend_handles_labels()
            _ax10s1.legend(_h10a + _h10b, _l10a + _l10b, fontsize=9, loc='lower right')
            _ax10s1.grid(True, axis='y', alpha=0.3)

            _bc10 = ['#388E3C' if v >= 0 else '#D32F2F' for v in _sc10]
            _ax10s2.bar(_xp10, _sc10, _w10, color=_bc10, alpha=0.65)
            _ax10s2.plot(_xp10, _sc10, color='#1976D2', marker='o', linewidth=2.0, label='Cumulatieve CF')
            _ax10s2.axhline(0, color='grey', linewidth=1.0, linestyle='--')
            _be10t = 'Break-even niet bereikt binnen tijdshorizon'
            for _bi10 in range(1, len(_sc10)):
                if _sc10[_bi10 - 1] < 0 <= _sc10[_bi10]:
                    _f10   = -_sc10[_bi10 - 1] / (_sc10[_bi10] - _sc10[_bi10 - 1])
                    _bex10 = _bi10 - 1 + _f10
                    _be10t = f'Break-even ≈ jaar {_bex10:.1f}'
                    _ax10s2.axvline(_bex10, color='#F57C00', linewidth=1.8, linestyle=':', label=_be10t)
                    break
            if _sc10[0] >= 0:
                _be10t = 'Break-even in jaar 0 (direct winstgevend)'
            _ax10s2.set_xticks(_xp10)
            _ax10s2.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10s2.set_ylabel('Cumulatieve cashflow (€)', fontsize=11)
            _ax10s2.set_title(f'Cumulatieve cashflow top 10 — {_be10t}', fontsize=12)
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
            _cmap10mt  = _mt10_cm.get_cmap('tab10')

            _fig10c, _ax10c = _plt_mt10.subplots(figsize=(12, 5))
            for _ci10, (_code10, _cdata10) in enumerate(_comp_cf10.items()):
                _cum10c = _np_mt10.array(_cdata10['cum'])
                _lbl10c = f"{_code10} – {_cdata10['descr'][:25]}"
                _ax10c.plot(_xp10, _cum10c, color=_cmap10mt(_ci10 / 10),
                            marker='o', linewidth=1.8, markersize=5, label=_lbl10c)
            _ax10c.axhline(0, color='grey', linewidth=1.0, linestyle='--')
            _ax10c.set_xticks(_xp10)
            _ax10c.set_xticklabels(_xl10, rotation=30, ha='right', fontsize=9)
            _ax10c.set_ylabel('Cumulatieve cashflow (€)', fontsize=11)
            _ax10c.set_title(
                f'Cumulatieve cashflow per component — top 10  (SL={_mt10p["sl"]:.1%})',
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
        "Aanname: λ schaalt lineair met N (λ = N × λ_huidig / N_huidig). "
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
                "N huidig":      _n,
                "S* huidig":     _s_now,
                "N voor S*+1":   _n_drempel if _n_drempel is not None else f">{_n + _MAX_N_SEARCH}",
                "Extra N nodig": _extra,
                "λ/jr":          round(_lam, 4),
                "μ = λ·L":       round(float(_row["mu"]), 4),
            })

        _tbl_d = pd.DataFrame(_drempel_rows).set_index("Code")
        _tbl_d_sorted = _tbl_d.sort_values("Extra N nodig", na_position="last")

        # Tabel weergeven met kleurcodering op basis van drempel
        def _kleur_drempel(row):
            v = row["Extra N nodig"]
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
                    "N huidig":      "{:.0f}",
                    "S* huidig":     "{:.0f}",
                    "λ/jr":          "{:.4f}",
                    "μ = λ·L":       "{:.4f}",
                    "Extra N nodig": lambda v: f"{int(v)}" if pd.notna(v) else "—",
                }),
            use_container_width=True,
            height=500,
        )

        # ── Bar chart: Extra N nodig per component ─────────────────────────
        _plot_d = _tbl_d_sorted[_tbl_d_sorted["Extra N nodig"].notna()].copy()
        if not _plot_d.empty:
            import matplotlib.pyplot as _plt_d
            _fig_d, _ax_d = _plt_d.subplots(
                figsize=(max(8, len(_plot_d) * 0.55), 5)
            )
            _ax_d.bar(
                range(len(_plot_d)),
                _plot_d["Extra N nodig"].astype(int),
                color="#1976D2",
            )
            _ax_d.set_xticks(range(len(_plot_d)))
            _ax_d.set_xticklabels(
                _plot_d.index, rotation=45, ha="right", fontsize=9
            )
            _ax_d.set_ylabel("Extra subscripties voor S*+1", fontsize=11)
            _ax_d.set_title(
                f"Subscriptiedrempel per component  (SL = {_sl_d:.1%})",
                fontsize=12,
            )
            _ax_d.grid(True, axis="y", alpha=0.3)
            _fig_d.tight_layout()
            st.pyplot(_fig_d)
            _plt_d.close(_fig_d)








