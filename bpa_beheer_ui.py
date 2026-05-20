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

tab_overzicht, tab_subscripties, tab_toevoegen, tab_verwijderen, tab_config, tab_historie, tab_kosten = st.tabs([
    "📊 Overzicht",
    "✏️ Subscripties aanpassen",
    "➕ Component toevoegen",
    "🗑️ Component verwijderen",
    "⚙️ Configuratie",
    "📈 Historiek",
    "💰 Kostenanalyse",
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

        # Tabel
        st.dataframe(
            df.reset_index().style.format({
                "lambda_jr": "{:.4f}",
                "mu":        "{:.4f}",
                **{c: "{:.0f}" for c in sl_cols},
            }),
            use_container_width=True,
            height=500,
        )

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
        cfg["standaard_n_klanten"] = int(nieuw_standaard)
        sla_config_op(cfg)
        st.success(f"Standaard aangepast naar {cfg['standaard_n_klanten']}")

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
        sla_config_op(cfg)
        st.success(f"Overrides opgeslagen — {len(n_ov)} N, {len(ip_ov)} IP, {len(lt_ov)} LT.")

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

        # Tabel weergeven
        _tbl_display = _tbl_d_sorted.copy()
        _tbl_display["Extra N nodig"] = _tbl_display["Extra N nodig"].apply(
            lambda v: int(v) if pd.notna(v) else "—"
        )
        st.dataframe(
            _tbl_display.style.format({
                "N huidig":  "{:.0f}",
                "S* huidig": "{:.0f}",
                "λ/jr":      "{:.4f}",
                "μ = λ·L": "{:.4f}",
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


