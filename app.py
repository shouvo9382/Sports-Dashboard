"""
Bangladesh Sports Ministry — National Athlete Performance Dashboard
===================================================================
A demonstration SaaS dashboard for a national sports ministry, built with
Streamlit + Plotly. It reads the athlete workbook directly (read-only), cleans
the human-formatted fields, and provides:

  • Executive Summary with headline KPIs and portfolio-level charts
  • Player Profile cards with cascading Sport -> Player filters
  • Performance & Fitness analytics
  • Injury & Training insights, plus Intake-vs-Burn calorie analysis
  • A configurable, within-sport ranking system + live leaderboard

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

All scoring/parsing logic lives in scoring.py (documented there). Weights are
adjustable from the sidebar and default to the values in scoring.DEFAULT_WEIGHTS.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import scoring

# --------------------------------------------------------------------------- #
#  CONSTANTS + THEME                                                           #
# --------------------------------------------------------------------------- #
APP_DIR = Path(__file__).parent
DATA_PATH = APP_DIR / "data" / "Dashboard_Data.xlsx"
EMBLEM_PATH = APP_DIR / "assets" / "emblem.svg"

NAVY = "#0B4F8A"
NAVY_DARK = "#083b68"
GREEN = "#1B8A4C"
GREEN_LIGHT = "#2ea966"
GOLD = "#F2C14E"
INK = "#12283A"
MUTED = "#5a6b7b"
BG_SOFT = "#F1F6FB"

# Ordered categorical palette used across charts (blue -> green government feel).
SEQ = [NAVY, "#1667a8", GREEN, GREEN_LIGHT, "#5aa9d6", "#7cc79a", GOLD, "#8a6d3b"]

STATUS_COLORS = {
    "Fully Fit": GREEN,
    "Cleared / Recovered": "#5aa9d6",
    "Actively Managed": GOLD,
    "Monitored / Restricted": "#c0563b",
}

st.set_page_config(
    page_title="BD Sports Ministry — Athlete Dashboard",
    page_icon="🏅",
    layout="wide",
    initial_sidebar_state="expanded",
)

CSS = f"""
<style>
:root {{
  --navy:{NAVY}; --green:{GREEN}; --gold:{GOLD}; --ink:{INK}; --muted:{MUTED};
}}
.block-container {{ padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px; }}

/* Header banner */
.gov-header {{
  background: linear-gradient(120deg, {NAVY} 0%, {NAVY_DARK} 55%, {GREEN} 130%);
  border-radius: 16px; padding: 20px 26px; margin-bottom: 18px;
  display:flex; align-items:center; gap:18px; color:#fff;
  box-shadow: 0 8px 22px rgba(11,79,138,.22);
}}
.gov-header h1 {{ font-size: 1.55rem; margin:0; line-height:1.15; color:#fff; }}
.gov-header .sub {{ opacity:.9; font-size:.92rem; margin-top:3px; }}
.gov-header .flag {{ height:6px; width:64px; border-radius:4px;
  background:linear-gradient(90deg,{GREEN} 0 70%, {GOLD} 70% 100%); margin-top:8px; }}

/* KPI tiles */
.kpi {{
  background:#fff; border:1px solid #e4edf5; border-left:5px solid {NAVY};
  border-radius:12px; padding:14px 16px; height:100%;
  box-shadow:0 2px 8px rgba(18,40,58,.05);
}}
.kpi.green {{ border-left-color:{GREEN}; }}
.kpi.gold  {{ border-left-color:{GOLD}; }}
.kpi .label {{ color:{MUTED}; font-size:.78rem; text-transform:uppercase;
  letter-spacing:.04em; font-weight:600; }}
.kpi .value {{ color:{INK}; font-size:1.7rem; font-weight:750; line-height:1.1; }}
.kpi .foot {{ color:{MUTED}; font-size:.78rem; margin-top:2px; }}

/* Profile card */
.profile {{
  background:#fff; border:1px solid #e4edf5; border-radius:16px; padding:20px 22px;
  box-shadow:0 4px 16px rgba(18,40,58,.07);
}}
.profile .avatar {{
  width:74px; height:74px; border-radius:50%;
  background:linear-gradient(135deg,{NAVY},{GREEN}); color:#fff;
  display:flex; align-items:center; justify-content:center;
  font-size:1.7rem; font-weight:750; flex-shrink:0;
}}
.profile .pname {{ font-size:1.35rem; font-weight:750; color:{INK}; line-height:1.1; }}
.profile .prole {{ color:{MUTED}; font-size:.95rem; }}
.chip {{ display:inline-block; padding:3px 11px; border-radius:999px;
  font-size:.78rem; font-weight:650; margin:3px 5px 0 0; }}
.chip.navy {{ background:{BG_SOFT}; color:{NAVY}; border:1px solid #cfe0ef; }}
.chip.green{{ background:#e8f6ee; color:{GREEN}; border:1px solid #bfe6cd; }}
.chip.gold {{ background:#fdf3dc; color:#8a6d1e; border:1px solid #f0dca0; }}

/* Rank badge */
.rankbadge {{
  background:linear-gradient(135deg,{NAVY},{GREEN}); color:#fff;
  border-radius:14px; padding:14px 18px; text-align:center;
}}
.rankbadge .r {{ font-size:2.1rem; font-weight:800; line-height:1; }}
.rankbadge .l {{ font-size:.78rem; opacity:.9; text-transform:uppercase; letter-spacing:.05em; }}

.metric-row {{ display:flex; justify-content:space-between; padding:6px 0;
  border-bottom:1px dashed #e4edf5; font-size:.92rem; }}
.metric-row .k {{ color:{MUTED}; }}
.metric-row .v {{ color:{INK}; font-weight:650; }}

.section-title {{ color:{NAVY}; font-weight:750; font-size:1.15rem;
  border-left:4px solid {GREEN}; padding-left:10px; margin:6px 0 10px; }}

hr {{ margin: 0.8rem 0; }}
[data-testid="stSidebar"] {{ background:{BG_SOFT}; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
#  ACCESS GATE                                                                 #
# --------------------------------------------------------------------------- #
# Protects the dashboard when it is hosted on the public internet (e.g.
# Streamlit Community Cloud), because it displays named athletes with health
# and injury data. The password is read from Streamlit "secrets":
#
#   • Locally  -> .streamlit/secrets.toml  (never commit this file)
#   • On Cloud -> the app's Settings ▸ Secrets panel
#
# If NO password is configured at all, the gate stays OPEN so local
# development and the run_dashboard.bat launcher keep working unchanged.
def _check_access() -> bool:
    try:
        expected = st.secrets.get("app_password", None)
    except Exception:
        expected = None  # no secrets file present -> gate disabled

    if not expected:
        return True  # no password set -> open (local/offline use)

    if st.session_state.get("_authed", False):
        return True

    st.markdown(
        "<div class='gov-header'><div style='font-size:2.4rem'>🏅</div>"
        "<div><h1>National Athlete Performance Dashboard</h1>"
        "<div class='sub'>Ministry of Youth &amp; Sports · Bangladesh</div>"
        "<div class='flag'></div></div></div>",
        unsafe_allow_html=True,
    )
    st.markdown("#### 🔒 Restricted access")
    st.caption("This portal contains athlete health information. "
               "Please enter the access code provided to you.")
    entered = st.text_input("Access code", type="password",
                            label_visibility="collapsed",
                            placeholder="Access code")
    if entered:
        if entered == expected:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Incorrect access code. Please try again.")
    st.stop()  # halt the script here until authenticated


_check_access()


# --------------------------------------------------------------------------- #
#  DATA (cached)                                                               #
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def get_clean_data(path_str: str) -> pd.DataFrame:
    return scoring.load_data(path_str)


@st.cache_data(show_spinner=False)
def get_scored_data(path_str: str, weights_items: tuple,
                    fweights_items: tuple, ideal_ratio: float) -> pd.DataFrame:
    base = get_clean_data(path_str)
    return scoring.compute_scores(
        base,
        weights=dict(weights_items),
        fitness_weights=dict(fweights_items),
        ideal_ratio=ideal_ratio,
    )


def kpi(col, label, value, foot="", cls=""):
    col.markdown(
        f"<div class='kpi {cls}'><div class='label'>{label}</div>"
        f"<div class='value'>{value}</div><div class='foot'>{foot}</div></div>",
        unsafe_allow_html=True,
    )


def initials(name: str) -> str:
    parts = [p for p in str(name).split() if p]
    if not parts:
        return "?"
    return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()


# --------------------------------------------------------------------------- #
#  SIDEBAR — branding, navigation, cascading filters, weights                 #
# --------------------------------------------------------------------------- #
with st.sidebar:
    if EMBLEM_PATH.exists():
        st.image(str(EMBLEM_PATH), width=76)
    st.markdown(
        f"<div style='font-weight:750;color:{NAVY};font-size:1.05rem;line-height:1.2'>"
        "Sports Ministry<br>Athlete Analytics</div>"
        f"<div style='color:{MUTED};font-size:.8rem;margin-bottom:8px'>"
        "National Performance Portal · Demo</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    page = st.radio(
        "Navigate",
        ["Executive Summary", "Player Profile", "Performance & Fitness",
         "Injury & Training", "Rankings & Leaderboard"],
        label_visibility="collapsed",
    )
    st.divider()

    # --- Cascading filters: Sport -> Player ---
    clean = get_clean_data(str(DATA_PATH))
    sports = sorted(clean["Sport"].dropna().unique().tolist())
    sport_choice = st.selectbox("🏆 Sport", ["All Sports"] + sports, index=0)

    if sport_choice == "All Sports":
        player_pool = clean.sort_values("Name")
    else:
        player_pool = clean[clean["Sport"] == sport_choice].sort_values("Name")

    # The player dropdown only ever lists athletes from the chosen sport.
    player_names = player_pool["Name"].tolist()
    player_choice = st.selectbox("👤 Player", player_names, index=0)

    st.caption(
        f"{len(player_pool)} athlete(s) in scope"
        + ("" if sport_choice == "All Sports" else f" · {sport_choice}")
    )
    st.divider()

    # --- Weight controls (drive the ranking model live) ---
    # Keyed sliders + session_state so "Reset" genuinely restores defaults.
    _slider_keys = {
        "performance": "w_perf", "fitness": "w_fit", "health": "w_health",
        "training": "w_train", "nutrition": "w_nutr",
    }
    # Seed defaults once, and honour a pending reset before widgets build.
    for comp, key in _slider_keys.items():
        st.session_state.setdefault(key, scoring.DEFAULT_WEIGHTS[comp])
    if st.session_state.pop("_reset_weights", False):
        for comp, key in _slider_keys.items():
            st.session_state[key] = scoring.DEFAULT_WEIGHTS[comp]

    with st.expander("⚙️ Ranking model weights", expanded=False):
        st.caption("Adjust the emphasis of each component. "
                   "Values are auto-normalised to 100%.")
        w_perf = st.slider("Performance (expert tier)", 0.0, 1.0,
                           step=0.05, key="w_perf")
        w_fit = st.slider("Fitness (VO₂ / HR / body-fat)", 0.0, 1.0,
                          step=0.05, key="w_fit")
        w_health = st.slider("Health / availability", 0.0, 1.0,
                             step=0.05, key="w_health")
        w_train = st.slider("Training load (burn)", 0.0, 1.0,
                            step=0.05, key="w_train")
        w_nutr = st.slider("Nutrition (fuelling)", 0.0, 1.0,
                           step=0.05, key="w_nutr")
        st.button("↺ Reset to defaults",
                  on_click=lambda: st.session_state.update(_reset_weights=True))

weights = {"performance": w_perf, "fitness": w_fit, "health": w_health,
           "training": w_train, "nutrition": w_nutr}
norm_w = scoring._normalise_weights(weights)

df = get_scored_data(
    str(DATA_PATH),
    tuple(sorted(weights.items())),
    tuple(sorted(scoring.DEFAULT_FITNESS_WEIGHTS.items())),
    scoring.IDEAL_INTAKE_BURN_RATIO,
)

# Scope used by pages that respect the Sport filter.
scoped = df if sport_choice == "All Sports" else df[df["Sport"] == sport_choice]


# --------------------------------------------------------------------------- #
#  HEADER BANNER                                                               #
# --------------------------------------------------------------------------- #
st.markdown(
    "<div class='gov-header'>"
    "<div style='font-size:2.4rem'>🏅</div>"
    "<div><h1>National Athlete Performance Dashboard</h1>"
    "<div class='sub'>Ministry of Youth &amp; Sports · Bangladesh — "
    "SaaS demonstration</div><div class='flag'></div></div></div>",
    unsafe_allow_html=True,
)


# =========================================================================== #
#  PAGE 1 — EXECUTIVE SUMMARY                                                  #
# =========================================================================== #
def page_summary():
    st.markdown("<div class='section-title'>Portfolio at a glance</div>",
                unsafe_allow_html=True)

    injured = df["injury_status"].isin(["Actively Managed", "Monitored / Restricted"])
    c1, c2, c3, c4, c5 = st.columns(5)
    kpi(c1, "Athletes", f"{len(df):,}", f"across {df['Sport'].nunique()} sports")
    kpi(c2, "Avg VO₂ Max", f"{df['vo2max'].mean():.1f}",
        "mL/kg/min", cls="green")
    kpi(c3, "Avg Body Fat", f"{df['body_fat'].mean():.1f}%",
        "squad mean", cls="green")
    kpi(c4, "On Injury Watch", f"{int(injured.sum())}",
        f"{injured.mean()*100:.0f}% of squad", cls="gold")
    kpi(c5, "Avg Energy Balance",
        f"{df['energy_balance'].mean():+,.0f}", "kcal intake − burn", cls="gold")

    st.write("")
    left, right = st.columns([1.1, 1])

    with left:
        st.markdown("<div class='section-title'>Athletes by sport</div>",
                    unsafe_allow_html=True)
        counts = (df["Sport"].value_counts()
                  .rename_axis("Sport").reset_index(name="Athletes"))
        fig = px.bar(counts, x="Athletes", y="Sport", orientation="h",
                     color="Sport", color_discrete_sequence=SEQ, text="Athletes")
        fig.update_layout(showlegend=False, height=380,
                          margin=dict(l=0, r=10, t=6, b=0),
                          yaxis=dict(categoryorder="total ascending"))
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, width='stretch')

    with right:
        st.markdown("<div class='section-title'>Squad availability</div>",
                    unsafe_allow_html=True)
        status = (df["injury_status"].value_counts()
                  .rename_axis("Status").reset_index(name="Count"))
        fig = px.pie(status, names="Status", values="Count", hole=0.55,
                     color="Status", color_discrete_map=STATUS_COLORS)
        fig.update_layout(height=380, margin=dict(l=0, r=0, t=6, b=0),
                          legend=dict(orientation="h", y=-0.1))
        st.plotly_chart(fig, width='stretch')

    st.markdown("<div class='section-title'>Average fitness profile by sport</div>",
                unsafe_allow_html=True)
    agg = (df.groupby("Sport")
           .agg(VO2=("vo2max", "mean"), BodyFat=("body_fat", "mean"),
                RestingHR=("resting_hr", "mean"), Athletes=("Name", "count"))
           .reset_index())
    fig = px.scatter(agg, x="VO2", y="BodyFat", size="Athletes", color="Sport",
                     color_discrete_sequence=SEQ, text="Sport",
                     labels={"VO2": "Avg VO₂ Max (mL/kg/min)",
                             "BodyFat": "Avg Body Fat (%)"})
    fig.update_traces(textposition="top center")
    fig.update_layout(height=430, showlegend=False,
                      margin=dict(l=0, r=10, t=6, b=0))
    fig.add_annotation(text="↑ leaner &amp; higher aerobic capacity is bottom-right",
                       xref="paper", yref="paper", x=1, y=1.06, showarrow=False,
                       font=dict(size=11, color=MUTED))
    st.plotly_chart(fig, width='stretch')

    st.caption("Body-fat norms differ by discipline (e.g. sprinters run leaner "
               "than throwers), which is why the ranking model compares each "
               "athlete only against peers in the same sport.")


# =========================================================================== #
#  PAGE 2 — PLAYER PROFILE                                                     #
# =========================================================================== #
def page_profile():
    row = df[df["Name"] == player_choice]
    if row.empty:
        st.warning("Select a player from the sidebar.")
        return
    a = row.iloc[0]

    top = st.columns([2.3, 1, 1])
    with top[0]:
        st.markdown(
            f"<div class='profile'><div style='display:flex;gap:16px;align-items:center'>"
            f"<div class='avatar'>{initials(a['Name'])}</div>"
            f"<div><div class='pname'>{a['Name']}</div>"
            f"<div class='prole'>{a.get('Position/Event','')}</div>"
            f"<div><span class='chip navy'>{a['Sport']}</span>"
            f"<span class='chip green'>{a['tier_label']}</span>"
            f"<span class='chip gold'>ID {a['ID']}</span></div></div></div>"
            f"<div style='margin-top:12px;color:{MUTED};font-size:.9rem'>"
            f"<b>Club / Team:</b> {a.get('Team/Club','—')}</div>"
            f"<div style='margin-top:6px;color:{INK};font-size:.9rem'>"
            f"<b>Recent form:</b> {a.get('Recent Stats','—')}</div></div>",
            unsafe_allow_html=True,
        )
    with top[1]:
        st.markdown(
            f"<div class='rankbadge'><div class='l'>Rank in {a['Sport']}</div>"
            f"<div class='r'>#{int(a['sport_rank'])}</div>"
            f"<div class='l'>of {int(a['sport_size'])}</div></div>",
            unsafe_allow_html=True,
        )
    with top[2]:
        st.markdown(
            f"<div class='rankbadge' style='background:linear-gradient(135deg,{GREEN},{NAVY})'>"
            f"<div class='l'>Percentile</div>"
            f"<div class='r'>{a['sport_percentile']:.0f}</div>"
            f"<div class='l'>overall score {a['overall_score']:.1f}</div></div>",
            unsafe_allow_html=True,
        )

    st.write("")
    c1, c2 = st.columns([1, 1.15])

    with c1:
        st.markdown("<div class='section-title'>Biometrics &amp; vitals</div>",
                    unsafe_allow_html=True)
        rows = [
            ("Age", f"{a['age']:.0f} yrs"),
            ("Height", f"{a['height_cm']:.0f} cm"),
            ("Weight", f"{a['weight_kg']:.0f} kg"),
            ("VO₂ Max", f"{a['vo2max']:.0f} mL/kg/min"),
            ("Resting HR", f"{a['resting_hr']:.0f} BPM"),
            ("Body Fat", f"{a['body_fat']:.1f}%"),
            ("Daily Intake", f"{a['intake_kcal']:,.0f} kcal"),
            ("Daily Burn", f"{a['burn_kcal']:,.0f} kcal"),
            ("Energy Balance", f"{a['energy_balance']:+,.0f} kcal"),
        ]
        html = "".join(
            f"<div class='metric-row'><span class='k'>{k}</span>"
            f"<span class='v'>{v}</span></div>" for k, v in rows
        )
        status_color = STATUS_COLORS.get(a["injury_status"], MUTED)
        html += (f"<div class='metric-row'><span class='k'>Injury status</span>"
                 f"<span class='v' style='color:{status_color}'>"
                 f"● {a['injury_status']}</span></div>")
        st.markdown(f"<div class='profile'>{html}</div>", unsafe_allow_html=True)
        st.caption(f"Injury note: {a.get('Injury History','—')}")

    with c2:
        st.markdown("<div class='section-title'>Component breakdown "
                    "(0–100, vs same-sport peers)</div>", unsafe_allow_html=True)
        comps = ["Performance", "Fitness", "Health", "Training", "Nutrition"]
        vals = [a["perf_n"], a["fitness_n"], a["health_n"],
                a["training_n"], a["nutrition_n"]]
        radar = go.Figure()
        radar.add_trace(go.Scatterpolar(
            r=vals + [vals[0]], theta=comps + [comps[0]], fill="toself",
            line=dict(color=NAVY, width=2), fillcolor="rgba(27,138,76,.28)",
            name=a["Name"]))
        radar.update_layout(
            polar=dict(radialaxis=dict(range=[0, 100], showline=False,
                                       gridcolor="#e4edf5")),
            showlegend=False, height=330, margin=dict(l=30, r=30, t=20, b=20))
        st.plotly_chart(radar, width='stretch')

        # Calorie intake vs burn for this athlete.
        cal = pd.DataFrame({
            "Metric": ["Intake", "Burn"],
            "kcal": [a["intake_kcal"], a["burn_kcal"]],
        })
        fig = px.bar(cal, x="Metric", y="kcal", color="Metric", text="kcal",
                     color_discrete_map={"Intake": NAVY, "Burn": GREEN})
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(showlegend=False, height=250,
                          margin=dict(l=0, r=0, t=24, b=0),
                          title=dict(text="Daily calories: intake vs burn",
                                     font=dict(size=13, color=INK)))
        st.plotly_chart(fig, width='stretch')


# =========================================================================== #
#  PAGE 3 — PERFORMANCE & FITNESS                                             #
# =========================================================================== #
def page_fitness():
    scope_label = "all sports" if sport_choice == "All Sports" else sport_choice
    st.markdown(f"<div class='section-title'>Fitness analytics — {scope_label}</div>",
                unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Avg VO₂ Max", f"{scoped['vo2max'].mean():.1f}", "mL/kg/min")
    kpi(c2, "Avg Resting HR", f"{scoped['resting_hr'].mean():.0f}", "BPM", cls="green")
    kpi(c3, "Avg Body Fat", f"{scoped['body_fat'].mean():.1f}%", "", cls="green")
    kpi(c4, "Fittest (overall)",
        f"{scoped.loc[scoped['fitness_n'].idxmax(), 'Name'].split()[0]}"
        if len(scoped) else "—",
        f"fitness index {scoped['fitness_n'].max():.0f}" if len(scoped) else "",
        cls="gold")

    st.write("")
    l, r = st.columns(2)
    with l:
        st.markdown("<div class='section-title'>VO₂ Max vs Body Fat</div>",
                    unsafe_allow_html=True)
        fig = px.scatter(scoped, x="vo2max", y="body_fat", color="Sport",
                         size="overall_score", hover_name="Name",
                         color_discrete_sequence=SEQ,
                         labels={"vo2max": "VO₂ Max (mL/kg/min)",
                                 "body_fat": "Body Fat (%)"})
        fig.update_layout(height=420, margin=dict(l=0, r=0, t=6, b=0),
                          legend=dict(orientation="h", y=-0.25)
                          if sport_choice == "All Sports" else dict())
        if sport_choice != "All Sports":
            fig.update_layout(showlegend=False)
        st.plotly_chart(fig, width='stretch')

    with r:
        st.markdown("<div class='section-title'>VO₂ Max distribution</div>",
                    unsafe_allow_html=True)
        if sport_choice == "All Sports":
            fig = px.box(df, x="Sport", y="vo2max", color="Sport",
                         color_discrete_sequence=SEQ,
                         labels={"vo2max": "VO₂ Max"})
            fig.update_layout(showlegend=False, height=420,
                              margin=dict(l=0, r=0, t=6, b=0),
                              xaxis=dict(tickangle=-40))
        else:
            fig = px.histogram(scoped, x="vo2max", nbins=12,
                               color_discrete_sequence=[NAVY],
                               labels={"vo2max": "VO₂ Max (mL/kg/min)"})
            fig.update_layout(height=420, margin=dict(l=0, r=0, t=6, b=0),
                              bargap=0.05)
        st.plotly_chart(fig, width='stretch')

    st.markdown("<div class='section-title'>Fitness component index by athlete</div>",
                unsafe_allow_html=True)
    top_n = scoped.nlargest(min(15, len(scoped)), "fitness_n")
    fig = px.bar(top_n.sort_values("fitness_n"), x="fitness_n", y="Name",
                 orientation="h", color="fitness_n",
                 color_continuous_scale=["#cfe0ef", NAVY, GREEN],
                 labels={"fitness_n": "Fitness index (0–100)", "Name": ""})
    fig.update_layout(height=max(300, 26 * len(top_n)),
                      margin=dict(l=0, r=0, t=6, b=0), coloraxis_showscale=False)
    st.plotly_chart(fig, width='stretch')


# =========================================================================== #
#  PAGE 4 — INJURY & TRAINING                                                 #
# =========================================================================== #
def page_injury():
    scope_label = "all sports" if sport_choice == "All Sports" else sport_choice
    st.markdown(f"<div class='section-title'>Injury &amp; training — {scope_label}</div>",
                unsafe_allow_html=True)

    watch = scoped["injury_status"].isin(
        ["Actively Managed", "Monitored / Restricted"])
    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Fully Fit", f"{(scoped['injury_status']=='Fully Fit').sum()}",
        f"{(scoped['injury_status']=='Fully Fit').mean()*100:.0f}% of scope")
    kpi(c2, "On Injury Watch", f"{int(watch.sum())}",
        "managed or restricted", cls="gold")
    kpi(c3, "Avg Training Burn", f"{scoped['burn_kcal'].mean():,.0f}",
        "kcal / day", cls="green")
    kpi(c4, "Avg Availability", f"{scoped['health_score'].mean():.0f}",
        "health index (0–100)", cls="green")

    st.write("")
    l, r = st.columns([1, 1.2])
    with l:
        st.markdown("<div class='section-title'>Availability breakdown</div>",
                    unsafe_allow_html=True)
        status = (scoped["injury_status"].value_counts()
                  .rename_axis("Status").reset_index(name="Count"))
        fig = px.bar(status, x="Count", y="Status", orientation="h",
                     color="Status", color_discrete_map=STATUS_COLORS, text="Count")
        fig.update_layout(showlegend=False, height=320,
                          margin=dict(l=0, r=0, t=6, b=0),
                          yaxis=dict(categoryorder="total ascending"))
        st.plotly_chart(fig, width='stretch')

    with r:
        st.markdown("<div class='section-title'>Training load vs availability</div>",
                    unsafe_allow_html=True)
        fig = px.scatter(scoped, x="burn_kcal", y="health_score",
                         color="injury_status", hover_name="Name",
                         size="intake_kcal", color_discrete_map=STATUS_COLORS,
                         labels={"burn_kcal": "Daily burn (kcal)",
                                 "health_score": "Availability index"})
        fig.update_layout(height=320, margin=dict(l=0, r=0, t=6, b=0),
                          legend=dict(orientation="h", y=-0.3))
        st.plotly_chart(fig, width='stretch')

    # ---- Calorie intake vs burn (headline requirement) ----
    st.markdown("<div class='section-title'>Calorie intake vs burn</div>",
                unsafe_allow_html=True)
    if sport_choice == "All Sports":
        cal = (df.groupby("Sport")
               .agg(Intake=("intake_kcal", "mean"), Burn=("burn_kcal", "mean"))
               .reset_index().sort_values("Intake"))
        m = cal.melt(id_vars="Sport", value_vars=["Intake", "Burn"],
                     var_name="Metric", value_name="kcal")
        fig = px.bar(m, x="Sport", y="kcal", color="Metric", barmode="group",
                     color_discrete_map={"Intake": NAVY, "Burn": GREEN},
                     labels={"kcal": "Avg kcal / day"})
        fig.update_layout(height=420, margin=dict(l=0, r=0, t=6, b=0),
                          xaxis=dict(tickangle=-40),
                          legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig, width='stretch')
        st.caption("Grouped bars compare average daily intake against burn for "
                   "each sport; the gap is the mean energy surplus fuelling "
                   "recovery and adaptation.")
    else:
        d = scoped.sort_values("burn_kcal")
        m = d.melt(id_vars="Name", value_vars=["intake_kcal", "burn_kcal"],
                   var_name="Metric", value_name="kcal")
        m["Metric"] = m["Metric"].map({"intake_kcal": "Intake",
                                       "burn_kcal": "Burn"})
        fig = px.bar(m, x="Name", y="kcal", color="Metric", barmode="group",
                     color_discrete_map={"Intake": NAVY, "Burn": GREEN},
                     labels={"kcal": "kcal / day", "Name": ""})
        fig.update_layout(height=440, margin=dict(l=0, r=0, t=6, b=0),
                          xaxis=dict(tickangle=-55),
                          legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, width='stretch')


# =========================================================================== #
#  PAGE 5 — RANKINGS & LEADERBOARD                                            #
# =========================================================================== #
def page_rankings():
    st.markdown("<div class='section-title'>Within-sport ranking model</div>",
                unsafe_allow_html=True)

    # Live weight readout.
    wc = st.columns(5)
    for col, (name, key) in zip(wc, [
        ("Performance", "performance"), ("Fitness", "fitness"),
        ("Health", "health"), ("Training", "training"),
        ("Nutrition", "nutrition")]):
        kpi(col, name, f"{norm_w[key]*100:.0f}%", "current weight",
            cls="green" if key in ("fitness", "health") else
                ("gold" if key in ("training", "nutrition") else ""))

    with st.expander("📐 How the overall score is calculated", expanded=False):
        st.markdown(f"""
Each athlete gets five components, every one rescaled to **0–100 within their
own sport** (so a swimmer is compared only to swimmers):

| Component | Signal (higher = better) | Weight |
|---|---|---|
| **Performance** | expert Group/Tier grade — *Grade A / Tier 1 = best* | {norm_w['performance']*100:.0f}% |
| **Fitness** | VO₂ Max (+), Resting HR (−), Body Fat (−) | {norm_w['fitness']*100:.0f}% |
| **Health** | availability score from the injury-history text rubric | {norm_w['health']*100:.0f}% |
| **Training** | daily calories burned (training-load proxy) | {norm_w['training']*100:.0f}% |
| **Nutrition** | closeness of the Intake/Burn ratio to the ideal (~{scoring.IDEAL_INTAKE_BURN_RATIO:.2f}) | {norm_w['nutrition']*100:.0f}% |

**Overall = Σ (weight × component)**, then within each sport
**Rank** = position by overall score (1 = best) and
**Percentile** = share of same-sport athletes scoring at or below.

Weights are set in `scoring.DEFAULT_WEIGHTS` and can be changed live from the
sidebar. Fitness sub-weights (VO₂ 50% / HR 25% / body-fat 25%) live in
`scoring.DEFAULT_FITNESS_WEIGHTS`.
""")

    # Leaderboard respects the sidebar Sport filter and updates with weights.
    if sport_choice == "All Sports":
        st.info("Showing the **top athlete of every sport**. "
                "Pick a sport in the sidebar for its full leaderboard.")
        board = (df.sort_values("sport_rank")
                 .groupby("Sport", as_index=False).head(1)
                 .sort_values("overall_score", ascending=False))
        title = "Sport leaders"
    else:
        board = scoped.sort_values("sport_rank")
        title = f"{sport_choice} leaderboard"

    st.markdown(f"<div class='section-title'>{title} "
                f"({len(board)} athletes)</div>", unsafe_allow_html=True)

    show = board.copy()
    show["Rank"] = show["sport_rank"].astype(int)
    disp = show[[
        "Rank", "Name", "Sport", "tier_label", "overall_score",
        "sport_percentile", "perf_n", "fitness_n", "health_n",
        "training_n", "nutrition_n", "injury_status",
    ]].rename(columns={
        "tier_label": "Tier", "overall_score": "Overall",
        "sport_percentile": "Pctl", "perf_n": "Perf", "fitness_n": "Fit",
        "health_n": "Health", "training_n": "Train", "nutrition_n": "Nutr",
        "injury_status": "Status",
    })

    # --- Styling without matplotlib -------------------------------------- #
    # pandas' Styler.background_gradient requires matplotlib, which we do not
    # want as a dependency just for a table shade. This hand-rolled gradient
    # interpolates alpha over the column's own min..max range and uses the
    # dashboard's navy/green palette directly.
    def _gradient(col: pd.Series, rgb: str) -> list[str]:
        lo, hi = col.min(), col.max()
        span = (hi - lo) or 1.0            # single-row / all-equal columns
        out = []
        for v in col:
            t = 0.0 if pd.isna(v) else (float(v) - lo) / span
            out.append(f"background-color: rgba({rgb},{0.08 + 0.55 * t:.2f})")
        return out

    # Highlight the currently-selected player if present in the board.
    def _highlight(r):
        if r["Name"] == player_choice:
            return ["background-color:#fdf3dc" for _ in r]
        return ["" for _ in r]

    comp_cols = ["Perf", "Fit", "Health", "Train", "Nutr"]
    styled = (disp.style
              .format({"Overall": "{:.1f}", "Pctl": "{:.0f}", "Perf": "{:.0f}",
                       "Fit": "{:.0f}", "Health": "{:.0f}", "Train": "{:.0f}",
                       "Nutr": "{:.0f}"})
              # gradients first...
              .apply(_gradient, rgb="27,138,76", subset=["Overall"])
              .apply(_gradient, rgb="11,79,138", subset=comp_cols)
              # ...then the row highlight, so it wins on the selected athlete.
              .apply(_highlight, axis=1))
    st.dataframe(styled, width='stretch', height=min(560, 60 + 35 * len(disp)),
                 hide_index=True)

    st.download_button(
        "⬇️ Download this leaderboard (CSV)",
        disp.to_csv(index=False).encode("utf-8"),
        file_name=f"leaderboard_{sport_choice.replace(' ', '_').lower()}.csv",
        mime="text/csv",
    )

    # Visual: stacked weighted contribution for the visible board (top 12).
    st.markdown("<div class='section-title'>Weighted score contribution</div>",
                unsafe_allow_html=True)
    top = board.nlargest(min(12, len(board)), "overall_score").copy()
    parts = []
    for comp, key, color in [
        ("Performance", "perf_n", NAVY), ("Fitness", "fitness_n", "#1667a8"),
        ("Health", "health_n", GREEN), ("Training", "training_n", GREEN_LIGHT),
        ("Nutrition", "nutrition_n", GOLD)]:
        wkey = {"perf_n": "performance", "fitness_n": "fitness",
                "health_n": "health", "training_n": "training",
                "nutrition_n": "nutrition"}[key]
        parts.append(pd.DataFrame({
            "Name": top["Name"], "Component": comp,
            "Contribution": top[key] * norm_w[wkey], "color": color}))
    contrib = pd.concat(parts)
    order = top.sort_values("overall_score")["Name"].tolist()
    fig = px.bar(contrib, x="Contribution", y="Name", color="Component",
                 orientation="h", color_discrete_map={
                     "Performance": NAVY, "Fitness": "#1667a8", "Health": GREEN,
                     "Training": GREEN_LIGHT, "Nutrition": GOLD},
                 category_orders={"Name": order})
    fig.update_layout(height=max(320, 30 * len(top)),
                      margin=dict(l=0, r=0, t=6, b=0),
                      legend=dict(orientation="h", y=1.08),
                      xaxis_title="Weighted points toward overall score",
                      yaxis_title="")
    st.plotly_chart(fig, width='stretch')


# --------------------------------------------------------------------------- #
#  ROUTER                                                                      #
# --------------------------------------------------------------------------- #
PAGES = {
    "Executive Summary": page_summary,
    "Player Profile": page_profile,
    "Performance & Fitness": page_fitness,
    "Injury & Training": page_injury,
    "Rankings & Leaderboard": page_rankings,
}
PAGES[page]()

st.divider()
st.caption(
    "Demonstration build · Data read directly from the source workbook "
    "(unmodified). Scores are model-derived from the supplied fields; no values "
    "were fabricated. Injury severity is inferred from free-text notes via a "
    "documented keyword rubric in scoring.py."
)
