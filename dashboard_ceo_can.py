# -*- coding: utf-8 -*-
"""
CAN INTERIORS — CEO EXECUTIVE DASHBOARD
Streamlit app · reads Master_Inventory_1.1.xlsx directly.
Run: streamlit run dashboard_ceo_can.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ── Constants ─────────────────────────────────────────────────────────────────
EXCEL_PATH = Path(__file__).parent / 'Master_Inventory_1.1.xlsx'
SHEET      = '📦 MASTER INVENTORY'

BG       = '#FAF8F5'
DARK     = '#3D2B1F'
CLAY     = '#A56F59'
MOCHA    = '#5C463F'
SAGE     = '#8A9E7D'
SAND     = '#C8A882'
CREAM    = '#EFE3D7'
SMOKE    = '#707760'
CHART_PALETTE = [CLAY, MOCHA, SAGE, SAND, '#C79B92', SMOKE, '#D7C2B5', '#2B211D', '#E8C8B8', '#375623']

# ── Page setup ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='CAN Interiors — CEO Dashboard',
    page_icon='🏠',
    layout='wide',
    initial_sidebar_state='collapsed',
)

st.markdown(f"""
<style>
  .stApp {{ background-color: {BG}; }}
  html, body, [class*="css"] {{ color: {DARK}; }}

  /* ── Login card ─────────────────────────────────────── */
  .login-wrap {{
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 3.5rem 0 2rem;
  }}
  .login-logo {{
    font-size: .78rem;
    letter-spacing: .38em;
    color: {CLAY};
    text-transform: uppercase;
    margin-bottom: .35rem;
  }}
  .login-brand {{
    font-size: 1.75rem;
    font-weight: 700;
    color: {MOCHA};
    letter-spacing: .12em;
    text-transform: uppercase;
    margin-bottom: .15rem;
  }}
  .login-sub {{
    font-size: .72rem;
    color: #B5A098;
    letter-spacing: .22em;
    text-transform: uppercase;
    margin-bottom: 2.4rem;
  }}
  .login-divider {{
    width: 38px;
    height: 2px;
    background: {CLAY};
    margin: 0 auto 2.2rem;
    border-radius: 2px;
  }}
  .login-card {{
    background: #FFFFFF;
    border-radius: 18px;
    padding: 2.6rem 2.6rem 2.2rem;
    box-shadow: 0 8px 40px rgba(61,43,31,.10), 0 1px 4px rgba(61,43,31,.06);
    border: 1px solid #EDE5DC;
    max-width: 420px;
    width: 100%;
  }}
  .login-card-title {{
    font-size: .7rem;
    letter-spacing: .24em;
    color: #B5A098;
    text-transform: uppercase;
    text-align: center;
    margin-bottom: 1.6rem;
  }}
  .login-error {{
    background: #FDF0EE;
    border: 1px solid #E8A49A;
    border-radius: 8px;
    padding: .65rem 1rem;
    font-size: .8rem;
    color: #A0442C;
    text-align: center;
    margin-top: .8rem;
  }}
  .login-footer {{
    font-size: .68rem;
    color: #C8BAB3;
    text-align: center;
    margin-top: 2.8rem;
    letter-spacing: .1em;
  }}

  /* ── Dashboard ──────────────────────────────────────── */
  .hero {{
    background: linear-gradient(135deg, {MOCHA} 0%, {CLAY} 100%);
    padding: 2rem 2.5rem 1.8rem;
    border-radius: 16px;
    margin-bottom: 1.8rem;
  }}
  .hero h1 {{ color: #fff; font-size: 1.9rem; margin: 0 0 .3rem; letter-spacing:.06em; }}
  .hero p  {{ color: rgba(255,255,255,.78); font-size: .88rem; margin: 0; }}

  .kpi-card {{
    background: #fff;
    border-radius: 14px;
    padding: 1.3rem 1.5rem 1.1rem;
    box-shadow: 0 2px 14px rgba(61,43,31,.07);
    border-top: 4px solid {CLAY};
  }}
  .kpi-label {{ font-size:.72rem; color:#9A8880; text-transform:uppercase;
                letter-spacing:.12em; margin-bottom:.25rem; }}
  .kpi-value {{ font-size:1.9rem; font-weight:700; color:{MOCHA}; line-height:1.1; }}
  .kpi-sub   {{ font-size:.78rem; color:#B5A098; margin-top:.3rem; }}

  .section-title {{
    font-size:1.05rem; font-weight:700; color:{DARK};
    border-bottom: 2px solid {CLAY};
    padding-bottom:.4rem; margin:1.6rem 0 .9rem;
    letter-spacing:.04em; text-transform:uppercase;
  }}
  div[data-testid="stHorizontalBlock"] > div {{ gap: .8rem; }}
</style>
""", unsafe_allow_html=True)

# ── Credentials ───────────────────────────────────────────────────────────────
_ACCESS_USER = st.secrets.get("ACCESS_USER", "CAN-Interiors")
_ACCESS_KEY  = st.secrets.get("ACCESS_KEY",  "CEO@CAN2026")

# ── Session state ─────────────────────────────────────────────────────────────
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'auth_error' not in st.session_state:
    st.session_state.auth_error = False

# ── Login gate ────────────────────────────────────────────────────────────────
if not st.session_state.authenticated:

    _, center, _ = st.columns([1, 1.35, 1])

    with center:
        st.markdown("""
        <div class="login-wrap">
          <div class="login-logo">est. 2024</div>
          <div class="login-brand">CAN Interiors</div>
          <div class="login-sub">Executive Access Portal</div>
          <div class="login-divider"></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="login-card-title">Secure sign-in required</div>',
                    unsafe_allow_html=True)

        username = st.text_input(
            'Username',
            placeholder='Enter executive username',
            label_visibility='collapsed',
        )
        password = st.text_input(
            'Access key',
            placeholder='Enter access key',
            type='password',
            label_visibility='collapsed',
        )

        st.markdown('<div style="height:.6rem"></div>', unsafe_allow_html=True)

        if st.button('Access Dashboard →', use_container_width=True, type='primary'):
            if username == _ACCESS_USER and password == _ACCESS_KEY:
                st.session_state.authenticated = True
                st.session_state.auth_error    = False
                st.rerun()
            else:
                st.session_state.auth_error = True

        if st.session_state.auth_error:
            st.markdown(
                '<div class="login-error">Invalid credentials. Please try again.</div>',
                unsafe_allow_html=True,
            )

        st.markdown(
            '<div class="login-footer">CAN Interiors &nbsp;·&nbsp; CEO Control Tower &nbsp;·&nbsp; Confidential</div>',
            unsafe_allow_html=True,
        )

    st.stop()

# ── Data ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET, header=1)
    df = df[df['EAN (Primary ID)'].notna()].copy()

    for col in ['Stock', 'Cost Price (€)', 'BOL Price (€)', 'Shop Price (€)',
                'Margin BOL (%)', 'Margin Shop (%)', 'Unieke_Fotos']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Season']   = df['Season'].fillna('All Season').str.strip()
    df['Category'] = df['Category'].fillna('Other').str.strip()

    df['BOL_Rev']    = df['BOL Price (€)']  * df['Stock']
    df['Shop_Rev']   = df['Shop Price (€)'] * df['Stock']
    df['Cost_Total'] = df['Cost Price (€)'] * df['Stock']
    df['BOL_Profit'] = df['BOL_Rev']  - df['Cost_Total']
    df['Shop_Profit']= df['Shop_Rev'] - df['Cost_Total']

    df['Marketing_Klaar'] = df['Unieke_Fotos'].apply(
        lambda x: 'Compleet' if x >= 4 else ('Pending' if x > 0 else 'Geen foto\'s'))

    return df

df = load()

bol_active  = df[df['Marketplace (BOL)'] == 'ACTIVE']
shop_active = df[df['Online Shop']       == 'ACTIVE']

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="hero">
  <h1>🏠 CAN INTERIORS — CEO DASHBOARD</h1>
  <p>Executive overview · {len(df):,} products · Live from Master Inventory</p>
</div>
""", unsafe_allow_html=True)

# ── KPI Row ───────────────────────────────────────────────────────────────────
total_bol_rev   = bol_active['BOL_Rev'].sum()
total_shop_rev  = shop_active['Shop_Rev'].sum()
total_revenue   = total_bol_rev + total_shop_rev
total_cost      = df['Cost_Total'].sum()
net_profit      = total_revenue - total_cost
margin_pct      = (net_profit / total_revenue * 100) if total_revenue else 0
total_stock     = df['Stock'].sum()
avg_margin_bol  = df[df['Margin BOL (%)'] > 0]['Margin BOL (%)'].mean()
marketing_ready = (df['Unieke_Fotos'] >= 4).sum()

k1, k2, k3, k4, k5 = st.columns(5)

def kpi(col, label, value, sub=''):
    col.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

kpi(k1, 'Total Revenue (BOL + Shop)',
    f'€ {total_revenue:,.0f}',
    f'BOL €{total_bol_rev:,.0f} · Shop €{total_shop_rev:,.0f}')

kpi(k2, 'Total Cost',
    f'€ {total_cost:,.0f}',
    f'{total_stock:,} units in stock')

kpi(k3, 'Net Profit',
    f'€ {net_profit:,.0f}',
    f'After cost of goods')

kpi(k4, 'Net Profit Margin',
    f'{margin_pct:.1f} %',
    f'Avg BOL margin {avg_margin_bol:.1f}%')

kpi(k5, 'Marketing Ready',
    f'{marketing_ready:,}',
    f'of {len(df):,} products have ≥4 photos')

st.markdown('<div style="margin-top:1rem"></div>', unsafe_allow_html=True)

# ── Best Sellers & Profitability ──────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Best Sellers & Profitability — Top 10 Margin</div>',
            unsafe_allow_html=True)

col_left, col_right = st.columns([3, 2])

with col_left:
    top10 = (df[df['Margin BOL (%)'] > 0]
             .nlargest(10, 'Margin BOL (%)')
             [['Product Name', 'SKU', 'Margin BOL (%)', 'BOL Price (€)', 'Stock', 'BOL_Rev']]
             .copy())
    top10['Label'] = top10['SKU'] + '  |  ' + top10['Product Name'].str[:38]

    fig_bar = go.Figure(go.Bar(
        x=top10['Margin BOL (%)'],
        y=top10['Label'],
        orientation='h',
        marker_color=CLAY,
        text=top10['Margin BOL (%)'].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        textfont=dict(color=DARK, size=11),
    ))
    fig_bar.update_layout(
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(color=DARK, size=11),
        margin=dict(l=10, r=60, t=10, b=10),
        xaxis=dict(title='Margin BOL (%)', showgrid=True, gridcolor='#EDE5DC',
                   range=[0, top10['Margin BOL (%)'].max() * 1.18]),
        yaxis=dict(autorange='reversed', tickfont=dict(size=10)),
        height=340,
        showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col_right:
    st.markdown('**Top 10 detail**')
    display = top10[['SKU', 'Margin BOL (%)', 'BOL Price (€)', 'Stock', 'BOL_Rev']].copy()
    display.columns = ['SKU', 'Margin %', 'Price €', 'Stock', 'Revenue €']
    display['Revenue €'] = display['Revenue €'].apply(lambda x: f'€{x:,.0f}')
    display['Margin %']  = display['Margin %'].apply(lambda x: f'{x:.1f}%')
    display['Price €']   = display['Price €'].apply(lambda x: f'€{x:.2f}')
    st.dataframe(display.reset_index(drop=True), use_container_width=True, height=320)

# ── Omnichannel Analysis ──────────────────────────────────────────────────────
st.markdown('<div class="section-title">🛒 Omnichannel Analysis — Bol.com vs Web Shop</div>',
            unsafe_allow_html=True)

oc1, oc2, oc3 = st.columns(3)

with oc1:
    cat_bol  = bol_active.groupby('Category')['BOL_Rev'].sum().sort_values(ascending=False).head(8)
    cat_shop = shop_active.groupby('Category')['Shop_Rev'].sum().reindex(cat_bol.index, fill_value=0)

    fig_oc = go.Figure()
    fig_oc.add_trace(go.Bar(name='BOL.com', x=cat_bol.index, y=cat_bol.values,
                            marker_color=CLAY, text=[f'€{v/1000:.0f}k' for v in cat_bol.values],
                            textposition='outside'))
    fig_oc.add_trace(go.Bar(name='Web Shop', x=cat_shop.index, y=cat_shop.values,
                            marker_color=SAGE, text=[f'€{v/1000:.0f}k' for v in cat_shop.values],
                            textposition='outside'))
    fig_oc.update_layout(
        barmode='group', plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(color=DARK, size=10),
        legend=dict(orientation='h', y=1.08, font=dict(size=10)),
        margin=dict(l=0, r=10, t=30, b=60),
        xaxis=dict(tickangle=-30, tickfont=dict(size=9)),
        yaxis=dict(showgrid=True, gridcolor='#EDE5DC', title='Revenue (€)'),
        height=310,
    )
    st.plotly_chart(fig_oc, use_container_width=True)

with oc2:
    st.markdown('**BOL.com Channel**')
    bol_metrics = {
        'Active products': f'{len(bol_active):,}',
        'Not listed':      f'{(df["Marketplace (BOL)"] == "NOT LISTED").sum()}',
        'Total revenue':   f'€ {total_bol_rev:,.0f}',
        'Avg price':       f'€ {bol_active["BOL Price (€)"].mean():.2f}',
        'Avg margin':      f'{bol_active["Margin BOL (%)"].mean():.1f} %',
        'Units in stock':  f'{bol_active["Stock"].sum():,}',
    }
    for k, v in bol_metrics.items():
        c1, c2 = st.columns([2, 1])
        c1.markdown(f'<span style="color:#888;font-size:.82rem">{k}</span>', unsafe_allow_html=True)
        c2.markdown(f'<span style="font-weight:700;color:{MOCHA}">{v}</span>', unsafe_allow_html=True)

    st.markdown('<hr style="margin:.8rem 0">', unsafe_allow_html=True)
    st.markdown('**Web Shop Channel**')
    shop_metrics = {
        'Active products': f'{len(shop_active):,}',
        'Not listed':      f'{(df["Online Shop"] == "NOT LISTED").sum()}',
        'Total revenue':   f'€ {total_shop_rev:,.0f}',
        'Avg price':       f'€ {shop_active["Shop Price (€)"].mean():.2f}',
        'Avg margin':      f'{shop_active["Margin Shop (%)"].mean():.1f} %',
        'Units in stock':  f'{shop_active["Stock"].sum():,}',
    }
    for k, v in shop_metrics.items():
        c1, c2 = st.columns([2, 1])
        c1.markdown(f'<span style="color:#888;font-size:.82rem">{k}</span>', unsafe_allow_html=True)
        c2.markdown(f'<span style="font-weight:700;color:{MOCHA}">{v}</span>', unsafe_allow_html=True)

with oc3:
    cat_stock = df.groupby('Category')['Stock'].sum().sort_values(ascending=False).head(8)
    fig_donut = go.Figure(go.Pie(
        labels=cat_stock.index,
        values=cat_stock.values,
        hole=0.52,
        marker_colors=CHART_PALETTE,
        textinfo='label+percent',
        textfont=dict(size=10, color=DARK),
        hovertemplate='%{label}<br>%{value:,} units<extra></extra>',
    ))
    fig_donut.update_layout(
        plot_bgcolor=BG, paper_bgcolor=BG,
        showlegend=False,
        margin=dict(l=0, r=0, t=10, b=0),
        annotations=[dict(text=f'<b>{total_stock:,}</b><br>units', x=.5, y=.5,
                          font=dict(size=13, color=MOCHA), showarrow=False)],
        height=310,
    )
    fig_donut.update_traces(pull=[0.04] + [0] * 7)
    st.plotly_chart(fig_donut, use_container_width=True)
    st.markdown(f'<center style="font-size:.78rem;color:#999">Stock distribution by category</center>',
                unsafe_allow_html=True)

# ── Marketing & Automation Coverage ──────────────────────────────────────────
st.markdown('<div class="section-title">🎯 Marketing & Automation Coverage</div>',
            unsafe_allow_html=True)

mc1, mc2, mc3 = st.columns(3)

with mc1:
    photo_dist = df['Marketing_Klaar'].value_counts()
    colors_pie = {
        'Compleet':    SAGE,
        'Pending':     SAND,
        "Geen foto's": '#F4A261',
    }
    fig_pie = go.Figure(go.Pie(
        labels=photo_dist.index,
        values=photo_dist.values,
        marker_colors=[colors_pie.get(l, CLAY) for l in photo_dist.index],
        hole=0.5,
        textinfo='label+value',
        textfont=dict(size=11, color=DARK),
        hovertemplate='%{label}: %{value} products (%{percent})<extra></extra>',
    ))
    fig_pie.update_layout(
        plot_bgcolor=BG, paper_bgcolor=BG, showlegend=False,
        margin=dict(l=0, r=0, t=10, b=0),
        annotations=[dict(
            text=f'<b>{marketing_ready}</b><br>ready',
            x=.5, y=.5, font=dict(size=14, color=SAGE), showarrow=False)],
        height=260,
    )
    st.plotly_chart(fig_pie, use_container_width=True)
    st.markdown('<center style="font-size:.78rem;color:#999">Photo sets: ≥4 unique = Compleet</center>',
                unsafe_allow_html=True)

with mc2:
    buckets = df.groupby(df['Unieke_Fotos'].apply(
        lambda x: '0 foto\'s' if x == 0 else '4 foto\'s' if x == 4
                  else '5 foto\'s' if x == 5 else '6+ foto\'s')
    ).size().reindex(["0 foto's", "4 foto's", "5 foto's", "6+ foto's"], fill_value=0)

    bcolors = ['#F4A261', SAGE, CLAY, MOCHA]
    fig_bkt = go.Figure(go.Bar(
        x=buckets.index, y=buckets.values,
        marker_color=bcolors,
        text=buckets.values,
        textposition='outside',
        textfont=dict(color=DARK, size=12),
    ))
    fig_bkt.update_layout(
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(color=DARK, size=11),
        margin=dict(l=0, r=0, t=20, b=0),
        yaxis=dict(showgrid=True, gridcolor='#EDE5DC'),
        xaxis=dict(showgrid=False),
        height=260,
        showlegend=False,
    )
    st.plotly_chart(fig_bkt, use_container_width=True)
    st.markdown('<center style="font-size:.78rem;color:#999">Unique photo count distribution</center>',
                unsafe_allow_html=True)

with mc3:
    season_rev = df.groupby('Season').agg(
        Revenue=('BOL_Rev', 'sum'),
        Products=('EAN (Primary ID)', 'count'),
        Ready=('Unieke_Fotos', lambda x: (x >= 4).sum()),
    ).sort_values('Revenue', ascending=False)

    fig_sea = go.Figure()
    fig_sea.add_trace(go.Bar(
        name='Revenue (€)', x=season_rev.index, y=season_rev['Revenue'],
        marker_color=CLAY, yaxis='y',
        text=[f'€{v/1000:.0f}k' for v in season_rev['Revenue']],
        textposition='outside', textfont=dict(size=9),
    ))
    fig_sea.add_trace(go.Scatter(
        name='Ready %', x=season_rev.index,
        y=(season_rev['Ready'] / season_rev['Products'] * 100).round(1),
        mode='lines+markers', line=dict(color=SAGE, width=2.5),
        marker=dict(size=7), yaxis='y2',
    ))
    fig_sea.update_layout(
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(color=DARK, size=9),
        margin=dict(l=0, r=40, t=20, b=70),
        xaxis=dict(tickangle=-30, tickfont=dict(size=8)),
        yaxis=dict(showgrid=True, gridcolor='#EDE5DC', title=dict(text='Revenue (€)', font=dict(size=9))),
        yaxis2=dict(title=dict(text='Ready %', font=dict(size=9, color=SAGE)),
                    overlaying='y', side='right', range=[0, 110]),
        legend=dict(orientation='h', y=1.12, font=dict(size=9)),
        height=280,
    )
    st.plotly_chart(fig_sea, use_container_width=True)
    st.markdown('<center style="font-size:.78rem;color:#999">Revenue + marketing readiness by season</center>',
                unsafe_allow_html=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown('<hr style="margin:2rem 0 1rem">', unsafe_allow_html=True)
st.markdown(f"""
<div style="text-align:center;color:#B5A098;font-size:.78rem;padding-bottom:1rem">
  CAN Interiors — Executive Dashboard &nbsp;·&nbsp;
  Source: <code>Master_Inventory_1.1.xlsx</code> &nbsp;·&nbsp;
  {len(df):,} products &nbsp;·&nbsp; Powered by Streamlit + Plotly
</div>
""", unsafe_allow_html=True)
