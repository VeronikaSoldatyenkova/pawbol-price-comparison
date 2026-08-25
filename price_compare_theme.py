import streamlit as st


def configure_theme():
    st.set_page_config(
        page_title="Price List Comparison",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown(
        """
        <style>
            :root {
                --pc-primary: #0f766e;
                --pc-primary-hover: #115e59;
                --pc-primary-soft: #ecfdf5;
                --pc-ink: #0f172a;
                --pc-muted: #64748b;
                --pc-border: #e2e8f0;
                --pc-surface: #ffffff;
                --pc-soft: #f8fafc;
            }

            .block-container {
                padding-top: 2rem;
                padding-bottom: 3rem;
                max-width: 1550px;
            }

            h1, h2, h3, h4 { color: var(--pc-ink); }

            .pc-hero {
                padding: 1.25rem 1.4rem;
                border: 1px solid var(--pc-border);
                border-radius: 16px;
                background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                margin-bottom: 1rem;
            }

            .pc-hero-title {
                margin: 0;
                font-size: 1.85rem;
                font-weight: 750;
                letter-spacing: -0.02em;
                color: var(--pc-ink);
            }

            .pc-hero-subtitle {
                margin: .35rem 0 0 0;
                color: var(--pc-muted);
                font-size: .98rem;
            }

            .pc-step {
                display: inline-flex;
                align-items: center;
                gap: .55rem;
                margin: .4rem 0 .65rem 0;
                font-size: 1.05rem;
                font-weight: 700;
                color: var(--pc-ink);
            }

            .pc-step-number {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 28px;
                height: 28px;
                border-radius: 999px;
                background: var(--pc-primary-soft);
                color: var(--pc-primary);
                font-size: .86rem;
                font-weight: 800;
            }

            div[data-testid="stMetric"] {
                background: var(--pc-surface);
                border: 1px solid var(--pc-border);
                border-radius: 12px;
                padding: .85rem 1rem;
            }

            div[data-testid="stMetricLabel"] { color: var(--pc-muted); }

            div[data-testid="stExpander"] {
                border: 1px solid var(--pc-border);
                border-radius: 12px;
            }

            .stButton > button[kind="primary"],
            .stDownloadButton > button[kind="primary"],
            button[kind="primary"] {
                background: var(--pc-primary) !important;
                border-color: var(--pc-primary) !important;
                color: white !important;
                box-shadow: none !important;
            }

            .stButton > button[kind="primary"]:hover,
            .stDownloadButton > button[kind="primary"]:hover,
            button[kind="primary"]:hover {
                background: var(--pc-primary-hover) !important;
                border-color: var(--pc-primary-hover) !important;
                color: white !important;
            }

            .stButton > button[kind="secondary"],
            .stDownloadButton > button[kind="secondary"],
            button[kind="secondary"] {
                background: #ffffff !important;
                border-color: #cbd5e1 !important;
                color: #334155 !important;
                box-shadow: none !important;
            }

            button:focus {
                box-shadow: 0 0 0 2px rgba(15, 118, 110, .18) !important;
            }

            .stTextInput input:focus,
            .stTextArea textarea:focus {
                border-color: var(--pc-primary) !important;
                box-shadow: 0 0 0 1px var(--pc-primary) !important;
            }

            div[data-baseweb="select"] > div:focus-within {
                border-color: var(--pc-primary) !important;
            }

            .pc-note {
                border-left: 3px solid var(--pc-primary);
                background: #f0fdfa;
                color: #334155;
                border-radius: 8px;
                padding: .75rem .9rem;
                margin: .35rem 0 .8rem 0;
                font-size: .92rem;
            }

            .pc-muted {
                color: var(--pc-muted);
                font-size: .9rem;
            }

            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def section_title(number: int, title: str):
    st.markdown(
        f"""
        <div class="pc-step">
            <span class="pc-step-number">{number}</span>
            <span>{title}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hero():
    st.markdown(
        """
        <div class="pc-hero">
            <div class="pc-hero-title">📊 Price List Comparison</div>
            <div class="pc-hero-subtitle">
                Compare your current prices with one or many supplier price lists,
                identify the cheapest offers, and export a complete Excel report.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="pc-note">
            Start by choosing the format of <b>Our current price list</b>.
            PowerBI mode keeps the existing fixed structure; Free format lets you map
            identifiers, price, and any extra columns you want to carry into the result.
        </div>
        """,
        unsafe_allow_html=True,
    )
