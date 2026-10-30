import json, re, sys, os, base64
from pathlib import Path
from datetime import date, datetime, timezone
from io import BytesIO
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="EICH | Production Planning", page_icon="🌱", layout="wide", initial_sidebar_state="expanded")

# ---------- Paths and authentication ----------
BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
WORKBOOK = DATA / "current_source.xlsx"
PROGRESS = DATA / "completion_progress.json"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown("""<div class="login-card"><div class="login-brand">🌱 EICH</div>
    <h1>Production Planning</h1><p>Secure operations workspace</p></div>""", unsafe_allow_html=True)
    with st.form("login"):
        password = st.text_input("Password / Passwort", type="password")
        submitted = st.form_submit_button("Sign in / Anmelden", use_container_width=True)
    if submitted:
        if password == os.getenv("APP_PASSWORD", st.secrets.get("APP_PASSWORD", "Eich123")):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password / Falsches Passwort")
    st.stop()

# ---------- Translation: source article descriptions are never translated ----------
D = {
"English": {
"language":"Language","nav":"WORKSPACE","dashboard":"Overview","production":"Production plan","delivery":"Delivery plan","orders":"Orders","stock":"Stock & articles","quality":"Data quality","exports":"Exports","progress":"Progress",
"data_source":"DATA SOURCE","upload":"Upload Excel workbook (.xlsx)","import":"Import / replace workbook","reload":"Reload last imported workbook","year":"Planning year","active":"Active workbook","data_note":"Only the first two workbook sheets are used: Orders first, Stock second.",
"tagline":"PLAN · MONITOR · DELIVER","overview_title":"Production intelligence","overview_sub":"A clear view of demand, stock and fulfilment across your operation.",
"ordered":"Ordered units","required":"Production required","stock_units":"Available stock","shortage":"Articles requiring production","trend":"Delivery demand by month","top":"Highest-demand articles","status":"Fulfilment status","complete":"Complete","partial":"In progress","not_started":"Not started","lines":"Order lines",
"prod_title":"Production planning","prod_sub":"Plan required production by article and calendar week.","del_title":"Delivery planning","del_sub":"Track delivery commitments by article and calendar week.",
"from":"From date","to":"To date","date_error":"The start date must be on or before the end date.","no_rows":"No records match the selected date range and filters.",
"article":"Article","description":"Description","week":"Week","required_qty":"Required quantity","produced":"Produced quantity","delivered":"Delivered quantity","remaining":"Remaining quantity","status":"Status","progress":"Progress tracker","progress_note":"Edit produced or delivered quantities. Progress is saved automatically in the app's data folder.",
"filters":"Article filter","article_filter":"Article filter","all_articles":"All articles","filter_note":"Filter by article and sort the visible results. PDF and Excel exports use the filtered rows.",
"search":"Search values","sort":"Sort by","ascending":"Ascending","descending":"Descending","reset":"Reset filters","rows":"Rows shown",
"download_pdf":"Download filtered PDF","download_excel":"Download filtered Excel","pdf_missing":"PDF export requires reportlab; it is listed in requirements.txt.",
"orders_title":"Order register","orders_sub":"Source order lines with calculated dates and stock allocation.",
"stock_title":"Stock & article intelligence","stock_sub":"Available stock, demand and production shortfall by article.",
"quality_title":"Data quality checks","quality_sub":"Review missing dates, stock matches and source-data anomalies.",
"exports_title":"Export planning data","exports_sub":"Download calculated data for reporting and offline analysis.",
"export_calc":"Download calculated Excel workbook","import_success":"Workbook imported","import_error":"Import failed","no_data_title":"Your planning workspace is ready","no_data":"Import your Excel workbook to calculate production requirements, delivery demand and stock allocation.",
"duration_warning":"No production-duration column was detected in Stock. Production dates currently equal delivery dates. Check the duration column mapping.",
"missing_delivery":"Missing delivery date","missing_duration":"Missing production duration","missing_stock":"Article not found in Stock","bad_qty":"Zero or negative quantity","duplicate":"Repeated article and delivery date","severity":"Severity","issue":"Issue","affected":"Rows affected","details":"Details",
"warning":"Warning","info":"Info","detail_date":"These order lines cannot be scheduled by delivery date.","detail_duration":"Check the production duration field in Stock.","detail_stock":"No matching stock record was found for the article number.","detail_qty":"Check whether the row is a cancellation or correction.","detail_duplicate":"These may be valid separate order lines; review if unexpected.",
"production_weekly":"Weekly production requirement","delivery_weekly":"Weekly delivery requirement","demand":"Demand","customers":"Customers","article_count":"Articles","date":"Date","production_date":"Production date","stock_allocated":"Stock allocated","closing_stock":"Remaining stock","qty":"Order quantity","name":"Article name","download_table":"Download this table as Excel","report":"Filtered report"
},
"Deutsch": {
"language":"Sprache","nav":"ARBEITSBEREICH","dashboard":"Übersicht","production":"Produktionsplan","delivery":"Lieferplan","orders":"Aufträge","stock":"Lager & Artikel","quality":"Datenqualität","exports":"Exporte","progress":"Fortschritt",
"data_source":"DATENQUELLE","upload":"Excel-Arbeitsmappe hochladen (.xlsx)","import":"Arbeitsmappe importieren / ersetzen","reload":"Zuletzt importierte Arbeitsmappe neu laden","year":"Planungsjahr","active":"Aktive Arbeitsmappe","data_note":"Es werden nur die ersten zwei Tabellenblätter verwendet: zuerst Orders, dann Stock.",
"tagline":"PLANEN · ÜBERWACHEN · LIEFERN","overview_title":"Produktionsübersicht","overview_sub":"Übersicht über Bedarf, Lagerbestand und Erfüllung im gesamten Betrieb.",
"ordered":"Bestellte Einheiten","required":"Benötigte Produktion","stock_units":"Verfügbarer Bestand","shortage":"Artikel mit Produktionsbedarf","trend":"Lieferbedarf nach Monat","top":"Artikel mit höchstem Bedarf","status":"Erfüllungsstatus","complete":"Abgeschlossen","partial":"In Bearbeitung","not_started":"Nicht begonnen","lines":"Auftragszeilen",
"prod_title":"Produktionsplanung","prod_sub":"Produktionsbedarf nach Artikel und Kalenderwoche planen.","del_title":"Lieferplanung","del_sub":"Lieferverpflichtungen nach Artikel und Kalenderwoche verfolgen.",
"from":"Von Datum","to":"Bis Datum","date_error":"Das Startdatum muss vor oder am Enddatum liegen.","no_rows":"Keine Einträge für den ausgewählten Zeitraum und die Filter.",
"article":"Artikel","description":"Beschreibung","week":"Kalenderwoche","required_qty":"Benötigte Menge","produced":"Produzierte Menge","delivered":"Gelieferte Menge","remaining":"Restmenge","status":"Status","progress":"Fortschritt","progress_note":"Produzierte oder gelieferte Mengen bearbeiten. Der Fortschritt wird automatisch im Datenordner der App gespeichert.",
"filters":"Artikelfilter","article_filter":"Artikelfilter","all_articles":"Alle Artikel","filter_note":"Nach Artikel filtern und sichtbare Ergebnisse sortieren. PDF- und Excel-Export verwenden die gefilterten Zeilen.",
"search":"Werte durchsuchen","sort":"Sortieren nach","ascending":"Aufsteigend","descending":"Absteigend","reset":"Filter zurücksetzen","rows":"Angezeigte Zeilen",
"download_pdf":"Gefiltertes PDF herunterladen","download_excel":"Gefiltertes Excel herunterladen","pdf_missing":"Für den PDF-Export wird reportlab benötigt; es ist in requirements.txt aufgeführt.",
"orders_title":"Auftragsübersicht","orders_sub":"Auftragszeilen mit berechneten Terminen und Lagerzuordnung.",
"stock_title":"Lager- und Artikelübersicht","stock_sub":"Bestand, Bedarf und Produktionsfehlmenge je Artikel.",
"quality_title":"Datenqualitätsprüfung","quality_sub":"Fehlende Daten, Lagerzuordnung und Auffälligkeiten prüfen.",
"exports_title":"Planungsdaten exportieren","exports_sub":"Berechnete Daten für Berichte und Offline-Analysen herunterladen.",
"export_calc":"Berechnete Excel-Arbeitsmappe herunterladen","import_success":"Arbeitsmappe importiert","import_error":"Import fehlgeschlagen","no_data_title":"Ihr Planungsbereich ist bereit","no_data":"Importieren Sie die Excel-Arbeitsmappe, um Produktionsbedarf, Liefermengen und Lagerzuordnung zu berechnen.",
"duration_warning":"Keine Spalte für Produktionsdauer im Lagerblatt erkannt. Produktionsdaten entsprechen derzeit den Lieferdaten. Bitte die Spaltenzuordnung prüfen.",
"missing_delivery":"Fehlendes Lieferdatum","missing_duration":"Fehlende Produktionsdauer","missing_stock":"Artikel nicht im Lager gefunden","bad_qty":"Null oder negative Menge","duplicate":"Doppelter Artikel und Liefertermin","severity":"Schweregrad","issue":"Problem","affected":"Betroffene Zeilen","details":"Details",
"warning":"Warnung","info":"Info","detail_date":"Diese Auftragszeilen können keinem Liefertermin zugeordnet werden.","detail_duration":"Produktionsdauer-Spalte im Lagerblatt prüfen.","detail_stock":"Kein passender Lagerdatensatz zur Artikelnummer gefunden.","detail_qty":"Prüfen, ob es sich um eine Stornierung oder Korrektur handelt.","detail_duplicate":"Es können gültige separate Auftragszeilen sein; bei Bedarf prüfen.",
"production_weekly":"Wöchentlicher Produktionsbedarf","delivery_weekly":"Wöchentlicher Lieferbedarf","demand":"Bedarf","customers":"Kunden","article_count":"Artikel","date":"Datum","production_date":"Produktionstermin","stock_allocated":"Zugeordneter Bestand","closing_stock":"Restbestand","qty":"Bestellmenge","name":"Artikelname","download_table":"Tabelle als Excel herunterladen","report":"Gefilterter Bericht"
}}
if "language" not in st.session_state:
    st.session_state.language = "English"
def T(key):
    return D[st.session_state.language].get(key, key)

st.markdown("""
<style>
:root{--eich-green:#174638;--eich-green-2:#246a52;--eich-mint:#e8f2ec;--eich-ink:#20352c;--eich-muted:#6d7d74;--eich-line:#dfe8e2;--eich-canvas:#f4f7f5;--eich-white:#ffffff}
.stApp{background:var(--eich-canvas);color:var(--eich-ink);font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
[data-testid="stAppViewContainer"]{background:var(--eich-canvas)}
.block-container{max-width:1680px;padding:1.35rem 2.2rem 3rem}
.stAppDeployButton{display:none}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#12392e 0%,#174638 58%,#205944 100%);border-right:1px solid #d7e2db}
[data-testid="stSidebar"]>div:first-child{padding-top:1.2rem}
[data-testid="stSidebar"] *{color:#f5faf6}
[data-testid="stSidebar"] input,[data-testid="stSidebar"] [data-baseweb="select"] *{color:#20352c!important}
[data-testid="stSidebar"] [data-testid="stRadio"]>label{display:none}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"]{gap:.35rem}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label{border:1px solid transparent;border-radius:10px;padding:.55rem .7rem;background:transparent;transition:all .15s ease}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label:hover{background:#ffffff12;border-color:#ffffff20}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked){background:#ffffff20;border-color:#ffffff35}
[data-testid="stSidebar"] hr{border-color:#ffffff30}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section{background:#f8fbf9;border:1px dashed #9bbca9;border-radius:12px}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section *{color:#20352c!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button{background:var(--eich-green-2)!important;color:#fff!important;border:0!important}
.stApp h1,.stApp h2,.stApp h3{color:var(--eich-green);letter-spacing:-.025em}
.stApp h1{font-weight:780;font-size:2rem}.stApp h2{font-weight:730}.stApp h3{font-weight:700}
.stApp p,.stApp label,.stApp [data-testid="stMarkdownContainer"]{color:inherit}
.stApp [data-testid="stCaptionContainer"]{color:var(--eich-muted)!important}
[data-testid="stMetric"]{background:#fff;border:1px solid var(--eich-line);border-radius:16px;padding:1rem 1.15rem;box-shadow:0 3px 12px #183e2d08;min-height:112px}
[data-testid="stMetricLabel"]{color:var(--eich-muted)!important;font-size:.82rem;font-weight:650}
[data-testid="stMetricValue"]{color:var(--eich-green)!important;font-weight:800;letter-spacing:-.035em}
[data-testid="stMetricDelta"]{font-size:.8rem}
.stApp input,.stApp textarea,.stApp [data-baseweb="select"]>div,.stApp [data-baseweb="input"]>div{background:#fff!important;color:#20352c!important;border-color:#cbd9d0!important;border-radius:9px!important}
.stApp input::placeholder,.stApp textarea::placeholder{color:#7c8c82!important;opacity:1!important}
.stApp [data-testid="stDateInput"] input{font-variant-numeric:tabular-nums}
/* Use the EICH green accent for the selected day in every Streamlit calendar. */
[data-baseweb="calendar"] button[aria-selected="true"],
[data-baseweb="calendar"] [aria-selected="true"],
[data-baseweb="calendar"] [role="gridcell"][aria-selected="true"]{
    background:#246a52!important;color:#ffffff!important;border-color:#246a52!important;
}
[data-baseweb="calendar"] button[aria-selected="true"] *,
[data-baseweb="calendar"] [aria-selected="true"] *{color:#ffffff!important}
[data-baseweb="calendar"] button:hover{background:#e3f0e8!important;color:#174638!important}
[data-baseweb="calendar"] button[aria-selected="true"]:hover{background:#174638!important;color:#ffffff!important}
.stApp [data-baseweb="tag"]{background:#e3f0e8!important;border:1px solid #c4ddce!important;color:#174638!important;border-radius:8px!important}
.stApp [data-baseweb="tag"] *{color:#174638!important}
.js-plotly-plot .plotly .modebar{display:none!important}.dashboard-chart-stack .js-plotly-plot{margin-top:.25rem}.dashboard-chart-stack h2,.dashboard-chart-stack h3{margin-bottom:.35rem}.dashboard-chart-stack [data-testid="stVerticalBlock"]{gap:.55rem}
.stApp [data-testid="stFileUploader"] section{background:#f9fbfa!important;border:1px dashed #91b5a1!important;border-radius:12px!important}
.stApp [data-testid="stFileUploader"] section *{color:#20352c!important}
.stApp [data-testid="stFileUploader"] button{background:#246a52!important;color:#fff!important;border:0!important}
.stApp [data-testid="stAlert"]{border-radius:12px;border:1px solid var(--eich-line)}
.stApp [data-testid="stAlert"] *{color:var(--eich-ink)!important}
.stButton>button,.stDownloadButton>button,.stFormSubmitButton>button{border-radius:10px;font-weight:680;border:1px solid #d0ddd4;min-height:2.55rem;transition:all .15s ease}
.stButton>button[kind="primary"],.stFormSubmitButton>button[kind="primary"]{background:var(--eich-green-2);color:white;border-color:var(--eich-green-2)}
.stButton>button:hover,.stDownloadButton>button:hover{border-color:var(--eich-green-2);transform:translateY(-1px)}
[data-testid="stDataFrame"],[data-testid="stDataEditor"],[data-testid="stTable"]{border:1px solid var(--eich-line);border-radius:13px;overflow:hidden;background:#fff}
[data-testid="stDataFrame"] thead tr th,[data-testid="stTable"] thead tr th,[data-testid="stDataEditor"] thead tr th{background:#eaf2ed!important;color:#214b3a!important;font-weight:750!important;border-bottom:1px solid #d4e2d9!important}
/* Weekly completion grid: keep the existing table styling, but use a red checked state. */
[data-testid="stDataEditor"]{--gdg-accent-color:#c62828!important;--gdg-accent-fg:#ffffff!important;--gdg-accent-light:rgba(198,40,40,.16)!important}
[data-testid="stDataFrame"] [role="grid"]{border-radius:12px}
[data-testid="stExpander"]{background:#fff;border:1px solid var(--eich-line);border-radius:13px}
[data-testid="stTabs"] button{font-weight:650}
[data-testid="stProgressBar"]>div>div{background:var(--eich-green-2)}
.hero{background:linear-gradient(115deg,#12392e 0%,#174638 58%,#2c7559 100%);padding:1.35rem 1.65rem;border-radius:18px;color:#fff;margin-bottom:1.4rem;box-shadow:0 8px 26px #12392e16;position:relative;overflow:hidden}
.hero:after{content:"";position:absolute;width:260px;height:260px;right:-75px;top:-155px;border:1px solid #ffffff24;border-radius:50%;box-shadow:0 0 0 28px #ffffff08,0 0 0 58px #ffffff05}
.hero .eyebrow{font-size:.7rem;letter-spacing:.17em;font-weight:800;color:#c4dfce}.hero .title{font-size:1.75rem;font-weight:820;line-height:1.2;margin-top:.35rem;color:#fff;letter-spacing:-.035em}.hero .subtitle{color:#d8e8de;margin-top:.45rem;font-size:.9rem}
.section-title{font-size:1.55rem;font-weight:800;color:#174638;letter-spacing:-.035em}.section-sub{color:#718078;margin:.25rem 0 1.2rem;font-size:.94rem}
.eyebrow{font-size:.72rem;letter-spacing:.13em;font-weight:800;color:#557366;margin:.2rem 0 .75rem}

.exec-card-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;width:100%;max-width:100%;box-sizing:border-box;align-items:stretch}
.exec-card{width:100%;min-width:0;min-height:150px;box-sizing:border-box;padding:18px;border-radius:15px;display:flex;flex-direction:column;justify-content:space-between;overflow:hidden;box-shadow:0 6px 18px #102a2012}
.exec-card-label{font-size:12px;font-weight:700;letter-spacing:.04em;line-height:1.35;overflow-wrap:anywhere}
.exec-card-value{font-size:clamp(21px,2vw,27px);font-weight:800;line-height:1.2;margin-top:12px;overflow-wrap:anywhere;font-variant-numeric:tabular-nums}
.exec-card-foot{font-size:11px;line-height:1.35;margin-top:8px;overflow-wrap:anywhere}
@media(max-width:1100px){.exec-card-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:600px){.exec-card-grid{grid-template-columns:minmax(0,1fr);gap:10px}.exec-card{min-height:128px}}

.login-card{background:linear-gradient(135deg,#12392e,#246a52);padding:2.2rem;border-radius:22px;color:#fff;margin:3rem auto 1rem;max-width:650px;text-align:center;box-shadow:0 18px 48px #12392e25}.login-card h1{color:#fff}.login-card p{color:#d8e8de}.login-brand{font-size:1.1rem;letter-spacing:.15em;font-weight:850;color:#d6e9dd}
[data-testid="stVerticalBlock"]>div:has(> .hero){margin-bottom:.5rem}
@media(max-width:900px){.block-container{padding:.9rem 1rem 2rem}.hero{padding:1.1rem 1.15rem}.hero .title{font-size:1.45rem}.stApp h1{font-size:1.65rem}}
/* EICH premium interaction layer */
@keyframes eichFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes eichGlow{0%,100%{box-shadow:0 8px 26px #12392e16}50%{box-shadow:0 12px 34px #12392e24}}
.block-container{animation:none}
.hero{animation:none}
[data-testid="stMetric"], [data-testid="stDataFrame"], [data-testid="stExpander"]{transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}
[data-testid="stMetric"]:hover{transform:translateY(-2px);box-shadow:0 8px 22px #183e2d12;border-color:#b8d0c1}
.stButton>button:focus-visible,.stDownloadButton>button:focus-visible{outline:3px solid #85b99b;outline-offset:2px}
.stApp [data-testid="stDataFrame"]{box-shadow:0 3px 12px #183e2d05}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label{min-height:42px}
/* Sidebar contrast: readable text on a distinct dark surface, with visible selected states. */
[data-testid="stSidebar"]{background:linear-gradient(180deg,#10263a 0%,#153b43 55%,#174638 100%)!important;border-right:1px solid #2c5960}
[data-testid="stSidebar"] *{color:#f3f7f8}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"]{color:#c5d7df!important}
[data-testid="stSidebar"] .sidebar-section-label{display:block;background:#214d57;color:#ffffff;padding:10px 12px;border:1px solid #3a6b72;border-radius:10px;font-weight:800;letter-spacing:.07em;font-size:.74rem;margin:.35rem 0 .7rem;box-shadow:0 4px 0 #0b202d,0 8px 18px #06152155}
[data-testid="stSidebar"] [data-testid="stRadio"] label{background:#1a3b4b;border:1px solid #315b68;border-radius:10px;padding:8px 10px;margin:4px 0;color:#f5f8fa!important;box-shadow:0 3px 0 #0c2532;transition:transform .15s ease,background .15s ease}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover{background:#285766;transform:translateY(-1px)}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked){background:#e0f2e8;border-color:#b5dfc7;box-shadow:0 3px 0 #8abca0}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) *{color:#12392e!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section{background:#f5faf7!important;border:2px dashed #9dcdb2!important;border-radius:12px}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section *{color:#17392d!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button{background:#287454!important;color:#ffffff!important;border:0!important}
[data-testid="stSidebar"] [data-testid="stSelectbox"] label{color:#f3f7f8!important;font-weight:700}
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] *{color:#17392d!important}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{color:#f5f8fa!important;opacity:1!important}
[data-testid="stSidebar"] button:disabled{opacity:.72!important;color:#e8f1ed!important}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] *{color:#17392d!important}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.01ms!important;animation-iteration-count:1!important;scroll-behavior:auto!important;transition-duration:.01ms!important}}

/* Navigation buttons use the same green as the Reload last imported workbook button. */
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {
    background: #246A52 !important;
    background-color: #246A52 !important;
    border: 1px solid #246A52 !important;
    color: #FFFFFF !important;
    box-shadow: 0 3px 0 #174936 !important;
    border-radius: 10px !important;
    transition: transform .15s ease, background .15s ease, border-color .15s ease;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {
    background: #1D5944 !important;
    background-color: #1D5944 !important;
    border-color: #1D5944 !important;
    transform: translateY(-1px);
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked),
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has([aria-checked="true"]) {
    background: #246A52 !important;
    background-color: #246A52 !important;
    border-color: #246A52 !important;
    box-shadow: 0 3px 0 #174936 !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) *,
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has([aria-checked="true"]) * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label *,
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] input[type="radio"] {
    accent-color: #FFFFFF !important;
}

/* Data Source controls: retain the green/white palette, but make text and
   button treatments consistent and readable. These rules do not affect nav. */
[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
    background: #F7FAF8 !important;
    border: 1px dashed #91B5A1 !important;
    border-radius: 12px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section,
[data-testid="stSidebar"] [data-testid="stFileUploader"] section p,
[data-testid="stSidebar"] [data-testid="stFileUploader"] section span,
[data-testid="stSidebar"] [data-testid="stFileUploader"] section small,
[data-testid="stSidebar"] [data-testid="stFileUploader"] section [data-testid="stFileUploaderDropzoneInstructions"] * {
    color: #20352C !important;
    opacity: 1 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section button {
    background: #246A52 !important;
    color: #FFFFFF !important;
    border: 1px solid #246A52 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section button * {
    color: #FFFFFF !important;
}
/* Import and reload are the two Data Source action buttons. Give both the
   same green treatment so neither label disappears against the sidebar. */
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    background: #246A52 !important;
    color: #FFFFFF !important;
    border: 1px solid #246A52 !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    min-height: 2.55rem !important;
    box-shadow: 0 2px 5px #071d1526 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button:hover:not(:disabled) {
    background: #1D5944 !important;
    border-color: #1D5944 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button:disabled {
    background: #56836F !important;
    border-color: #56836F !important;
    color: #FFFFFF !important;
    opacity: .78 !important;
}
/* White select field with dark, readable value and menu options. */
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border-color: #C7D8CE !important;
    border-radius: 9px !important;
}
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] *,
[data-baseweb="popover"] [role="option"] * {
    color: #20352C !important;
}
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] svg {
    fill: #49665A !important;
}

</style>
""", unsafe_allow_html=True)

# ---------- Navigation ----------
with st.sidebar:
    st.markdown("## 🌱 EICH")
    st.caption("PRODUCTION PLANNING")
    st.selectbox("Language / Sprache", ["English","Deutsch"], key="language")
    st.divider()
    nav = [T("dashboard"),T("production"),T("delivery"),T("progress"),T("stock"),T("exports")]
    page_label = st.radio("Navigation", nav, label_visibility="collapsed", key="nav")
    st.divider()
    uploaded = st.file_uploader(T("upload"), type=["xlsx"], key="xlsx_upload")
    import_click = st.button(T("import"), type="primary", use_container_width=True, disabled=uploaded is None)
    reload_click = st.button(T("reload"), use_container_width=True)

st.markdown(f"""<div class="hero"><div class="eyebrow">EICH · OPERATIONS</div>
<div class="title">🌱 Production Planning</div><div class="subtitle">{T('tagline')}</div></div>""", unsafe_allow_html=True)

# ---------- Workbook parsing ----------
def normalize(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())
def find_column(df, aliases):
    cols = {normalize(c): c for c in df.columns}
    for alias in aliases:
        if normalize(alias) in cols: return cols[normalize(alias)]
    return None
def numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)
def read_workbook(file):
    xls = pd.ExcelFile(file)
    if len(xls.sheet_names) < 2:
        raise ValueError("Excel workbook needs at least two sheets: Orders first, Stock second.")
    return pd.read_excel(xls, sheet_name=0, dtype=object), pd.read_excel(xls, sheet_name=1, dtype=object), xls.sheet_names[:2]
def prepare(orders_raw, stock_raw):
    o, s = orders_raw.copy(), stock_raw.copy()
    oc = {
        "article":find_column(o,["Article Number","Artikelnummer","Artikel Nr.","Artikle Nr.","ArtikelID","Artikel-Nr."]),
        "description":find_column(o,["Name","Bezeichnung","Beschreibung","Description"]),
        "qty":find_column(o,["Order Qty","Menge","Quantity","Qty","Bestellmenge"]),
        "delivery_date":find_column(o,["Delivery Date","Lieferdatum","L-Datum"]),
        "customer":find_column(o,["Customer Name","Kunde","Customer"]),
        "location_order":find_column(o,["Plz-Ort","PLZ-Ort","Ort"]),
    }
    missing = [k for k in ("article","qty","delivery_date") if oc[k] is None]
    if missing: raise ValueError("Missing required Orders columns: "+", ".join(missing)+". Required: Article Number, Order Qty, Delivery Date.")
    o = o.rename(columns={v:k for k,v in oc.items() if v is not None})
    for c in ("description","customer","location_order"):
        if c not in o: o[c] = ""
    o["article"] = o["article"].astype("string").str.strip()
    o["description"] = o["description"].fillna("").astype(str) # keep source descriptions exactly; never translate
    o["qty"] = numeric(o["qty"])
    o["delivery_date"] = pd.to_datetime(o["delivery_date"], errors="coerce", dayfirst=True)
    o = o[o["article"].notna() & o["article"].ne("")].copy()
    sc = {
        "article":find_column(s,["Article Number","ArtikelID","Artikelnummer","Artikel Nr.","Artikle Nr.","Artikel-Nr."]),
        "description_stock":find_column(s,["Name","Beschreibung","Bezeichnung","Description"]),
        "lead_days":find_column(s,["Endzeit","Production Duration","Lead Time","Dauer","lead_days","Production Time","Duration","Produktionsdauer","Fertigungszeit","Bearbeitungszeit","Durchlaufzeit","Production Days","Lead Days","Fertigungsdauer","Zeit"]),
        "available_stock":find_column(s,["Stock","Lagerbestand","Bestand","Available Stock","Lager","Quantity in Stock","Menge","Qty"]),
        "location":find_column(s,["StdLager","Lagerort","Storage","Location"]),
        "material_id":find_column(s,["Material ID","MaterialID","Materialnummer","Material Number","Seed Material ID","Saatgutnummer","Rohstoffnummer"]),
        "seeds_per_article":find_column(s,["Seeds per Article","Seeds per Unit","Seed Factor","Seeds","Saatgut je Artikel","Samen je Artikel","Anzahl Samen","Seed Qty per Article","Material Qty per Article","Verbrauch je Artikel","Menge je Artikel"]),
    }
    if sc["article"] is None or sc["available_stock"] is None:
        raise ValueError("Stock sheet needs an article number and a stock quantity column.")
    duration_missing = sc["lead_days"] is None
    s = s.rename(columns={v:k for k,v in sc.items() if v is not None})
    if "description_stock" not in s: s["description_stock"] = ""
    if "lead_days" not in s: s["lead_days"] = 0
    if "material_id" not in s: s["material_id"] = ""
    if "seeds_per_article" not in s: s["seeds_per_article"] = 0
    s["article"] = s["article"].astype("string").str.strip()
    s["available_stock"] = numeric(s["available_stock"])
    s["lead_days"] = numeric(s["lead_days"])
    s["seeds_per_article"] = numeric(s["seeds_per_article"])
    s["material_id"] = s["material_id"].fillna("").astype(str).str.strip()
    s["description_stock"] = s["description_stock"].fillna("").astype(str)
    s = s[s["article"].notna() & s["article"].ne("")].copy()
    s = s.groupby("article",as_index=False).agg(description_stock=("description_stock","first"),lead_days=("lead_days","max"),available_stock=("available_stock","sum"),material_id=("material_id","first"),seeds_per_article=("seeds_per_article","max"))
    o = o.merge(s[["article","lead_days","available_stock","description_stock","material_id","seeds_per_article"]],on="article",how="left")
    o["material_id"] = o["material_id"].fillna("")
    o["seeds_per_article"] = numeric(o["seeds_per_article"])
    o["seeds_required"] = o["qty"].clip(lower=0) * 1.05
    o["description"] = o["description"].where(o["description"].str.strip().ne(""),o["description_stock"].fillna(""))
    o["lead_days"] = numeric(o["lead_days"])
    o["production_date"] = o["delivery_date"] - pd.to_timedelta(o["lead_days"],unit="D")
    for field, source in [("delivery_year","delivery_date"),("delivery_week","delivery_date"),("production_year","production_date"),("production_week","production_date")]:
        iso = o[source].dt.isocalendar()
        o[field] = iso.year.astype("Int64") if field.endswith("year") else iso.week.astype("Int64")
    o["delivery_period"] = o["delivery_date"].dt.strftime("%Y-%m")
    o["production_period"] = o["production_date"].dt.strftime("%Y-%m")
    return o,s,duration_missing

def make_excel(sheets):
    out = BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter",datetime_format="dd-mmm-yyyy") as writer:
        for name,df in sheets.items():
            df.to_excel(writer,sheet_name=name[:31],index=False)
            ws = writer.sheets[name[:31]]
            ws.freeze_panes(1,0)
            if len(df.columns): ws.autofilter(0,0,max(len(df),1),len(df.columns)-1)
            fmt = writer.book.add_format({"bold":True,"font_color":"white","bg_color":"#17634D","text_wrap":True})
            for i,col in enumerate(df.columns):
                ws.write(0,i,col,fmt)
                sizes = [len(str(col))]+[len(str(x)) for x in df[col].head(250)]
                ws.set_column(i,i,min(max(sizes+[10])+2,36))
    return out.getvalue()
def load_progress():
    try:
        val = json.loads(PROGRESS.read_text(encoding="utf-8")) if PROGRESS.exists() else {}
        return val if isinstance(val,dict) else {}
    except Exception: return {}
def save_progress(obj):
    tmp = PROGRESS.with_suffix(".tmp")
    tmp.write_text(json.dumps(obj,indent=2),encoding="utf-8")
    tmp.replace(PROGRESS)
    cloud_save_state(progress=obj)
def pkey(kind,year,article,week): return f"{kind}|{int(year)}|{str(article)}|{int(week):02d}"

# ---------- Column filters, sorting, PDF and filtered exports ----------
def filter_table(frame, key):
    result = frame.copy()
    columns = list(result.columns)
    # Restore the compact single Article filter used in the earlier layout.
    article_col = next((c for c in columns if normalize(c) in {"article", "artikel", "articlenumber", "artikelnummer", "articleid"}), None)
    if article_col is not None:
        c_filter, c_sort, c_reset = st.columns([2.4, 2.4, 1])
        with c_filter:
            vals = sorted(result[article_col].dropna().astype(str).unique().tolist(), key=str.casefold)
            options = [T("all_articles")] + vals
            selected = st.selectbox(T("article_filter"), options, index=0, key=f"{key}_article_filter")
            if selected != T("all_articles"):
                result = result[result[article_col].astype(str) == selected]
        with c_sort:
            sortcol = st.selectbox(T("sort"), ["—"] + list(result.columns), key=f"{key}_sortcol")
        with c_reset:
            st.write("")
            if st.button(T("reset"), key=f"{key}_reset", use_container_width=True):
                for k in list(st.session_state.keys()):
                    if k.startswith(key + "_"):
                        del st.session_state[k]
                st.rerun()
    else:
        c_sort, c_reset = st.columns([3, 1])
        with c_sort:
            sortcol = st.selectbox(T("sort"), ["—"] + list(result.columns), key=f"{key}_sortcol")
        with c_reset:
            st.write("")
            if st.button(T("reset"), key=f"{key}_reset", use_container_width=True):
                for k in list(st.session_state.keys()):
                    if k.startswith(key + "_"):
                        del st.session_state[k]
                st.rerun()
    if sortcol != "—":
        direction = st.radio(T("ascending") + " / " + T("descending"), [T("ascending"), T("descending")], horizontal=True, key=f"{key}_direction")
        result = result.sort_values(sortcol, ascending=direction == T("ascending"), na_position="last", kind="stable")
    st.caption(f"{T('rows')}: {german_number(len(result))} / {german_number(len(frame))}")
    return result

def german_number(value):
    """Format whole quantities with German thousands dots (e.g. 566.702).
    Fractional quantities use a decimal point and no thousands grouping to avoid
    confusing two different meanings of the same separator.
    """
    try:
        number = float(value)
        if pd.isna(number): return ""
        if abs(number - round(number)) < 1e-9:
            return f"{int(round(number)):,}".replace(",", ".")
        return f"{number:.2f}".rstrip("0").rstrip(".")
    except Exception:
        return str(value)

def display_frame(frame):
    out = frame.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        elif pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].map(german_number)
    return out

def pdf_bytes(frame,title,subtitle=""):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4,landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
    out=BytesIO()
    doc=SimpleDocTemplate(out,pagesize=landscape(A4),rightMargin=7*mm,leftMargin=7*mm,topMargin=8*mm,bottomMargin=8*mm)
    styles=getSampleStyleSheet()
    story=[Paragraph(title.replace("&","&amp;"),styles["Title"])]
    if subtitle: story += [Paragraph(subtitle.replace("&","&amp;"),styles["Normal"]),Spacer(1,4*mm)]
    clean=frame.copy().where(pd.notna(frame),"")
    for col in clean.columns:
        if pd.api.types.is_datetime64_any_dtype(clean[col]):
            clean[col]=pd.to_datetime(clean[col],errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
    data=[list(map(str,clean.columns))]
    data += [[str(v)[:160] for v in row] for row in clean.astype(str).values.tolist()]
    n=max(1,len(data[0])); width=(landscape(A4)[0]-14*mm)/n
    tbl=Table(data,colWidths=[width]*n,repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17634D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),("LEADING",(0,0),(-1,-1),8),
        ("GRID",(0,0),(-1,-1),.35,colors.HexColor("#CAD8D0")),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F1F6F3")]),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
    story.append(tbl); doc.build(story); return out.getvalue()
def show_table(frame,key,title,subtitle=""):
    filtered=filter_table(frame,key)
    st.dataframe(display_frame(filtered),use_container_width=True,hide_index=True,height=470)
    c1,c2=st.columns(2)
    safe=re.sub(r"[^A-Za-z0-9]+","_",title).strip("_")
    with c1:
        st.download_button(T("download_excel"),data=make_excel({"Filtered data":filtered}),file_name=f"EICH_{safe}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"{key}_xlsx")
    with c2:
        try:
            pdf=pdf_bytes(filtered,title,subtitle)
            st.download_button(T("download_pdf"),data=pdf,file_name=f"EICH_{safe}.pdf",mime="application/pdf",key=f"{key}_pdf")
        except ImportError:
            st.warning(T("pdf_missing"))
    return filtered

def show_weekly_detail_table(frame, key, title, order_source, date_column, quantity_column, detail_kind, subtitle=""):
    """Filter/sort weekly article rows while keeping completion checkboxes beside each week."""
    st.caption("Each week shows the planned quantity and a completion checkbox beside it. Tick the box when complete.")
    prog = load_progress()
    article_col = T("article")

    # Keep ISO year/week internally for progress keys, but show compact week labels.
    original_week_cols = [c for c in frame.columns if re.fullmatch(r"\d{4}-W\d{2}", str(c))]
    display_names = {w: f"W {int(w[-2:]):02d}" for w in original_week_cols}
    display_to_original = {display_names[w]: w for w in original_week_cols}
    display_frame = frame.rename(columns=display_names).copy()

    # Reuse the same article filtering and sorting controls as Stock & articles.
    filtered = filter_table(display_frame, f"{key}_filters").reset_index(drop=True)
    week_cols = [c for c in filtered.columns if re.fullmatch(r"W \d{2}", str(c))]
    editor = filtered[[c for c in filtered.columns if c not in week_cols]].copy()
    ordered_cols = list(editor.columns)
    config = {}
    for display_week in week_cols:
        original_week = display_to_original[display_week]
        # Use a fixed width for every week/checkbox pair so Streamlit does not
        # auto-size later columns differently based on their contents.
        raw_values = pd.to_numeric(filtered[display_week], errors="coerce").fillna(0).astype(float)
        editor[display_week] = [f"{int(round(value)):^8}" for value in raw_values]
        done_col = f"✓ {display_week}"
        editor[done_col] = [
            bool(prog.get(pkey(detail_kind, int(original_week[:4]), row.get(article_col, ""), int(original_week[-2:])), 0))
            for _, row in filtered.iterrows()
        ]
        ordered_cols.extend([display_week, done_col])
        config[display_week] = st.column_config.TextColumn(display_week, disabled=True, width=84)
        config[done_col] = st.column_config.CheckboxColumn("✓", help=f"Mark {display_week} complete", default=False, width=58)

    editor = editor[ordered_cols]
    edited = st.data_editor(
        editor, key=f"{key}_editor_v2", use_container_width=True, hide_index=True,
        height=470, column_config=config,
        disabled=[c for c in ordered_cols if c not in week_cols and not c.startswith("✓ ")]
    )
    changed = False
    for i, row in edited.iterrows():
        article = str(row.get(article_col, ""))
        for display_week in week_cols:
            original_week = display_to_original[display_week]
            done_col = f"✓ {display_week}"
            k = pkey(detail_kind, int(original_week[:4]), article, int(original_week[-2:]))
            was = bool(prog.get(k, 0))
            now = bool(row.get(done_col, False))
            if now != was:
                source_row = filtered.iloc[i]
                prog[k] = float(source_row[display_week]) if now else 0
                changed = True
    if changed:
        save_progress(prog)
        st.toast("Completion status saved.")
    st.caption("For order details, use the Orders tab and filter by article and date.")

# ---------- Optional persistent cloud storage (for ephemeral hosts such as Render) ----------
def cloud_client():
    url = os.getenv("SUPABASE_URL") or (st.secrets.get("SUPABASE_URL", "") if hasattr(st, "secrets") else "")
    key = os.getenv("SUPABASE_KEY") or (st.secrets.get("SUPABASE_KEY", "") if hasattr(st, "secrets") else "")
    if not url or not key:
        return None
    try:
        from supabase import create_client
        return create_client(url, key)
    except Exception:
        return None

def cloud_load_state():
    client = cloud_client()
    if client is None: return None
    try:
        rows = client.table("production_state").select("*").eq("id", 1).limit(1).execute().data or []
        return rows[0] if rows else None
    except Exception as exc:
        st.sidebar.warning("Cloud storage is configured but could not be read. Check the production_state table schema and permissions.")
        return None

def cloud_save_state(workbook_bytes=None, filename=None, progress=None):
    client = cloud_client()
    if client is None: return False
    try:
        current = cloud_load_state() or {"id": 1}
        payload = {"id": 1}
        # Compatible with a JSONB state column; schema should contain id and state.
        state = current.get("state") if isinstance(current.get("state"), dict) else {}
        state = dict(state)
        if workbook_bytes is not None:
            state["workbook_base64"] = base64.b64encode(workbook_bytes).decode("ascii")
            state["filename"] = filename or "current_source.xlsx"
            state["imported_at"] = datetime.now(timezone.utc).isoformat()
        if progress is not None: state["progress"] = progress
        payload["state"] = state
        client.table("production_state").upsert(payload).execute()
        return True
    except Exception:
        st.sidebar.warning("Cloud save failed. Verify Supabase credentials and that public.production_state has columns id (integer primary key) and state (jsonb).")
        return False

def restore_cloud_files():
    if WORKBOOK.exists(): return
    state_row = cloud_load_state()
    state = (state_row or {}).get("state") or {}
    encoded = state.get("workbook_base64") if isinstance(state, dict) else None
    if encoded:
        try: WORKBOOK.write_bytes(base64.b64decode(encoded))
        except Exception: pass
    saved_progress = state.get("progress") if isinstance(state, dict) else None
    if isinstance(saved_progress, dict) and not PROGRESS.exists():
        try: PROGRESS.write_text(json.dumps(saved_progress), encoding="utf-8")
        except Exception: pass

restore_cloud_files()

# ---------- Load saved workbook / import new workbook ----------
if reload_click:
    st.session_state.pop("data_ready", None)
    st.session_state.pop("calculated_workbook", None)
    st.session_state.pop("export_workbook_bytes", None)
    st.rerun()
if "data_ready" not in st.session_state:
    st.session_state["data_ready"]=False
    if WORKBOOK.exists():
        try:
            ro,rs,names=read_workbook(WORKBOOK); po,ps,dm=prepare(ro,rs)
            st.session_state["data"]={"orders":po,"stock":ps,"filename":WORKBOOK.name,"sheets":names,"duration_missing":dm}
            st.session_state.pop("calculated_workbook", None)
            st.session_state["data_ready"]=True
        except Exception as exc: st.session_state["load_error"]=str(exc)
if import_click and uploaded is not None:
    try:
        ro,rs,names=read_workbook(uploaded); po,ps,dm=prepare(ro,rs)
        uploaded.seek(0); workbook_bytes=uploaded.getvalue(); WORKBOOK.write_bytes(workbook_bytes)
        cloud_saved = cloud_save_state(workbook_bytes=workbook_bytes, filename=uploaded.name)
        st.session_state["data"]={"orders":po,"stock":ps,"filename":uploaded.name,"sheets":names,"duration_missing":dm,"imported_at":datetime.now().strftime("%d-%b-%Y %H:%M") ,"cloud_saved":cloud_saved}
        st.session_state.pop("calculated_workbook", None)
        st.session_state.pop("export_workbook_bytes", None)
        st.session_state["data_ready"]=True; st.session_state.pop("load_error",None)
        st.success(f"{T('import_success')}: {len(po):,} order lines · {len(ps):,} stock articles")
        st.rerun()
    except Exception as exc: st.error(f"{T('import_error')}: {exc}")
if not st.session_state.get("data_ready") or "data" not in st.session_state:
    st.subheader(T("no_data_title")); st.info(T("no_data"))
    if "load_error" in st.session_state: st.error(st.session_state["load_error"])
    st.stop()

data=st.session_state["data"]
if data.get("duration_missing"):
    st.warning(T("duration_warning"))

# ---------- One-time workbook calculations ----------
# Expensive calculations are performed once per imported workbook and reused
# across Streamlit reruns, date changes, and tab navigation.
if "calculated_workbook" not in st.session_state:
    orders = data["orders"].copy()
    stock = data["stock"].copy()

    # FIFO stock allocation, vectorised by article (no Python row-by-row loop).
    orders["production_required"] = (
        pd.to_numeric(orders["qty"], errors="coerce").fillna(0).clip(lower=0).astype("float64")
    )
    for _col in ("stock_at_start", "stock_allocated", "remaining_stock_after_order"):
        orders[_col] = 0.0

    orders = orders.sort_values(
        ["delivery_date", "article"], na_position="last", kind="stable"
    ).copy()

    valid_delivery = orders["delivery_date"].notna()
    alloc_rows = orders.loc[valid_delivery]
    if not alloc_rows.empty:
        alloc_qty = pd.to_numeric(alloc_rows["qty"], errors="coerce").fillna(0).clip(lower=0)
        stock_by_row = (
            alloc_rows["article"].map(stock.set_index("article")["available_stock"])
            .fillna(0).clip(lower=0).astype("float64")
        )
        cumulative_qty = alloc_qty.groupby(alloc_rows["article"], sort=False).cumsum()
        previous_qty = cumulative_qty - alloc_qty
        stock_at_start = (stock_by_row - previous_qty).clip(lower=0)
        allocated = pd.Series(
            np.minimum(stock_at_start.to_numpy(), alloc_qty.to_numpy()),
            index=alloc_rows.index, dtype="float64"
        )
        remaining_after = (stock_by_row - cumulative_qty).clip(lower=0)

        orders.loc[valid_delivery, "stock_at_start"] = stock_at_start
        orders.loc[valid_delivery, "stock_allocated"] = allocated
        orders.loc[valid_delivery, "remaining_stock_after_order"] = remaining_after
        orders.loc[valid_delivery, "production_required"] = (alloc_qty - allocated).clip(lower=0)

    stock_total = stock["available_stock"].sum()
    demand = orders.groupby("article", as_index=False).agg(
        ordered_qty=("qty", "sum"), order_lines=("qty", "size")
    )
    alloc = orders.groupby("article", as_index=False).agg(
        stock_allocated_total=("stock_allocated", "sum"),
        production_required_total=("production_required", "sum")
    )
    article_summary = stock.merge(demand, on="article", how="left").merge(alloc, on="article", how="left")
    for c in ["ordered_qty", "order_lines", "stock_allocated_total", "production_required_total"]:
        article_summary[c] = article_summary[c].fillna(0)
    article_summary["shortage"] = article_summary["production_required_total"]
    unmatched = sorted(set(orders["article"]) - set(stock["article"]))

    production_plan = (
        orders.dropna(subset=["production_date"])
        .groupby(["article", "description", "production_year", "production_week"], dropna=False)["production_required"]
        .sum().reset_index(name="production_qty")
    )
    delivery_plan = (
        orders.dropna(subset=["delivery_date"])
        .groupby(["article", "description", "delivery_year", "delivery_week"], dropna=False)
        .agg(
            delivery_qty=("qty", "sum"),
            stock_allocated=("stock_allocated", "sum"),
            production_required=("production_required", "sum"),
            closing_stock=("remaining_stock_after_order", "last")
        ).reset_index()
    )

    years = sorted(set(
        [int(y) for y in orders["delivery_year"].dropna() if int(y) > 1900]
        + [int(y) for y in orders["production_year"].dropna() if int(y) > 1900]
    ))
    if not years:
        years = [datetime.now().year]
    default_year = 2027 if 2027 in years else years[-1]
    year = default_year

    st.session_state["calculated_workbook"] = {
        "orders": orders,
        "stock": stock,
        "stock_total": stock_total,
        "article_summary": article_summary,
        "unmatched": unmatched,
        "production_plan": production_plan,
        "delivery_plan": delivery_plan,
        "years": years,
        "year": year,
    }

calculated = st.session_state["calculated_workbook"]
orders = calculated["orders"]
stock = calculated["stock"]
stock_total = calculated["stock_total"]
article_summary = calculated["article_summary"]
unmatched = calculated["unmatched"]
production_plan = calculated["production_plan"]
delivery_plan = calculated["delivery_plan"]
years = calculated["years"]
year = calculated["year"]

page_map={T("dashboard"):"dashboard",T("production"):"production",T("delivery"):"delivery",T("progress"):"progress",T("stock"):"stock",T("exports"):"exports"}
page=page_map[page_label]

# ---------- Shared date range (all tabs, including Exports) ----------
# Initialise once per Streamlit session. The same widget keys are used on every
# tab, so changing either date anywhere updates the shared range everywhere.
# Values remain unchanged during reruns and tab navigation; a fresh app session
# starts again with today's date.
if "shared_from_date" not in st.session_state:
    st.session_state["shared_from_date"] = date.today()
if "shared_to_date" not in st.session_state:
    st.session_state["shared_to_date"] = date.today()

st.markdown("<div class='section-title'>Date range</div>", unsafe_allow_html=True)
date_col_a, date_col_b = st.columns(2)
with date_col_a:
    shared_start = st.date_input(
        T("from"), key="shared_from_date", format="DD/MM/YYYY"
    )
with date_col_b:
    shared_end = st.date_input(
        T("to"), key="shared_to_date", format="DD/MM/YYYY"
    )
if shared_start > shared_end:
    st.error(T("date_error"))
    st.stop()

# ---------- Overview ----------
if page=="dashboard":
    yorders=orders.dropna(subset=["delivery_date"]).copy()
    yorders=yorders[(yorders["delivery_date"].dt.date>=shared_start)&(yorders["delivery_date"].dt.date<=shared_end)]
    yprod=orders.dropna(subset=["production_date"]).copy()
    yprod=yprod[(yprod["production_date"].dt.date>=shared_start)&(yprod["production_date"].dt.date<=shared_end)]
    prog=load_progress()
    produced_total=0.0; delivered_total=0.0
    for _, rr in yprod.groupby(["article","production_year","production_week"])["production_required"].sum().reset_index().iterrows():
        produced_total += float(prog.get(pkey("production",int(rr.production_year),rr.article,int(rr.production_week)),0))
    for _, rr in yorders.groupby(["article","delivery_year","delivery_week"])["qty"].sum().reset_index().iterrows():
        delivered_total += float(prog.get(pkey("delivery",int(rr.delivery_year),rr.article,int(rr.delivery_week)),0))
    req=float(yprod["production_required"].sum()); ordered=float(yorders["qty"].sum()); outstanding_prod=max(0,req-produced_total); outstanding_del=max(0,ordered-delivered_total)
    st.markdown("<div class='eyebrow'>EXECUTIVE OPERATIONS SNAPSHOT</div>",unsafe_allow_html=True)
    metrics=[
        (T("ordered"),ordered,"#173D63","#D7E9FF"),
        (T("required"),req,"#163F35","#A7E3C5"),
        ("Seeds required",float(yprod["seeds_required"].sum()),"#60461A","#FFE0A0"),
        ("Open production",outstanding_prod,"#4B315F","#E4C7FF")
    ]
    metric_descriptions = {
        T("ordered"): "Total customer demand in the selected date range",
        T("required"): "Units that need to be produced in the selected date range",
        "Seeds required": "Total seed quantity needed for the planned production",
        "Open production": "Production quantity still outstanding",
    }
    metric_cards = []
    for label,value,bg,accent in metrics:
        description = metric_descriptions.get(label, "")
        metric_cards.append(
            f"<div class='exec-card' style='background:{bg};color:#fff'>"
            f"<div class='exec-card-label' style='color:{accent}'>{label}</div>"
            f"<div class='exec-card-value' style='color:#fff'>{german_number(value)}</div>"
            f"<div class='exec-card-foot' style='color:{accent}'>{description}</div>"
            f"</div>"
        )
    st.markdown("<div class='exec-card-grid'>" + "".join(metric_cards) + "</div>",unsafe_allow_html=True)
    st.write("")
    # Operational follow-up cards use the same descriptive layout as the
    # executive snapshot cards above.
    follow_up_metrics = [
        (T("stock_units"), stock_total, "#245A48", "#C8E8D5",
         "Total units currently available in stock"),
        ("Produced to date", produced_total, "#234C70", "#D1E5F8",
         "Units recorded as produced within the selected date range"),
        ("Delivered to date", delivered_total, "#315B66", "#D1E9EC",
         "Units recorded as delivered within the selected date range"),
        ("Open deliveries", outstanding_del, "#65452E", "#F3D8BE",
         "Ordered units still awaiting delivery in the selected date range"),
    ]
    follow_up_cards = []
    for label, value, bg, accent, description in follow_up_metrics:
        follow_up_cards.append(
            f"<div class='exec-card' style='background:{bg};color:#fff'>"
            f"<div class='exec-card-label' style='color:{accent}'>{label}</div>"
            f"<div class='exec-card-value' style='color:#fff'>{german_number(value)}</div>"
            f"<div class='exec-card-foot' style='color:{accent}'>{description}</div>"
            f"</div>"
        )
    st.markdown("<div class='exec-card-grid'>" + "".join(follow_up_cards) + "</div>", unsafe_allow_html=True)
    st.divider()
    st.markdown("<div class='dashboard-chart-stack'>", unsafe_allow_html=True)
    with st.container():
        # The charts use the full year selected by the From Date, independently
        # of the From/To filters used by the KPI cards above.
        chart_year = shared_start.year

        # Delivery demand: all ISO weeks belonging to the selected calendar year.
        # Use Delivery Date (not Production Date) and keep zero-demand weeks visible.
        weeks_in_year = date(chart_year, 12, 28).isocalendar().week
        delivery_week_labels = [f"W {week:02d}" for week in range(1, weeks_in_year + 1)]
        delivery_weekly_source = orders.dropna(subset=["delivery_date"]).copy()
        if not delivery_weekly_source.empty:
            delivery_iso = delivery_weekly_source["delivery_date"].dt.isocalendar()
            delivery_weekly_source["iso_year"] = delivery_iso.year.astype(int)
            delivery_weekly_source["iso_week"] = delivery_iso.week.astype(int)
            delivery_weekly_source = delivery_weekly_source[delivery_weekly_source["iso_year"] == chart_year]
            delivery_weekly = (delivery_weekly_source.groupby("iso_week")["qty"].sum()
                               .reindex(range(1, weeks_in_year + 1), fill_value=0)
                               .rename_axis("iso_week").reset_index())
        else:
            delivery_weekly = pd.DataFrame({"iso_week": range(1, weeks_in_year + 1), "qty": 0})
        delivery_weekly["week_label"] = delivery_weekly["iso_week"].map(lambda week: f"W {int(week):02d}")
        st.subheader("Delivery demand by week")
        fig=px.area(delivery_weekly,x="week_label",y="qty",template="plotly_white",markers=True,
                    category_orders={"week_label": delivery_week_labels})
        fig.update_traces(line_color="#17634d",fillcolor="rgba(35,118,94,0.18)",
                          hovertemplate="%{x}<br>Units %{y:,.0f}<extra></extra>")
        fig.update_layout(height=390,margin=dict(l=8,r=10,t=12,b=65),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font_color="#193d32",xaxis_title="Week",yaxis_title="Units",xaxis=dict(type="category",tickmode="array",tickvals=delivery_week_labels,ticktext=delivery_week_labels,tickangle=-45,automargin=True,tickfont=dict(size=9)),yaxis=dict(automargin=True),showlegend=False)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "displaylogo": False, "responsive": True})

        # Production workload: all ISO weeks belonging to the selected year.
        # Reindexing keeps zero-workload weeks visible.
        weeks_in_year = date(chart_year, 12, 28).isocalendar().week
        week_labels = [f"W {week:02d}" for week in range(1, weeks_in_year + 1)]
        weekly_source = orders.dropna(subset=["production_date"]).copy()
        if not weekly_source.empty:
            iso = weekly_source["production_date"].dt.isocalendar()
            weekly_source["iso_year"] = iso.year.astype(int)
            weekly_source["iso_week"] = iso.week.astype(int)
            weekly_source = weekly_source[weekly_source["iso_year"] == chart_year]
            weekly = (weekly_source.groupby("iso_week")["production_required"].sum()
                      .reindex(range(1, weeks_in_year + 1), fill_value=0)
                      .rename_axis("iso_week").reset_index())
        else:
            weekly = pd.DataFrame({"iso_week": range(1, weeks_in_year + 1), "production_required": 0})
        weekly["week_label"] = weekly["iso_week"].map(lambda week: f"W {int(week):02d}")
        st.subheader("Production workload by week")
        fig=px.bar(weekly,x="week_label",y="production_required",category_orders={"week_label":week_labels},template="plotly_white")
        fig.update_traces(marker_color="#2b8068", hovertemplate="%{x}<br>Units %{y:,.0f}<extra></extra>")
        fig.update_layout(height=390,margin=dict(l=8,r=10,t=12,b=65),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font_color="#193d32",xaxis_title="Week",yaxis_title="Units",xaxis=dict(type="category",tickmode="array",tickvals=week_labels,ticktext=week_labels,tickangle=-45,automargin=True,tickfont=dict(size=9)),yaxis=dict(automargin=True),showlegend=False)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "displaylogo": False, "responsive": True})
    st.markdown("</div>", unsafe_allow_html=True)
    st.subheader("Operational watchlist")
    watch=[]
    missing_dates=int(orders["delivery_date"].isna().sum())
    missing_material=int((orders["material_id"].fillna("").astype(str).str.strip()=="").sum())
    missing_seed_factor=int((orders["seeds_per_article"]<=0).sum())
    if missing_dates: watch.append({"Priority":"High","Signal":"Orders without delivery date","Count":missing_dates,"Action":"Review source order dates"})
    if unmatched: watch.append({"Priority":"High","Signal":"Articles without stock match","Count":len(unmatched),"Action":"Check article IDs in Stock sheet"})
    if missing_material: watch.append({"Priority":"Medium","Signal":"Orders without material ID","Count":missing_material,"Action":"Map material IDs in Stock sheet"})
    if missing_seed_factor: watch.append({"Priority":"Medium","Signal":"Articles without seed factor","Count":missing_seed_factor,"Action":"Add seeds-per-article factor to Stock sheet"})
    if watch: st.dataframe(pd.DataFrame(watch),use_container_width=True,hide_index=True)
    else: st.success("No data-quality alerts detected in the current workbook.")

# ---------- Production planning ----------
elif page=="production":
    st.markdown(f"<div class='section-title'>{T('prod_title')}</div><div class='section-sub'>{T('prod_sub')}</div>",unsafe_allow_html=True)
    start, end = shared_start, shared_end
    v=orders.dropna(subset=["production_date"]).copy()
    v=v[(v["production_date"].dt.date>=start)&(v["production_date"].dt.date<=end)]
    if v.empty: st.info(T("no_rows"))
    else:
        weeks=[]
        for _d in pd.date_range(start, end, freq="W-MON"):
            _iso = _d.isocalendar(); _label = f"{int(_iso.year)}-W{int(_iso.week):02d}"
            if _label not in weeks: weeks.append(_label)
        for _d in (pd.Timestamp(start), pd.Timestamp(end)):
            _iso = _d.isocalendar(); _label = f"{int(_iso.year)}-W{int(_iso.week):02d}"
            if _label not in weeks: weeks.append(_label)
        weeks.sort(key=lambda _w: (int(_w[:4]), int(_w[-2:])))
        v["Week"]=v["production_date"].apply(lambda d:f"{int(d.isocalendar().year)}-W{int(d.isocalendar().week):02d}")
        pivot=v.pivot_table(index=["article","description","material_id"],columns="Week",values="production_required",aggfunc="sum",fill_value=0).reset_index(); pivot.columns.name=None
        seed_totals=v.groupby(["article","description","material_id"],dropna=False,as_index=False).agg(seeds_required=("seeds_required","sum"))
        pivot=pivot.merge(seed_totals,on=["article","description","material_id"],how="left")
        for w in weeks:
            if w not in pivot.columns: pivot[w]=0
        pivot=pivot[["article","description","material_id","seeds_required"]+weeks].rename(columns={"article":T("article"),"description":T("description"),"material_id":"Material ID","seeds_required":"Seeds required"})
        st.metric(T("required"),german_number(v['production_required'].sum()))
        show_weekly_detail_table(pivot,"prod_weekly",T("production"),v,"production_date","production_required","production",f"{start:%d-%b-%Y} – {end:%d-%b-%Y}")

# ---------- Delivery planning ----------
elif page=="delivery":
    st.markdown(f"<div class='section-title'>{T('del_title')}</div><div class='section-sub'>{T('del_sub')}</div>",unsafe_allow_html=True)
    start, end = shared_start, shared_end
    v=orders.dropna(subset=["delivery_date"]).copy()
    v=v[(v["delivery_date"].dt.date>=start)&(v["delivery_date"].dt.date<=end)]
    if v.empty: st.info(T("no_rows"))
    else:
        weeks=[]
        for _d in pd.date_range(start, end, freq="W-MON"):
            _iso = _d.isocalendar(); _label = f"{int(_iso.year)}-W{int(_iso.week):02d}"
            if _label not in weeks: weeks.append(_label)
        for _d in (pd.Timestamp(start), pd.Timestamp(end)):
            _iso = _d.isocalendar(); _label = f"{int(_iso.year)}-W{int(_iso.week):02d}"
            if _label not in weeks: weeks.append(_label)
        weeks.sort(key=lambda _w: (int(_w[:4]), int(_w[-2:])))
        v["Week"]=v["delivery_date"].apply(lambda d:f"{int(d.isocalendar().year)}-W{int(d.isocalendar().week):02d}")
        pivot=v.pivot_table(index=["article","description"],columns="Week",values="qty",aggfunc="sum",fill_value=0).reset_index(); pivot.columns.name=None
        for w in weeks:
            if w not in pivot.columns: pivot[w]=0
        pivot=pivot[["article","description"]+weeks].rename(columns={"article":T("article"),"description":T("description")})
        st.metric(T("demand"),german_number(v['qty'].sum()))
        show_weekly_detail_table(pivot,"del_weekly",T("delivery"),v,"delivery_date","qty","delivery",f"{start:%d-%b-%Y} – {end:%d-%b-%Y}")

# ---------- Progress tracker ----------
elif page=="progress":
    st.markdown(f"<div class='section-title'>{T('progress')}</div><div class='section-sub'>{T('progress_note')}</div>",unsafe_allow_html=True)
    all_progress = load_progress()
    prog = {}
    for progress_key, progress_value in all_progress.items():
        parts = progress_key.split("|", 3)
        if len(parts) == 4:
            try:
                week_start = date.fromisocalendar(int(parts[1]), int(parts[3]), 1)
                week_end = week_start + pd.Timedelta(days=6)
                if week_start <= shared_end and week_end >= shared_start:
                    prog[progress_key] = progress_value
            except (ValueError, TypeError):
                continue
    completed = sum(1 for value in prog.values() if float(value or 0) > 0)
    m = st.columns(3)
    m[0].metric("Completed article-weeks", german_number(completed))
    m[1].metric("Tracked article-weeks", german_number(len(prog)))
    m[2].metric("Completion rate", f"{(completed / len(prog) * 100):.1f}%" if prog else "0.0%")
    if prog:
        progress_rows=[]
        for k,value in prog.items():
            parts=k.split("|",3)
            if len(parts)==4:
                kind,yr,article,wk=parts
                progress_rows.append({"Plan":kind.title(),"Year":yr,"Article":article,"Week":f"W{wk}","Completed":bool(float(value or 0)>0),"Completed quantity":float(value or 0)})
        if progress_rows:
            st.dataframe(pd.DataFrame(progress_rows),use_container_width=True,hide_index=True)
        if st.button("Reset all completion marks", type="secondary"):
            save_progress({})
            st.rerun()
    else:
        st.info("No production or delivery weeks have been marked complete yet.")

# ---------- Stock and articles ----------
elif page=="stock":
    st.markdown(f"<div class='section-title'>{T('stock_title')}</div><div class='section-sub'>{T('stock_sub')}</div>",unsafe_allow_html=True)
    stock_period_orders = orders.dropna(subset=["delivery_date"]).copy()
    stock_period_orders = stock_period_orders[
        (stock_period_orders["delivery_date"].dt.date >= shared_start)
        & (stock_period_orders["delivery_date"].dt.date <= shared_end)
    ]
    period_demand = stock_period_orders.groupby("article",as_index=False).agg(
        ordered_qty=("qty","sum"), order_lines=("qty","size")
    )
    period_alloc = stock_period_orders.groupby("article",as_index=False).agg(
        stock_allocated_total=("stock_allocated","sum"),
        production_required_total=("production_required","sum")
    )
    period_summary = stock.merge(period_demand,on="article",how="left").merge(period_alloc,on="article",how="left")
    for c in ["ordered_qty","order_lines","stock_allocated_total","production_required_total"]:
        period_summary[c] = period_summary[c].fillna(0)
    period_summary["shortage"] = period_summary["production_required_total"]
    display=period_summary.rename(columns={"article":T("article"),"description_stock":T("description"),"available_stock":T("stock_units"),"ordered_qty":T("ordered"),"order_lines":T("lines"),"stock_allocated_total":T("stock_allocated"),"production_required_total":T("required"),"shortage":T("shortage")}).sort_values(T("shortage"),ascending=False)
    m=st.columns(3); m[0].metric(T("stock_units"),german_number(period_summary['available_stock'].sum())); m[1].metric(T("ordered"),german_number(period_summary['ordered_qty'].sum())); m[2].metric(T("shortage"),german_number(int((period_summary['shortage']>0).sum())))
    show_table(display,"stock_table",T("stock"))

# ---------- Data quality ----------
elif page=="quality":
    st.markdown(f"<div class='section-title'>{T('quality_title')}</div><div class='section-sub'>{T('quality_sub')}</div>",unsafe_allow_html=True)
    issues=pd.DataFrame([
      {T("severity"):T("warning"),T("issue"):T("missing_delivery"),T("affected"):int(orders["delivery_date"].isna().sum()),T("details"):T("detail_date")},
      {T("severity"):T("warning"),T("issue"):T("missing_duration"),T("affected"):int(orders["lead_days"].isna().sum()),T("details"):T("detail_duration")},
      {T("severity"):T("warning"),T("issue"):T("missing_stock"),T("affected"):int(orders["article"].isin(unmatched).sum()),T("details"):T("detail_stock")},
      {T("severity"):T("info"),T("issue"):T("bad_qty"),T("affected"):int((orders["qty"]<=0).sum()),T("details"):T("detail_qty")},
      {T("severity"):T("info"),T("issue"):T("duplicate"),T("affected"):int(orders.duplicated(["article","delivery_date"],keep=False).sum()),T("details"):T("detail_duplicate")},
    ])
    show_table(issues,"quality_table",T("quality_title"))

# ---------- Exports ----------
elif page=="exports":
    st.markdown(f"<div class='section-title'>{T('exports_title')}</div><div class='section-sub'>{T('exports_sub')}</div>",unsafe_allow_html=True)
    quality=pd.DataFrame([{"Issue":T("missing_delivery"),"Rows":int(orders["delivery_date"].isna().sum())},{"Issue":T("missing_duration"),"Rows":int(orders["lead_days"].isna().sum())},{"Issue":T("missing_stock"),"Rows":len(unmatched)}])
    if "export_workbook_bytes" not in st.session_state:
        st.session_state["export_workbook_bytes"] = make_excel({
            "Orders": orders, "Stock": stock, "Production Plan": production_plan,
            "Delivery Plan": delivery_plan, "Article Summary": article_summary,
            "Import Quality": quality
        })
    st.download_button(T("export_calc"),data=st.session_state["export_workbook_bytes"],file_name=f"EICH_Production_Planning_{year}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.caption(T("data_note"))