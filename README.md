Below is a complete solution specification for a Streamlit Freight Cost Analytics Dashboard that:

Accepts only Manual Accruals and SAP ERP uploads.
Uses local static master data files:
ERP Carrier Name
ERP Customers
ERP Plants
ERP Shipping Point
Currency conversion table
Standardizes data from both sources.
Converts all values into EUR.
Creates a unified freight spend dataset.
Provides interactive filtering and dashboard analytics.
requirements.txt
streamlit>=1.38.0
pandas>=2.2.2
numpy>=1.26.4
plotly>=5.24.0
openpyxl>=3.1.5
xlrd>=2.0.1

app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path

# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(
    page_title="Freight Cost Analytics",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Freight Cost Analytics Dashboard")
st.markdown(
    """
    Upload:
    - Manual Accruals file
    - SAP ERP file

    Master files are automatically loaded locally.
    """
)

# ==========================================================
# CONFIGURATION
# ==========================================================

MASTER_FOLDER = "master_data"

ERP_CARRIER_FILE = f"{MASTER_FOLDER}/ERP Carrier Name.xlsx"
ERP_CUSTOMERS_FILE = f"{MASTER_FOLDER}/ERP Customers.xlsx"
ERP_PLANTS_FILE = f"{MASTER_FOLDER}/ERP Plants.xlsx"
ERP_SHIPPING_FILE = f"{MASTER_FOLDER}/ERP Shipping Point.xlsx"
CURRENCY_FILE = f"{MASTER_FOLDER}/Currency conversion table.xlsx"

# ==========================================================
# HELPERS
# ==========================================================

@st.cache_data
def load_excel(file):
    return pd.read_excel(file)


def clean_carrier(value):
    if pd.isna(value):
        return value

    value = str(value)

    if "/" in value:
        return value.split("/")[0].strip()

    return value.strip()


def normalize_currency(value):
    if pd.isna(value):
        return None

    return str(value).strip().upper()


# ==========================================================
# LOAD STATIC FILES
# ==========================================================

try:
    carrier_master = load_excel(ERP_CARRIER_FILE)
    customer_master = load_excel(ERP_CUSTOMERS_FILE)
    plant_master = load_excel(ERP_PLANTS_FILE)
    shipping_master = load_excel(ERP_SHIPPING_FILE)
    currency_master = load_excel(CURRENCY_FILE)

except Exception as e:
    st.error(f"Unable to load master files: {e}")
    st.stop()

# ==========================================================
# FILE UPLOADS
# ==========================================================

manual_file = st.file_uploader(
    "Upload Manual Accruals",
    type=["xlsx", "xls"]
)

sap_file = st.file_uploader(
    "Upload SAP ERP",
    type=["xlsx", "xls"]
)

if not manual_file or not sap_file:
    st.stop()

# ==========================================================
# LOAD DATA
# ==========================================================

manual_raw = pd.read_excel(manual_file)
sap_raw = pd.read_excel(sap_file)

# ==========================================================
# MANUAL ACCRUALS TRANSFORMATION
# ==========================================================

manual = pd.DataFrame()

manual["Carrier"] = (
    manual_raw["Carrier Description"]
    .astype(str)
    .apply(clean_carrier)
)

manual["Type of Goods"] = manual_raw["PBU"]

manual["Date"] = pd.to_datetime(
    manual_raw["Planned Arrival Date-Last Stop"],
    errors="coerce"
)

manual["Origin"] = manual_raw["Source Location Description"]

manual["Destination"] = manual_raw["Destination Location Descripti"]

manual["Value"] = pd.to_numeric(
    manual_raw["Net Amt in Doc Crcy"],
    errors="coerce"
)

manual["Currency"] = (
    manual_raw["Currency"]
    .astype(str)
    .apply(normalize_currency)
)

manual["Source"] = "Manual Accruals"

# ==========================================================
# SAP ERP TRANSFORMATION
# ==========================================================

sap = sap_raw.copy()

# ----------------------------------------------------------
# CARRIER
# ----------------------------------------------------------

carrier_master["CarrierKey"] = (
    carrier_master["ServcAgent"]
    .astype(str)
    .str.strip()
)

sap["CarrierKey"] = (
    sap["ServcAgent"]
    .astype(str)
    .str.strip()
)

sap = sap.merge(
    carrier_master[["CarrierKey", "Name 1"]],
    on="CarrierKey",
    how="left"
)

sap["Carrier"] = sap["Name 1"]

# ----------------------------------------------------------
# TYPE OF GOODS
# ----------------------------------------------------------

def classify_goods(x):

    if pd.isna(x):
        return "NFG"

    value = str(x).strip()

    if value.startswith("1") or value.startswith("4"):
        return "FG"

    return "NFG"

sap["Type of Goods"] = (
    sap["Product Hierarchy"]
    .apply(classify_goods)
)

# ----------------------------------------------------------
# DATE
# ----------------------------------------------------------

sap["Date"] = pd.to_datetime(
    sap["Deliv.Date"],
    errors="coerce"
)

# ----------------------------------------------------------
# ORIGIN
# ----------------------------------------------------------

shipping_master["ShPtKey"] = (
    shipping_master["ShPt"]
    .astype(str)
    .str.strip()
)

sap["ShPtKey"] = (
    sap["ShPt"]
    .astype(str)
    .str.strip()
)

sap = sap.merge(
    shipping_master[["ShPtKey", "Description"]],
    on="ShPtKey",
    how="left"
)

sap["Origin"] = np.where(
    sap["ShPt"].isna(),
    "Import",
    sap["Description"]
)

# ----------------------------------------------------------
# DESTINATION
# ----------------------------------------------------------

customer_master["ShipToKey"] = (
    customer_master["Ship-To"]
    .astype(str)
    .str.strip()
)

sap["ShipToKey"] = (
    sap["Ship-To"]
    .astype(str)
    .str.strip()
)

sap = sap.merge(
    customer_master[["ShipToKey", "Name 1"]],
    on="ShipToKey",
    how="left",
    suffixes=("", "_Customer")
)

plant_master["PlantKey"] = (
    plant_master["Plnt"]
    .astype(str)
    .str.strip()
)

sap["PlantKey"] = (
    sap["Plnt"]
    .astype(str)
    .str.strip()
)

sap = sap.merge(
    plant_master[["PlantKey", "Name 1"]],
    on="PlantKey",
    how="left",
    suffixes=("", "_Plant")
)

sap["Destination"] = np.where(
    sap["Ship-To"].isna(),
    sap["Name 1_Plant"],
    sap["Name 1"]
)

# ----------------------------------------------------------
# VALUE / CURRENCY
# ----------------------------------------------------------

sap["Value"] = pd.to_numeric(
    sap["Loc.curr.amount"],
    errors="coerce"
)

sap["Currency"] = (
    sap["Local Curr."]
    .astype(str)
    .apply(normalize_currency)
)

sap["Source"] = "SAP ERP"

sap_final = sap[
    [
        "Carrier",
        "Type of Goods",
        "Date",
        "Origin",
        "Destination",
        "Value",
        "Currency",
        "Source",
    ]
]

# ==========================================================
# EUR CONVERSION
# ==========================================================

currency_master.columns = [
    c.strip()
    for c in currency_master.columns
]

currency_master["Currency"] = (
    currency_master["Currency"]
    .astype(str)
    .str.upper()
)

combined = pd.concat(
    [manual, sap_final],
    ignore_index=True
)

combined = combined.merge(
    currency_master,
    on="Currency",
    how="left"
)

rate_column = None

for c in currency_master.columns:
    if "rate" in c.lower():
        rate_column = c
        break

if rate_column is None:
    st.error(
        "Currency conversion file must contain a column with 'Rate' in the name."
    )
    st.stop()

combined["EUR Rate"] = combined[rate_column]

combined["Value EUR"] = np.where(
    combined["Currency"] == "EUR",
    combined["Value"],
    combined["Value"] * combined["EUR Rate"]
)

combined["Year"] = combined["Date"].dt.year
combined["Month"] = combined["Date"].dt.strftime("%Y-%m")

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================

st.sidebar.header("Filters")

carrier_filter = st.sidebar.multiselect(
    "Carrier",
    sorted(combined["Carrier"].dropna().unique())
)

goods_filter = st.sidebar.multiselect(
    "Type of Goods",
    sorted(combined["Type of Goods"].dropna().unique())
)

origin_filter = st.sidebar.multiselect(
    "Origin",
    sorted(combined["Origin"].dropna().unique())
)

destination_filter = st.sidebar.multiselect(
    "Destination",
    sorted(combined["Destination"].dropna().unique())
)

source_filter = st.sidebar.multiselect(
    "Source",
    sorted(combined["Source"].dropna().unique())
)

filtered = combined.copy()

if carrier_filter:
    filtered = filtered[
        filtered["Carrier"].isin(carrier_filter)
    ]

if goods_filter:
    filtered = filtered[
        filtered["Type of Goods"].isin(goods_filter)
    ]

if origin_filter:
    filtered = filtered[
        filtered["Origin"].isin(origin_filter)
    ]

if destination_filter:
    filtered = filtered[
        filtered["Destination"].isin(destination_filter)
    ]

if source_filter:
    filtered = filtered[
        filtered["Source"].isin(source_filter)
    ]

# ==========================================================
# KPIs
# ==========================================================

k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "Spend EUR",
    f"€ {filtered['Value EUR'].sum():,.0f}"
)

k2.metric(
    "Shipments",
    f"{len(filtered):,}"
)

k3.metric(
    "Carriers",
    filtered["Carrier"].nunique()
)

k4.metric(
    "Origins",
    filtered["Origin"].nunique()
)

# ==========================================================
# CHARTS
# ==========================================================

st.subheader("Monthly Spend")

monthly = (
    filtered
    .groupby("Month", as_index=False)["Value EUR"]
    .sum()
)

fig = px.line(
    monthly,
    x="Month",
    y="Value EUR",
    markers=True
)

st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------

col1, col2 = st.columns(2)

with col1:

    spend_by_carrier = (
        filtered
        .groupby("Carrier", as_index=False)["Value EUR"]
        .sum()
        .sort_values("Value EUR", ascending=False)
        .head(15)
    )

    fig = px.bar(
        spend_by_carrier,
        x="Carrier",
        y="Value EUR",
        title="Top Carriers"
    )

    st.plotly_chart(fig, use_container_width=True)

with col2:

    spend_by_goods = (
        filtered
        .groupby("Type of Goods", as_index=False)["Value EUR"]
        .sum()
    )

    fig = px.pie(
        spend_by_goods,
        names="Type of Goods",
        values="Value EUR",
        title="FG vs NFG"
    )

    st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------

route_spend = (
    filtered
    .groupby(
        ["Origin", "Destination"],
        as_index=False
    )["Value EUR"]
    .sum()
    .sort_values("Value EUR", ascending=False)
    .head(20)
)

fig = px.bar(
    route_spend,
    x="Origin",
    y="Value EUR",
    color="Destination",
    title="Top Freight Lanes"
)

st.plotly_chart(fig, use_container_width=True)

# ==========================================================
# DETAIL TABLE
# ==========================================================

st.subheader("Detailed Freight Cost Data")

st.dataframe(
    filtered.sort_values(
        "Date",
        ascending=False
    ),
    use_container_width=True,
    height=500
)

# ==========================================================
# EXPORT
# ==========================================================

csv = filtered.to_csv(index=False).encode("utf-8")

st.download_button(
    "Download Filtered Data",
    csv,
    "freight_spend.csv",
    "text/csv"
)

README.md
Freight Cost Analytics Dashboard
Overview

This application consolidates freight costs from:

Manual Accruals
SAP ERP

and produces a unified freight spend dashboard with all costs converted to EUR.

The application allows logistics, transportation, finance, and supply chain teams to:

Analyze actual freight spend
Compare carriers
Compare FG vs NFG costs
Analyze shipping lanes
Filter costs dynamically
Export filtered datasets
Business Logic

The application standardizes both data sources into the following model:

Field	DescriptionCarrier	Carrier name
Type of Goods	FG / NFG
Date	Shipment date
Origin	Shipping origin
Destination	Delivery location
Value	Original freight amount
Currency	Original currency
Value EUR	Converted amount in EUR
Manual Accruals Mapping
Dashboard Field	Manual Accruals ColumnCarrier	Carrier Description
Type of Goods	PBU
Date	Planned Arrival Date-Last Stop
Origin	Source Location Description
Destination	Destination Location Descripti
Value	Net Amt in Doc Crcy
Currency	Currency
Carrier Cleaning Rule

When Carrier Description contains:

DHL/ABC


the dashboard keeps:

DHL


Everything after "/" is removed.

SAP ERP Mapping
Dashboard Field	SAP ERP ColumnCarrier	ServcAgent
Type of Goods	Product Hierarchy
Date	Deliv.Date
Origin	ShPt
Destination	Ship-To
Value	Loc.curr.amount
Currency	Local Curr.
Carrier Enrichment

SAP ERP ServcAgent is linked to:

ERP Carrier Name


Relationship:

SAP ERP.ServcAgent
=
ERP Carrier Name.ServcAgent


Result:

Carrier = ERP Carrier Name.Name 1

Goods Classification

Product Hierarchy logic:

Starts with 1 -> FG
Starts with 4 -> FG
Anything else -> NFG


Examples:

1000000 → FG
4000000 → FG
7000000 → NFG
8000000 → NFG

Origin Enrichment

Relationship:

SAP ERP.ShPt
=
ERP Shipping Point.ShPt


Result:

Origin = ERP Shipping Point.Description


Special case:

ShPt empty


Result:

Origin = Import

Destination Enrichment
Standard Flow

Relationship:

SAP ERP.Ship-To
=
ERP Customers.Ship-To


Result:

Destination = ERP Customers.Name 1

Fallback Logic

When Ship-To is empty:

SAP ERP.Plnt
=
ERP Plants.Plnt


Result:

Destination = ERP Plants.Name 1

Currency Conversion

All costs are converted to EUR.

Relationship:

Currency
=
Currency conversion table.Currency


Formula:

Value EUR = Value × Conversion Rate


Exception:

Currency = EUR


Formula:

Value EUR = Value

Folder Structure
project/
│
├── app.py
├── requirements.txt
│
├── master_data/
│   ├── Currency conversion table.xlsx
│   ├── ERP Carrier Name.xlsx
│   ├── ERP Customers.xlsx
│   ├── ERP Plants.xlsx
│   └── ERP Shipping Point.xlsx
│
└── uploads/

Installation

Create virtual environment:

python -m venv venv


Activate:

Windows:

venv\Scripts\activate


Linux / Mac:

source venv/bin/activate


Install dependencies:

pip install -r requirements.txt

Running the Application
streamlit run app.py


The dashboard opens automatically in the browser.

Default URL:

http://localhost:8501

Dashboard Features
KPI Cards
Total Spend EUR
Shipment Count
Number of Carriers
Number of Origins
Interactive Filters
Carrier
Type of Goods
Origin
Destination
Source System
Visualizations
Monthly Spend Trend

Tracks freight spend evolution over time.

Top Carriers

Shows highest spending carriers.

FG vs NFG Analysis

Distribution of freight spend by product category.

Top Freight Lanes

Origin → Destination spend analysis.

Detailed Dataset

Complete filtered transaction list available on screen.

Export Functionality

Users can export filtered data as CSV.

Recommended Enhancements

Future versions may include:

Budget vs Actual comparison
Carrier performance KPIs
Cost per KG
Cost per Shipment
Freight accrual reconciliation
Power BI embedding
User authentication
Multi-currency reporting
Year-over-year trend analysis
Forecasting and spend prediction

This solution provides a complete freight cost analytics application ready for deployment with Streamlit.
